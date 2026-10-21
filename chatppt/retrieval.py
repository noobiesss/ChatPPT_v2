"""Lightweight chunk retrieval over the source document (BM25, no vector store, no extra dependencies).

Edit when: changing how long sources are cut into chunks or how the relevant chunks are picked for a question.
"""

import math
import re
from collections import Counter
from functools import lru_cache
from typing import List, Tuple

# A chunk is a run of whole lines of roughly this size; 12k chars of context ~= 12 chunks
CHUNK_TARGET_CHARS = 1000
BM25_K1, BM25_B = 1.5, 0.75
# Marker placed between non-adjacent excerpts so the model knows text was skipped
GAP_MARKER = "\n\n[...]\n\n"

# Words and numbers, keeping figures like "3.2%" or "1,200" as one token
_WORD_RE = re.compile(r"[a-z0-9]+(?:[.,][0-9]+)*%?")
# Runs of CJK ideographs (Extension A, Unified Ideographs, Compatibility Ideographs)
_CJK_RUN_RE =re.compile(r"[㐀-䶿一-鿿豈-﫿]+")
_STOPWORDS = frozenset(
    "a an and are as at be but by can do does for from has have how i in is it its me my of on or our so that the "
    "their them there these they this to was we were what when where which who why will with you your about "
    "slide slides page pages deck please could would should".split()
)


def tokenize(text: str) -> List[str]:
    """Lowercase word tokens (stopwords dropped) plus CJK character bigrams, so Chinese text needs no segmenter."""
    text = (text or "").lower()
    tokens = [w for w in _WORD_RE.findall(text) if w not in _STOPWORDS and (len(w) > 1 or w.isdigit())]
    for run in _CJK_RUN_RE.findall(text):
        if len(run) == 1:
            tokens.append(run)
        else:
            tokens.extend(run[i:i + 2] for i in range(len(run) - 1))
    return tokens


def split_into_chunks(text: str, target_chars: int = CHUNK_TARGET_CHARS) -> List[str]:
    """
    Cut text into chunks of whole lines (~target_chars each). A blank line closes a chunk once it is
    half full, so paragraph breaks are preferred; very long lines (PDF text without newlines) are hard-split.
    """
    chunks: List[str] = []
    buf: List[str] = []
    size = 0

    def flush() -> None:
        nonlocal buf, size
        joined = "\n".join(buf).strip()
        if joined:
            chunks.append(joined)
        buf, size = [], 0

    for line in (text or "").splitlines():
        if not line.strip():
            if size >= target_chars // 2:
                flush()
            continue
        pieces = [line[i:i + target_chars] for i in range(0, len(line), target_chars)]
        for piece in pieces:
            if size and size + len(piece) > target_chars:
                flush()
            buf.append(piece)
            size += len(piece) + 1
    flush()
    return chunks


@lru_cache(maxsize=4)
def _index(text: str) -> Tuple[Tuple[str, ...], Tuple[Counter, ...], Tuple[int, ...], Counter, float]:
    """Chunk + tokenize once per source text (the same source is queried on every chat turn)."""
    chunks = tuple(split_into_chunks(text))
    term_freqs = tuple(Counter(tokenize(c)) for c in chunks)
    lengths = tuple(sum(tf.values()) for tf in term_freqs)
    doc_freq: Counter = Counter()
    for tf in term_freqs:
        doc_freq.update(tf.keys())
    avg_len = (sum(lengths) / len(lengths)) if lengths else 0.0
    return chunks, term_freqs, lengths, doc_freq, avg_len


def bm25_scores(text: str, query: str) -> List[float]:
    """BM25 score of every chunk of `text` for `query` (same order as split_into_chunks)."""
    chunks, term_freqs, lengths, doc_freq, avg_len = _index(text)
    n = len(chunks)
    terms = set(tokenize(query))
    scores = [0.0] * n
    if not n or not terms:
        return scores
    # Standard BM25: rare terms weigh more (idf), repeated terms saturate (k1), long chunks are penalized (b)
    for term in terms:
        df = doc_freq.get(term, 0)
        if not df:
            continue
        idf = math.log(1 + (n - df + 0.5) / (df + 0.5))
        for i, tf in enumerate(term_freqs):
            f = tf.get(term, 0)
            if f:
                norm = BM25_K1 * (1 - BM25_B + BM25_B * lengths[i] / (avg_len or 1))
                scores[i] += idf * f * (BM25_K1 + 1) / (f + norm)
    return scores


def select_relevant_context(text: str, query: str, max_chars: int) -> str:
    """
    Return at most ~max_chars of `text` for answering `query`.
    Short texts are returned whole. Long texts keep the first chunk (title / intro) plus the best BM25 chunks,
    re-assembled in document order with GAP_MARKER between non-adjacent excerpts.
    If nothing matches the query, fall back to the head of the document (the previous behavior).
    """
    text = (text or "").strip()
    if len(text) <= max_chars:
        return text
    chunks = _index(text)[0]
    scores = bm25_scores(text, query)
    if not any(scores):
        return text[:max_chars] + GAP_MARKER

    chosen = {0}  # the first chunk (title / intro) is always kept
    used = len(chunks[0])

    def take(i: int) -> None:
        nonlocal used
        if not 0 <= i < len(chunks) or i in chosen:
            return
        cost = len(chunks[i]) + len(GAP_MARKER)
        if used + cost <= max_chars:
            chosen.add(i)
            used += cost

    # 1) best matches; 2) their neighbors (a hit often continues in the next chunk);
    # 3) leftover budget goes to the head of the document, so broad questions never get less than before
    hits = [i for i in sorted(range(1, len(chunks)), key=lambda i: scores[i], reverse=True) if scores[i] > 0]
    for i in hits:
        take(i)
    for i in [h for h in hits if h in chosen]:
        take(i + 1)
        take(i - 1)
    for i in range(1, len(chunks)):
        take(i)

    parts: List[str] = []
    prev = -1
    for i in sorted(chosen):
        if parts:
            parts.append(GAP_MARKER if i != prev + 1 else "\n")
        parts.append(chunks[i])
        prev = i
    if prev != len(chunks) - 1:
        parts.append(GAP_MARKER)
    return "".join(parts)
