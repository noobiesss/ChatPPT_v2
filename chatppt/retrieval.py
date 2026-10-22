"""Lightweight chunk retrieval over the source document (BM25, no vector store, no extra dependencies).

Edit when: changing how long sources are cut into chunks or how the relevant chunks are picked for a question.
"""

import math  # log() for the BM25 idf term
import re  # tokenizer patterns
from collections import Counter  # term frequencies per chunk and document frequencies
from functools import lru_cache  # caches the index of the last few source texts
from typing import List, Tuple

# A chunk is a run of whole lines of roughly this size; 12k chars of context ~= 12 chunks
CHUNK_TARGET_CHARS = 1000
# BM25 tuning: k1 = how fast repeated terms stop adding score, b = how strongly long chunks are penalized
BM25_K1, BM25_B = 1.5, 0.75
# Marker placed between non-adjacent excerpts so the model knows text was skipped
GAP_MARKER = "\n\n[...]\n\n"

# Words and numbers, keeping figures like "3.2%" or "1,200" as one token
# [a-z0-9]+ a word or number, (?:[.,][0-9]+)* decimal / thousands parts, %? an optional percent sign
_WORD_RE = re.compile(r"[a-z0-9]+(?:[.,][0-9]+)*%?")
# Runs of CJK ideographs (Extension A, Unified Ideographs, Compatibility Ideographs)
_CJK_RUN_RE = re.compile(r"[㐀-䶿一-鿿豈-﫿]+")
# Words too common to tell chunks apart; "slide", "deck" etc. are added because every question mentions them
_STOPWORDS = frozenset(
    "a an and are as at be but by can do does for from has have how i in is it its me my of on or our so that the "
    "their them there these they this to was we were what when where which who why will with you your about "
    "slide slides page pages deck please could would should".split()
)


def tokenize(text: str) -> List[str]:
    """Lowercase word tokens (stopwords dropped) plus CJK character bigrams, so Chinese text needs no segmenter."""
    text = (text or "").lower()  # case-insensitive matching
    # Keep words that are not stopwords and are longer than 1 character (single digits are kept: "3 options")
    tokens = [w for w in _WORD_RE.findall(text) if w not in _STOPWORDS and (len(w) > 1 or w.isdigit())]
    for run in _CJK_RUN_RE.findall(text):  # each unbroken run of Chinese / Japanese characters
        if len(run) == 1:
            tokens.append(run)  # a lone character is its own token
        else:
            # Overlapping 2-character pieces: "营收增长" -> "营收", "收增", "增长".
            # Most Chinese words are 2 characters, so bigrams match words without a word segmenter
            tokens.extend(run[i:i + 2] for i in range(len(run) - 1))
    return tokens


def split_into_chunks(text: str, target_chars: int = CHUNK_TARGET_CHARS) -> List[str]:
    """
    Cut text into chunks of whole lines (~target_chars each). A blank line closes a chunk once it is
    half full, so paragraph breaks are preferred; very long lines (PDF text without newlines) are hard-split.
    """
    chunks: List[str] = []  # finished chunks
    buf: List[str] = []  # lines of the chunk being built
    size = 0  # characters in buf (including one newline per line)

    def flush() -> None:
        # Close the current chunk: join its lines, keep it if it has any text, start an empty one
        nonlocal buf, size
        joined = "\n".join(buf).strip()
        if joined:
            chunks.append(joined)
        buf, size = [], 0

    for line in (text or "").splitlines():
        if not line.strip():  # a blank line = paragraph break
            if size >= target_chars // 2:  # end the chunk here only if it is at least half full
                flush()
            continue  # blank lines themselves are never stored
        # Cut lines longer than one chunk into target_chars-sized pieces (most lines give a single piece)
        pieces = [line[i:i + target_chars] for i in range(0, len(line), target_chars)]
        for piece in pieces:
            if size and size + len(piece) > target_chars:  # this piece would overfill a non-empty chunk
                flush()
            buf.append(piece)
            size += len(piece) + 1  # +1 for the newline that will join it
    flush()  # the last, partly filled chunk
    return chunks


@lru_cache(maxsize=4)
def _index(text: str) -> Tuple[Tuple[str, ...], Tuple[Counter, ...], Tuple[int, ...], Counter, float]:
    """Chunk + tokenize once per source text (the same source is queried on every chat turn)."""
    chunks = tuple(split_into_chunks(text))
    term_freqs = tuple(Counter(tokenize(c)) for c in chunks)  # per chunk: how often each token occurs
    lengths = tuple(sum(tf.values()) for tf in term_freqs)  # per chunk: number of tokens
    doc_freq: Counter = Counter()  # per token: in how many chunks it occurs
    for tf in term_freqs:
        doc_freq.update(tf.keys())  # keys() so each chunk counts a token once
    avg_len = (sum(lengths) / len(lengths)) if lengths else 0.0  # average chunk length, used by BM25's b term
    # Tuples (not lists) so the cached result cannot be modified by a caller
    return chunks, term_freqs, lengths, doc_freq, avg_len


def bm25_scores(text: str, query: str) -> List[float]:
    """BM25 score of every chunk of `text` for `query` (same order as split_into_chunks)."""
    chunks, term_freqs, lengths, doc_freq, avg_len = _index(text)
    n = len(chunks)
    terms = set(tokenize(query))  # a set: a word repeated in the question counts once
    scores = [0.0] * n
    if not n or not terms:
        return scores  # empty document or empty question: every chunk scores 0
    # Standard BM25: rare terms weigh more (idf), repeated terms saturate (k1), long chunks are penalized (b)
    for term in terms:
        df = doc_freq.get(term, 0)  # number of chunks containing the term
        if not df:
            continue  # the term appears nowhere: it cannot raise any score
        # Inverse document frequency: high for terms found in few chunks; the "1 +" keeps it positive
        idf = math.log(1 + (n - df + 0.5) / (df + 0.5))
        for i, tf in enumerate(term_freqs):
            f = tf.get(term, 0)  # occurrences of the term in chunk i
            if f:
                # Length normalization: chunks longer than average need more occurrences for the same score
                norm = BM25_K1 * (1 - BM25_B + BM25_B * lengths[i] / (avg_len or 1))
                # Term-frequency part: grows with f but levels off (saturates) as f gets large
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
        return text  # fits in the budget: nothing to select
    chunks = _index(text)[0]  # cached, so bm25_scores below reuses the same index
    scores = bm25_scores(text, query)
    if not any(scores):
        # No chunk shares a term with the query: send the start of the document, marked as cut
        return text[:max_chars] + GAP_MARKER

    chosen = {0}  # the first chunk (title / intro) is always kept
    used = len(chunks[0])  # characters spent so far

    def take(i: int) -> None:
        # Add chunk i if it exists, is not chosen yet and still fits in the budget (with its gap marker)
        nonlocal used
        if not 0 <= i < len(chunks) or i in chosen:
            return
        cost = len(chunks[i]) + len(GAP_MARKER)  # counted with a marker even when it ends up adjacent
        if used + cost <= max_chars:
            chosen.add(i)
            used += cost

    # 1) best matches; 2) their neighbors (a hit often continues in the next chunk);
    # 3) leftover budget goes to the head of the document, so broad questions never get less than before
    # Chunks with a positive score, best first (chunk 0 is already chosen, so the range starts at 1)
    hits = [i for i in sorted(range(1, len(chunks)), key=lambda i: scores[i], reverse=True) if scores[i] > 0]
    for i in hits:
        take(i)
    for i in [h for h in hits if h in chosen]:  # only hits that fit get their neighbors
        take(i + 1)
        take(i - 1)
    for i in range(1, len(chunks)):  # fill what is left with chunks in document order
        take(i)

    parts: List[str] = []
    prev = -1  # index of the previously emitted chunk
    for i in sorted(chosen):  # document order, not score order, so the excerpts read naturally
        if parts:
            # Consecutive chunks are joined by a plain newline; a jump gets the [...] marker
            parts.append(GAP_MARKER if i != prev + 1 else "\n")
        parts.append(chunks[i])
        prev = i
    if prev != len(chunks) - 1:
        parts.append(GAP_MARKER)  # the end of the document was skipped too
    return "".join(parts)
