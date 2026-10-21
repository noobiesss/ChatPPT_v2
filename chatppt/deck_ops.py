"""Deck-level operations on a chat's outline that need no LLM: undo / redo history and
structure edits (move / insert / delete slides).

Edit when: changing how outline versions are kept or restored, or the rules for moving / adding / removing slides.
Pure functions (no Streamlit), so they are covered by smoke_test.py.
"""

import copy
from typing import List, Optional, Tuple

from chatppt.outline import MAX_SLIDES, _sanitize_slide
from chatppt.schemas import LayoutType, SlideSpec

# Outline snapshots kept per chat; older ones are dropped first
HISTORY_MAX = 20


def _stacks(chat: dict):
    # Chats created before history existed get their stacks on first use
    return chat.setdefault("slides_history", []), chat.setdefault("slides_future", [])


def push_history(chat: dict) -> None:
    """Snapshot the current outline before it changes. A new change clears the redo stack."""
    if not chat.get("slides"):
        return
    history, future = _stacks(chat)
    history.append(copy.deepcopy(chat["slides"]))
    del history[:-HISTORY_MAX]
    future.clear()


def can_undo(chat: dict) -> bool:
    return bool(chat.get("slides_history"))


def can_redo(chat: dict) -> bool:
    return bool(chat.get("slides_future"))


def _swap(chat: dict, source: List, target: List) -> Optional[List[dict]]:
    # Shared by undo and redo: save the current outline on one stack, restore the top of the other
    if not source:
        return None
    if chat.get("slides"):
        target.append(copy.deepcopy(chat["slides"]))
        del target[:-HISTORY_MAX]
    chat["slides"] = source.pop()
    return chat["slides"]


def undo(chat: dict) -> Optional[List[dict]]:
    """Restore the previous outline version (the current one becomes redoable). None if nothing to undo."""
    history, future = _stacks(chat)
    return _swap(chat, history, future)


def redo(chat: dict) -> Optional[List[dict]]:
    """Re-apply the version that was undone last. None if nothing to redo."""
    history, future = _stacks(chat)
    return _swap(chat, future, history)


# --- Structure edits. Rules: the cover stays first, a closing slide stays last ---
MIN_DECK_SLIDES = 2


def _is_closing(slide: dict) -> bool:
    return slide.get("layout_type") == LayoutType.CLOSING.value


def _renumbered(slides: List[dict]) -> List[dict]:
    for i, s in enumerate(slides):
        s["slide_number"] = i + 1
    return slides


def can_move(slides: List[dict], idx: int, delta: int) -> bool:
    j = idx + delta
    if not (0 <= idx < len(slides) and 0 <= j < len(slides)) or 0 in (idx, j):
        return False
    return not (_is_closing(slides[idx]) or _is_closing(slides[j]))


def move_slide(slides: List[dict], idx: int, delta: int) -> List[dict]:
    """Swap slide `idx` with its neighbor `idx + delta` (-1 up, +1 down). Unchanged copy if not allowed."""
    out = list(slides)
    if can_move(out, idx, delta):
        out[idx], out[idx + delta] = out[idx + delta], out[idx]
    return _renumbered(out)


def can_delete(slides: List[dict], idx: int) -> bool:
    return 0 < idx < len(slides) and len(slides) > MIN_DECK_SLIDES


def delete_slide(slides: List[dict], idx: int) -> List[dict]:
    out = list(slides)
    if can_delete(out, idx):
        del out[idx]
    return _renumbered(out)


def can_insert(slides: List[dict]) -> bool:
    return len(slides) < MAX_SLIDES


def new_blank_slide(position: int, image_mode: str) -> dict:
    """An editable bullet-cards slide with placeholder text, sanitized like a generated one."""
    spec = SlideSpec(
        slide_number=position + 1,
        layout_type=LayoutType.BULLET_CARDS,
        action_title="New slide: write a conclusion-first title",
        points=[{"heading": f"Point {i}", "detail": "Replace with a short supporting detail."} for i in (1, 2, 3)],
        bottom_takeaway="Replace with the so-what of this slide.",
    )
    return _sanitize_slide(spec, position, image_mode)


def insert_slide(slides: List[dict], after_idx: int, image_mode: str) -> Tuple[List[dict], int]:
    """Insert a blank slide after `after_idx` (never after a closing slide). Returns (slides, new index)."""
    out = list(slides)
    if not can_insert(out):
        return _renumbered(out), after_idx
    pos = max(1, after_idx + 1)
    if out and _is_closing(out[-1]):
        pos = min(pos, len(out) - 1)
    out.insert(pos, new_blank_slide(pos, image_mode))
    return _renumbered(out), pos
