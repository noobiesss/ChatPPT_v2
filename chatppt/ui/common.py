"""Tiny UI helpers shared by several UI modules.

Edit when: adding a small helper used by 2+ UI modules.
"""




def _short(text: str, n: int = 16) -> str:
    text = (text or "").split("/")[-1]
    return text if len(text) <= n else text[: n - 1] + "…"
