"""Text <-> structured data helpers for the outline editor text areas.

Edit when: changing the 'Heading :: Detail' editing format.
"""

from typing import List


def format_points_for_editor(points: List[dict]) -> str:
    """Serialize structured points as 'Heading :: Detail' lines for easy text-area editing."""
    lines = []
    for pt in points or []:
        h = (pt.get("heading") or "").strip()
        d = (pt.get("detail") or "").strip()
        lines.append(f"{h} :: {d}" if d else h)
    return "\n".join(lines)


def parse_points_from_editor(raw_text: str) -> List[dict]:
    """Parse 'Heading :: Detail' lines back into structured point dictionaries."""
    parsed = []
    for line in raw_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        if "::" in line:
            parts = line.split("::", 1)
            parsed.append({"heading": parts[0].strip(), "detail": parts[1].strip()})
        else:
            parsed.append({"heading": line[:40], "detail": line})
    return parsed


def format_kpis_for_editor(kpis: List[dict]) -> str:
    """Serialize KPIs as 'Value | Label | Context' lines."""
    return "\n".join(
        f"{k.get('value', '')} | {k.get('label', '')} | {k.get('context', '')}"
        for k in (kpis or [])
    )


def parse_kpis_from_editor(raw_text: str) -> List[dict]:
    """Parse 'Value | Label | Context' lines back into KPI dictionaries."""
    parsed = []
    for line in raw_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 3:
            parsed.append({"value": parts[0], "label": parts[1], "context": parts[2]})
        elif len(parts) == 2:
            parsed.append({"value": parts[0], "label": parts[1], "context": ""})
        else:
            parsed.append({"value": parts[0][:8], "label": "Metric", "context": parts[0]})
    return parsed
