"""Text <-> structured data helpers for the outline editor text areas.

Edit when: changing the 'Heading :: Detail', KPI, timeline or table editing formats.
"""

from typing import List, Tuple


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


def format_steps_for_editor(steps: List[dict]) -> str:
    """Serialize timeline steps as 'Stage | Title | Description' lines."""
    return "\n".join(
        f"{s.get('stage_label', '')} | {s.get('title', '')} | {s.get('description', '')}" for s in (steps or [])
    )


def parse_steps_from_editor(raw_text: str) -> List[dict]:
    """Parse 'Stage | Title | Description' lines; a line without '|' becomes a step title."""
    parsed = []
    for line in raw_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split("|", 2)]
        if len(parts) == 3:
            stage, title, desc = parts
        elif len(parts) == 2:
            stage, title, desc = parts[0], parts[1], ""
        else:
            stage, title, desc = "", parts[0], ""
        parsed.append({"stage_label": stage or f"Step {len(parsed) + 1}", "title": title, "description": desc})
    return parsed


def format_table_for_editor(headers: List[str], rows: List[List[str]]) -> str:
    """Serialize a table as lines of ' | '-separated cells; the first line holds the headers."""
    return "\n".join(" | ".join(r) for r in [headers or []] + list(rows or []) if r)


def parse_table_from_editor(raw_text: str) -> Tuple[List[str], List[List[str]]]:
    """Parse ' | '-separated lines (first line = headers). Rows are padded / cut to the header count."""
    lines = [ln.strip() for ln in raw_text.split("\n") if ln.strip()]
    if not lines:
        return [], []
    headers = [c.strip() for c in lines[0].split("|")]
    rows = []
    for ln in lines[1:]:
        cells = [c.strip() for c in ln.split("|")][:len(headers)]
        rows.append(cells + [""] * (len(headers) - len(cells)))
    return headers, rows
