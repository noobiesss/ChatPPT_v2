"""Tiny UI helpers and option values shared by several UI modules.

Edit when: adding a small helper or option constant used by 2+ UI modules.
"""

from pathlib import Path

from chatppt.schemas import IMAGE_MODE_PLACEHOLDER, IMAGE_MODE_REPORT_FIRST, IMAGE_MODE_UNSPLASH_FIRST

# Pixel-art brand logo (assets/ at the project root); absolute so any working directory works
ASSETS_DIR = Path(__file__).resolve().parents[2] / "assets"
LOGO_PATH = str(ASSETS_DIR / "logo.png")  # 520x440, sidebar logo and hero
LOGO_ICON_PATH = str(ASSETS_DIR / "logo_square.png")  # 520x520, favicon, chat avatar, collapsed sidebar

ASSISTANT_AVATAR = LOGO_ICON_PATH

# Session-state values of the template popover controls (stored values; labels are shown via format_func)
TPL_MODE_BUILTIN = "Built-in 16:9 theme"
TPL_MODE_CUSTOM = "Custom .pptx template"
TPL_COLORS_FROM_TEMPLATE = "Auto-extract colors from .pptx"
TPL_COLORS_FALLBACK = "Use fallback palette below"

# Short segmented-control labels for the image modes (the full mode strings stay the stored values)
IMAGE_MODE_LABELS = {
    IMAGE_MODE_UNSPLASH_FIRST: ":material/photo_library: Unsplash / AI",
    IMAGE_MODE_PLACEHOLDER: ":material/crop_original: Placeholder",
    IMAGE_MODE_REPORT_FIRST: ":material/description: Report figures",
}
IMAGE_MODE_HELP = (
    "Unsplash / AI: search Unsplash first, generate an AI image if no photo is found. "
    "Placeholder: reserve an empty photo box to fill in PowerPoint. "
    "Report figures: use figures extracted from the uploaded PDF/DOCX; slides without one fall back to Unsplash / AI."
)


def _short(text: str, n: int = 16) -> str:
    text = (text or "").split("/")[-1]
    return text if len(text) <= n else text[: n - 1] + "…"
