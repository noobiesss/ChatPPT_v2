"""Text and embedded-image extraction from uploaded PDF / DOCX / TXT / MD files.

Edit when: changing how uploaded reports are read.
"""

import hashlib
import io
from typing import List, Dict, Optional

from docx import Document
from PIL import Image
from pypdf import PdfReader

# Embedded pictures smaller than this (logos, icons, bullets) are not useful as slide figures
REPORT_IMAGE_MIN_SIDE = 200
REPORT_IMAGE_MIN_AREA = 200 * 260
# Very large scans are downscaled so session state stays small; still sharp on a 13.3" slide
REPORT_IMAGE_MAX_SIDE = 2400
REPORT_IMAGE_MAX_COUNT = 30


def prepare_report_image(raw: bytes) -> Optional[Dict[str, object]]:
    """
    Validate one embedded picture: decode it, drop tiny ones, convert exotic formats/modes to PNG or JPEG
    (browsers and python-pptx both accept those), downscale huge ones. Returns None for unusable data.
    """
    try:
        with Image.open(io.BytesIO(raw)) as img:
            img.load()
            w, h = img.size
            if min(w, h) < REPORT_IMAGE_MIN_SIDE or w * h < REPORT_IMAGE_MIN_AREA:
                return None
            if img.format in ("JPEG", "PNG") and max(w, h) <= REPORT_IMAGE_MAX_SIDE:
                return {"bytes": raw, "width": w, "height": h}
            has_alpha = img.mode in ("RGBA", "LA", "P")
            out_img = img.convert("RGBA" if has_alpha else "RGB")
            if max(w, h) > REPORT_IMAGE_MAX_SIDE:
                out_img.thumbnail((REPORT_IMAGE_MAX_SIDE, REPORT_IMAGE_MAX_SIDE))
            buf = io.BytesIO()
            if has_alpha:
                out_img.save(buf, format="PNG")
            else:
                out_img.save(buf, format="JPEG", quality=90)
            return {"bytes": buf.getvalue(), "width": out_img.size[0], "height": out_img.size[1]}
    except Exception:  # noqa: BLE001 - JBIG2/CCITT and other undecodable streams are skipped
        return None


def extract_text_from_uploaded_file(uploaded_file) -> str:
    """Extract plain text from uploaded PDF, DOCX, TXT, or MD files."""
    if uploaded_file is None:
        return ""

    filename = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    if filename.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(file_bytes))
        return "\n".join([page.extract_text() or "" for page in reader.pages])

    elif filename.endswith(".docx"):
        doc = Document(io.BytesIO(file_bytes))
        return "\n".join([para.text for para in doc.paragraphs if para.text.strip()])

    elif filename.endswith((".txt", ".md")):
        return file_bytes.decode("utf-8", errors="ignore")

    return ""


def extract_images_from_report(uploaded_file) -> List[Dict[str, bytes]]:
    """
    Extract embedded images from an uploaded PDF or DOCX report so users can reuse
    original diagrams, charts, or photos directly inside their slides.
    Each item: {"name", "bytes", "width", "height"}. Tiny pictures (logos, icons) and exact duplicates
    (a logo repeated on every page) are dropped; at most REPORT_IMAGE_MAX_COUNT are kept, in document order.
    """
    extracted: List[Dict[str, bytes]] = []
    if uploaded_file is None:
        return extracted
    seen = set()

    def add(name: str, raw: bytes) -> None:
        digest = hashlib.sha1(raw).hexdigest()
        if digest in seen or len(extracted) >= REPORT_IMAGE_MAX_COUNT:
            return
        seen.add(digest)
        info = prepare_report_image(raw)
        if info:
            extracted.append({"name": name, **info})

    filename = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    try:
        if filename.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(file_bytes))
            for page_no, page in enumerate(reader.pages, start=1):
                try:
                    page_images = list(page.images)
                except Exception:  # noqa: BLE001 - one broken page must not lose the others
                    continue
                for image_file_object in page_images:
                    add(f"Page {page_no} · {image_file_object.name}", image_file_object.data)

        elif filename.endswith(".docx"):
            doc = Document(io.BytesIO(file_bytes))
            for rel in doc.part.rels.values():
                if "image" in rel.reltype and not rel.is_external:
                    add(rel.target_ref.rsplit("/", 1)[-1], rel.target_part.blob)
    except Exception:
        pass

    return extracted
