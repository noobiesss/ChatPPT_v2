"""Pydantic schemas and enums for the structured slide outline.

Edit when: adding a layout type, a new slide field, or a new chart/image option.
Used by: outline.py, pptx_*.py, ui/*.
"""

from enum import Enum
from typing import List, Optional

from pydantic import BaseModel, Field


class LayoutType(str, Enum):
    COVER = "cover"
    SECTION_DIVIDER = "section_divider"
    BULLET_CARDS = "bullet_cards"
    THREE_CARDS = "three_cards"
    TWO_COLUMN = "two_column"
    KPI_METRICS = "kpi_metrics"
    CHART_SLIDE = "chart_slide"
    IMAGE_SPLIT = "image_split"
    CLOSING = "closing"


class ChartType(str, Enum):
    BAR = "bar"
    LINE = "line"
    PIE = "pie"


class ImagePlacement(str, Enum):
    LEFT = "left"
    RIGHT = "right"


class VisualStyle(str, Enum):
    STOCK_PHOTO = "stock_photo"
    AI_ILLUSTRATION = "ai_illustration"


# Two-option image mode (Option 1 is the universal default when the user makes no selection)
IMAGE_MODE_UNSPLASH_FIRST = "Unsplash Photo First (Fallback to AI)"


IMAGE_MODE_PLACEHOLDER = "Reserve Photo Placeholder"


IMAGE_MODE_OPTIONS = [IMAGE_MODE_UNSPLASH_FIRST, IMAGE_MODE_PLACEHOLDER]


class StructuredPoint(BaseModel):
    heading: str = Field(
        description="Short, bold anchor phrase (3 to 6 words) summarizing the point, e.g., '3x Faster Query Speed'."
    )
    detail: str = Field(
        description="Concise explanation (12 to 25 words max). Avoid long paragraphs."
    )


class ColumnContent(BaseModel):
    column_title: str = Field(
        description="Header for this comparison column, e.g., 'Legacy System' or 'Proposed Architecture'."
    )
    items: List[StructuredPoint] = Field(
        description="2 to 4 structured points belonging to this column."
    )


class KPIMetric(BaseModel):
    value: str = Field(
        description="High-impact short metric string (max 8 chars), e.g., '+62%', '$2.4M', '99.9%', '3.5x'."
    )
    label: str = Field(
        description="Short metric title (2 to 4 words), e.g., 'Latency Reduction'."
    )
    context: str = Field(
        description="Brief 5 to 12 word context explaining the metric."
    )


class ChartSpec(BaseModel):
    chart_type: ChartType = Field(
        description="Select 'bar' for category comparisons, 'line' for time trends, or 'pie' for percentage shares."
    )
    chart_title: str = Field(
        description="Descriptive title for the chart graphic itself."
    )
    categories: List[str] = Field(
        description="Labels for each data point on the X-axis or pie slices (2 to 6 items)."
    )
    series_name: str = Field(
        description="Name of the data series including unit, e.g., 'Revenue ($M)' or 'Response Time (ms)'."
    )
    values: List[float] = Field(
        description="Numerical values corresponding 1-to-1 with 'categories'. Length MUST match 'categories'."
    )


class ImageSpec(BaseModel):
    preferred_style: VisualStyle = Field(
        default=VisualStyle.STOCK_PHOTO,
        description="Default to 'stock_photo' so Unsplash photography is prioritized before falling back to AI."
    )
    visual_prompt: str = Field(
        description="Detailed English prompt (20 to 35 words) for an AI image generator fallback describing a clean, "
                    "modern illustration or photo that visually explains this slide's concept. "
                    "Specify 'clean minimal corporate style, no text inside image'."
    )
    search_keywords: str = Field(
        description="2 to 3 concise, concrete English keywords for Unsplash stock photo search, "
                    "e.g., 'server room', 'team strategy meeting', or 'renewable energy'."
    )
    caption: str = Field(
        description="Short caption (8 to 15 words, in the target presentation language) displayed below the image."
    )
    placement: ImagePlacement = Field(
        default=ImagePlacement.LEFT,
        description="Alternate between 'left' and 'right' across different image slides."
    )


class SlideSpec(BaseModel):
    slide_number: int = Field(description="Sequential slide number starting from 1.")
    layout_type: LayoutType = Field(description="Visual layout archetype selected for this slide.")
    action_title: str = Field(
        description="Conclusion-first Action Title (6 to 14 words) stating the core takeaway, NOT a vague label."
    )
    subtitle_or_context: Optional[str] = Field(
        default=None,
        description="Used for 'cover' (presentation subtitle), 'section_divider' (section overview), or 'closing' "
                    "(optional one-line presenter note such as 'Presented by ...'; leave empty if the user gave no name)."
    )
    points: Optional[List[StructuredPoint]] = Field(
        default=None,
        description="Populate for 'bullet_cards' (3-4 items), 'three_cards' (3 items), "
                    "'chart_slide' (2-3 insights), or 'image_split' (2-3 key takeaways)."
    )
    left_column: Optional[ColumnContent] = Field(
        default=None,
        description="Populate ONLY when layout_type is 'two_column'."
    )
    right_column: Optional[ColumnContent] = Field(
        default=None,
        description="Populate ONLY when layout_type is 'two_column'."
    )
    kpis: Optional[List[KPIMetric]] = Field(
        default=None,
        description="Populate ONLY when layout_type is 'kpi_metrics' (2 to 4 headline numbers)."
    )
    chart: Optional[ChartSpec] = Field(
        default=None,
        description="Populate ONLY when layout_type is 'chart_slide'."
    )
    image: Optional[ImageSpec] = Field(
        default=None,
        description="Populate ONLY when layout_type is 'image_split'."
    )
    bottom_takeaway: Optional[str] = Field(
        default=None,
        description="Single punchy sentence (10 to 20 words) summarizing the 'So What?' in a bottom banner. "
                    "Leave None for 'cover', 'section_divider' and 'closing'."
    )


class PresentationOutline(BaseModel):
    deck_title: str = Field(description="Overall title of the slide deck.")
    slides: List[SlideSpec] = Field(description="Ordered list of slides matching the requested slide count.")