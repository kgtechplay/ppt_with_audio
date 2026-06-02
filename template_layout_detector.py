from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER


TITLE_TYPES = {
    PP_PLACEHOLDER.CENTER_TITLE,
    PP_PLACEHOLDER.TITLE,
    PP_PLACEHOLDER.VERTICAL_TITLE,
}

BODY_TYPES = {
    PP_PLACEHOLDER.BODY,
    PP_PLACEHOLDER.OBJECT,
    PP_PLACEHOLDER.VERTICAL_BODY,
    PP_PLACEHOLDER.VERTICAL_OBJECT,
}

SUBTITLE_TYPES = {
    PP_PLACEHOLDER.SUBTITLE,
}


@dataclass
class TemplateStyleProfile:
    template_path: Path
    title_layout_index: int
    content_layout_index: int
    title_slide_index: int | None
    content_slide_index: int | None
    title_score: float
    content_score: float
    confidence: float
    selection_reason: str
    title_layout_name: str
    content_layout_name: str


TemplateLayoutProfile = TemplateStyleProfile


def _placeholder_type(shape):
    if not shape.is_placeholder:
        return None
    return shape.placeholder_format.type


def _placeholder_summary(shape) -> dict:
    summary = {
        "name": shape.name,
        "is_placeholder": bool(shape.is_placeholder),
        "left": int(shape.left),
        "top": int(shape.top),
        "width": int(shape.width),
        "height": int(shape.height),
    }
    if shape.is_placeholder:
        summary["idx"] = shape.placeholder_format.idx
        summary["type"] = str(shape.placeholder_format.type)
    if getattr(shape, "has_text_frame", False):
        summary["text"] = shape.text_frame.text[:120]
    return summary


def _layout_features(layout) -> dict:
    placeholders = list(layout.placeholders)
    placeholder_types = [_placeholder_type(shape) for shape in placeholders]
    layout_name = (layout.name or "").lower()

    return {
        "name": layout.name or "",
        "name_lower": layout_name,
        "placeholder_count": len(placeholders),
        "title_count": sum(1 for item in placeholder_types if item in TITLE_TYPES),
        "body_count": sum(1 for item in placeholder_types if item in BODY_TYPES),
        "subtitle_count": sum(1 for item in placeholder_types if item in SUBTITLE_TYPES),
        "has_title_name": "title" in layout_name,
        "has_content_name": any(word in layout_name for word in ("content", "text", "body")),
        "has_section_name": "section" in layout_name,
        "is_blank": "blank" in layout_name,
    }


def _slide_text_shapes(slide) -> list:
    return [
        shape
        for shape in slide.shapes
        if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip()
    ]


def _slide_features(slide) -> dict:
    text_shapes = _slide_text_shapes(slide)
    texts = [shape.text_frame.text.strip() for shape in text_shapes]
    line_count = sum(len([line for line in text.splitlines() if line.strip()]) for text in texts)
    bulletish_count = sum(text.count("\n") for text in texts)
    largest_area = max((shape.width * shape.height for shape in text_shapes), default=0)
    top_text_count = sum(1 for shape in text_shapes if shape.top < 2_000_000)

    return {
        "text_shape_count": len(text_shapes),
        "line_count": line_count,
        "bulletish_count": bulletish_count,
        "largest_area": largest_area,
        "top_text_count": top_text_count,
        "texts": texts,
    }


def _score_title_slide(features: dict) -> float:
    score = 0.0
    score += 3.0 if features["text_shape_count"] <= 3 else 0.0
    score += 2.0 if 1 <= features["line_count"] <= 4 else 0.0
    score += 1.0 if features["top_text_count"] else 0.0
    score -= 2.0 if features["line_count"] > 8 else 0.0
    score -= 1.0 if features["bulletish_count"] > 3 else 0.0
    return score


def _score_content_slide(features: dict) -> float:
    score = 0.0
    score += 2.5 if features["text_shape_count"] >= 2 else 0.0
    score += 3.0 if features["line_count"] >= 5 else 0.0
    score += 1.0 if features["bulletish_count"] >= 2 else 0.0
    score += 0.5 if features["top_text_count"] else 0.0
    score -= 1.5 if features["line_count"] <= 3 else 0.0
    return score


def _score_title_layout(features: dict) -> float:
    score = 0.0
    score += 4.0 if features["title_count"] else 0.0
    score += 2.5 if features["subtitle_count"] else 0.0
    score += 1.5 if features["has_title_name"] else 0.0
    score += 0.5 if features["placeholder_count"] <= 3 else 0.0
    score -= 1.5 if features["body_count"] > 1 else 0.0
    score -= 3.0 if features["is_blank"] else 0.0
    return score


def _score_content_layout(features: dict) -> float:
    score = 0.0
    score += 3.5 if features["title_count"] else 0.0
    score += 4.0 if features["body_count"] else 0.0
    score += 1.5 if features["has_content_name"] else 0.0
    score += 0.5 if features["placeholder_count"] >= 2 else 0.0
    score -= 2.0 if features["subtitle_count"] and not features["body_count"] else 0.0
    score -= 3.0 if features["is_blank"] else 0.0
    return score


def _used_layout_indices(prs: Presentation) -> set[int]:
    indices = set()
    layouts = list(prs.slide_layouts)
    for slide in prs.slides:
        try:
            indices.add(layouts.index(slide.slide_layout))
        except ValueError:
            continue
    return indices


def analyze_template_style(template_path) -> TemplateStyleProfile:
    template_path = Path(template_path)
    prs = Presentation(template_path)
    layouts = list(prs.slide_layouts)
    if not layouts:
        raise ValueError("Template does not contain any slide layouts.")

    used_indices = _used_layout_indices(prs)
    scored = []
    for index, layout in enumerate(layouts):
        features = _layout_features(layout)
        title_score = _score_title_layout(features)
        content_score = _score_content_layout(features)

        if index in used_indices:
            title_score += 0.5
            content_score += 0.5

        scored.append((index, layout, features, title_score, content_score))

    title_choice = max(scored, key=lambda item: item[3])
    content_candidates = [item for item in scored if item[0] != title_choice[0]] or scored
    content_choice = max(content_candidates, key=lambda item: item[4])

    title_score = max(title_choice[3], 0.0)
    content_score = max(content_choice[4], 0.0)
    title_slide_index = None
    content_slide_index = None
    slide_title_score = 0.0
    slide_content_score = 0.0

    slide_scored = []
    for index, slide in enumerate(prs.slides, start=1):
        features = _slide_features(slide)
        slide_scored.append(
            (
                index,
                slide,
                features,
                _score_title_slide(features),
                _score_content_slide(features),
            )
        )

    if slide_scored and (title_score < 2.0 or content_score < 2.0):
        title_slide = max(slide_scored, key=lambda item: item[3])
        content_slide_candidates = [item for item in slide_scored if item[0] != title_slide[0]] or slide_scored
        content_slide = max(content_slide_candidates, key=lambda item: item[4])
        title_slide_index = title_slide[0]
        content_slide_index = content_slide[0]
        slide_title_score = max(title_slide[3], 0.0)
        slide_content_score = max(content_slide[4], 0.0)

    confidence = min(
        0.95,
        max(
            0.15,
            max(
                ((title_score / 8.0) + (content_score / 9.0)) / 2,
                ((slide_title_score / 6.0) + (slide_content_score / 7.0)) / 2,
            ),
        ),
    )

    reason = (
        f"Selected '{title_choice[2]['name']}' for title "
        f"and '{content_choice[2]['name']}' for content using placeholder scoring."
    )
    if title_slide_index or content_slide_index:
        reason += (
            f" Placeholder layouts were weak, so slide {title_slide_index} "
            f"and slide {content_slide_index} were selected as visual samples."
        )

    return TemplateStyleProfile(
        template_path=template_path,
        title_layout_index=title_choice[0],
        content_layout_index=content_choice[0],
        title_slide_index=title_slide_index,
        content_slide_index=content_slide_index,
        title_score=title_score,
        content_score=content_score,
        confidence=confidence,
        selection_reason=reason,
        title_layout_name=title_choice[2]["name"],
        content_layout_name=content_choice[2]["name"],
    )


def detect_template_layout(template_path, use_ai: bool = False) -> TemplateLayoutProfile:
    profile = analyze_template_style(template_path)
    return profile


def summarize_template_layouts(template_path) -> list[dict]:
    prs = Presentation(template_path)
    summaries = []
    used_indices = _used_layout_indices(prs)

    for index, layout in enumerate(prs.slide_layouts):
        features = _layout_features(layout)
        summaries.append(
            {
                "layout_index": index,
                "layout_name": layout.name,
                "used_by_template_slide": index in used_indices,
                "title_score": _score_title_layout(features),
                "content_score": _score_content_layout(features),
                "features": features,
                "placeholders": [
                    _placeholder_summary(shape)
                    for shape in layout.placeholders
                ],
            }
        )

    return summaries


def summarize_template_slides(template_path) -> list[dict]:
    prs = Presentation(template_path)
    layout_list = list(prs.slide_layouts)
    summaries = []

    for index, slide in enumerate(prs.slides, start=1):
        features = _slide_features(slide)
        try:
            layout_index = layout_list.index(slide.slide_layout)
        except ValueError:
            layout_index = None

        summaries.append(
            {
                "slide_index": index,
                "layout_index": layout_index,
                "layout_name": slide.slide_layout.name,
                "title_score": _score_title_slide(features),
                "content_score": _score_content_slide(features),
                "features": features,
                "text_shapes": [
                    _placeholder_summary(shape)
                    for shape in _slide_text_shapes(slide)
                ],
            }
        )

    return summaries


def format_template_detection_report(template_path, profile: TemplateStyleProfile) -> str:
    layouts = summarize_template_layouts(template_path)
    lines = [
        "Template detection report",
        f"Template: {Path(template_path)}",
        "",
        "Selected layouts",
        f"- Title: layout {profile.title_layout_index} ({profile.title_layout_name})",
        f"- Content: layout {profile.content_layout_index} ({profile.content_layout_name})",
        f"- Title sample slide: {profile.title_slide_index or 'not needed'}",
        f"- Content sample slide: {profile.content_slide_index or 'not needed'}",
        f"- Confidence: {profile.confidence:.2f}",
        f"- Reason: {profile.selection_reason}",
        "",
        "All layout scores",
    ]

    for item in layouts:
        marker = []
        if item["layout_index"] == profile.title_layout_index:
            marker.append("TITLE")
        if item["layout_index"] == profile.content_layout_index:
            marker.append("CONTENT")
        marker_text = f" [{' / '.join(marker)}]" if marker else ""
        lines.append(
            f"- {item['layout_index']}: {item['layout_name']}{marker_text} | "
            f"title_score={item['title_score']:.1f}, "
            f"content_score={item['content_score']:.1f}, "
            f"placeholders={item['features']['placeholder_count']}, "
            f"used={item['used_by_template_slide']}"
        )
        for placeholder in item["placeholders"]:
            placeholder_type = placeholder.get("type", "non-placeholder")
            text = placeholder.get("text", "").replace("\n", " ").strip()
            text_part = f", text='{text}'" if text else ""
            lines.append(
                f"    - {placeholder['name']} | {placeholder_type} | "
                f"{placeholder['width']}x{placeholder['height']} at "
                f"({placeholder['left']},{placeholder['top']}){text_part}"
            )

    lines.append("")
    lines.append("All slide scores")
    for item in summarize_template_slides(template_path):
        marker = []
        if item["slide_index"] == profile.title_slide_index:
            marker.append("TITLE_SAMPLE")
        if item["slide_index"] == profile.content_slide_index:
            marker.append("CONTENT_SAMPLE")
        marker_text = f" [{' / '.join(marker)}]" if marker else ""
        preview = " | ".join(item["features"]["texts"])[:180].replace("\n", " ")
        lines.append(
            f"- slide {item['slide_index']}: layout={item['layout_index']} "
            f"({item['layout_name']}){marker_text} | "
            f"title_score={item['title_score']:.1f}, "
            f"content_score={item['content_score']:.1f}, "
            f"text_shapes={item['features']['text_shape_count']}, "
            f"lines={item['features']['line_count']} | {preview}"
        )

    return "\n".join(lines)
