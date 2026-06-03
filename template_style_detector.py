from __future__ import annotations

from dataclasses import dataclass


# Describes editable/protected text roles on a title sample slide.
@dataclass
class TitleSlideStyle:
    title_id: int | None
    subtitle_id: int | None
    deletable_ids: set[int]


# Describes editable/protected text roles on a content sample slide.
@dataclass
class ContentSlideStyle:
    title_id: int | None
    body_id: int | None
    drop_cap_id: int | None
    drop_cap_body_id: int | None
    repeated_component_ids: list[int]
    repeated_component_description_ids: list[int]
    use_body_region: bool
    use_repeated_components: bool
    deletable_text_ids: set[int]
    footer_ids: set[int]
    body_confidence: float = 0.0
    body_candidate_ids: list[int] | None = None


# Normalizes text for loose label matching.
def normalized_label(text: str) -> str:
    return " ".join((text or "").lower().replace("\r", "\n").split())


# Counts non-empty words in a text string.
def word_count(text: str) -> int:
    return len([word for word in (text or "").replace("\r", "\n").split() if word.strip()])


# Counts non-empty lines in a text string.
def line_count(text: str) -> int:
    return len([line for line in (text or "").replace("\r", "\n").splitlines() if line.strip()])


# Detects small section/category labels that should not be treated as titles.
def is_short_eyebrow(item) -> bool:
    text = normalized_label(item["text"])
    if not text:
        return False
    if word_count(text) <= 4 and item["height"] < 40:
        return True
    return any(word in text for word in ("chapter", "section", "part", "issue")) and word_count(text) <= 5


# Detects a decorative single-letter drop cap.
def is_drop_cap_item(item) -> bool:
    text = item["text"].strip()
    return len(text) == 1 and text.isalpha() and item["width"] <= 120 and item["height"] >= 60


# Finds the body text box visually paired with a drop cap.
def drop_cap_body_partner(drop_cap_item, shapes):
    candidates = []
    drop_right = drop_cap_item["left"] + drop_cap_item["width"]
    drop_mid_y = drop_cap_item["top"] + drop_cap_item["height"] / 2

    for item in shapes:
        if item["id"] == drop_cap_item["id"]:
            continue
        if item["left"] < drop_cap_item["left"]:
            continue
        if item["left"] > drop_right + 160:
            continue
        item_mid_y = item["top"] + item["height"] / 2
        if abs(item_mid_y - drop_mid_y) > max(drop_cap_item["height"], item["height"]):
            continue
        if word_count(item["text"]) < 8:
            continue
        candidates.append(item)

    return min(candidates, key=lambda item: abs(item["left"] - drop_right)) if candidates else None


# Finds the first drop-cap and body-text pair on a slide.
def drop_cap_body_pair(shapes):
    for item in shapes:
        if not is_drop_cap_item(item):
            continue
        partner = drop_cap_body_partner(item, shapes)
        if partner is not None:
            return item, partner
    return None, None


# Scores how likely a text shape is to be a general body-copy region.
def body_region_confidence(item) -> float:
    lines = line_count(item["text"])
    words = word_count(item["text"])
    score = 0.0
    score += 0.35 if lines >= 4 else 0.18 if lines >= 2 else 0.0
    score += 0.30 if words >= 18 else 0.18 if words >= 8 else 0.0
    score += 0.20 if item["height"] >= 90 else 0.10 if item["height"] >= 50 else 0.0
    score += 0.15 if item["width"] >= 280 else 0.08 if item["width"] >= 180 else 0.0
    return min(score, 1.0)


# Converts the body-region score into the current boolean strategy flag.
def looks_like_body_region(item, confidence_threshold: float = 0.55) -> bool:
    return body_region_confidence(item) >= confidence_threshold


# Detects footer or page-number text that should usually be preserved.
def looks_like_footer_text(text: str) -> bool:
    label = normalized_label(text)
    if not label:
        return False
    if any(word in label for word in ("confidential", "copyright")):
        return True
    return label.replace(" ", "").isdigit() and len(label.replace(" ", "")) <= 3


# Returns a font-size value, falling back to line-adjusted box height.
def inferred_font_size(item) -> float:
    return item.get("font_size") or item["height"] / max(line_count(item["text"]), 1)


# Checks whether two numeric values are close enough for visual clustering.
def close_enough(left, right, absolute=10, relative=0.12) -> bool:
    return abs(left - right) <= max(absolute, max(abs(left), abs(right)) * relative)


# Checks whether two shapes have similar repeated-component geometry.
def similar_geometry(first, second) -> bool:
    return (
        close_enough(first["left"], second["left"], absolute=18, relative=0.08)
        and close_enough(first["width"], second["width"], absolute=28, relative=0.12)
        and close_enough(first["height"], second["height"], absolute=12, relative=0.20)
        and close_enough(inferred_font_size(first), inferred_font_size(second), absolute=4, relative=0.18)
    )


# Groups shapes that repeat the same visual geometry.
def geometry_clusters(shapes) -> list[list[dict]]:
    clusters = []
    for item in shapes:
        for cluster in clusters:
            if similar_geometry(item, cluster[0]):
                cluster.append(item)
                break
        else:
            clusters.append([item])
    return [
        sorted(cluster, key=lambda item: (item["top"], item["left"]))
        for cluster in clusters
        if len(cluster) >= 2
    ]


# Finds repeated component pairs using text-position heuristics that worked well before.
def text_row_component_pairs(shapes):
    pairs = []
    for heading in shapes:
        if looks_like_footer_text(heading["text"]):
            continue
        if word_count(heading["text"]) > 5:
            continue
        if heading["height"] > 48:
            continue
        if heading["width"] < 180:
            continue

        candidates = []
        for description in shapes:
            if description["id"] == heading["id"]:
                continue
            if description["top"] <= heading["top"]:
                continue
            if description["top"] - (heading["top"] + heading["height"]) > 24:
                continue
            if abs(description["left"] - heading["left"]) > 45:
                continue
            if description["width"] < heading["width"] * 0.6:
                continue
            if word_count(description["text"]) < 6:
                continue
            candidates.append(description)

        if candidates:
            pairs.append((heading, min(candidates, key=lambda item: item["top"])))

    if len(pairs) < 2:
        return []

    lefts = [heading["left"] for heading, _ in pairs]
    widths = [heading["width"] for heading, _ in pairs]
    if max(widths) - min(widths) > max(120, max(widths) * 0.25):
        return []

    return sorted(pairs, key=lambda pair: (pair[0]["top"], pair[0]["left"]))


# Finds repeated component pairs by matching repeated geometry clusters.
def geometry_row_component_pairs(shapes):
    clusters = geometry_clusters(
        [
            item
            for item in shapes
            if not looks_like_footer_text(item["text"])
        ]
    )
    pairs = []

    for heading_cluster in clusters:
        for description_cluster in clusters:
            if heading_cluster is description_cluster:
                continue
            cluster_pairs = []
            for heading in heading_cluster:
                candidates = [
                    description
                    for description in description_cluster
                    if description["top"] > heading["top"]
                    and description["top"] - (heading["top"] + heading["height"]) <= 36
                    and abs(description["left"] - heading["left"]) <= 60
                    and description["width"] >= heading["width"] * 0.6
                ]
                if candidates:
                    cluster_pairs.append((heading, min(candidates, key=lambda item: item["top"])))

            if len(cluster_pairs) >= 2:
                pairs.extend(cluster_pairs)

    by_heading_id = {}
    for heading, description in pairs:
        current = by_heading_id.get(heading["id"])
        if current is None or description["top"] < current[1]["top"]:
            by_heading_id[heading["id"]] = (heading, description)

    return sorted(by_heading_id.values(), key=lambda pair: (pair[0]["top"], pair[0]["left"]))


# Finds repeated heading/description pairs using both semantic-position and geometry signals.
def row_component_pairs(shapes):
    pairs_by_heading_id = {
        heading["id"]: (heading, description)
        for heading, description in text_row_component_pairs(shapes)
    }
    for heading, description in geometry_row_component_pairs(shapes):
        pairs_by_heading_id.setdefault(heading["id"], (heading, description))

    return sorted(pairs_by_heading_id.values(), key=lambda pair: (pair[0]["top"], pair[0]["left"]))


# Finds standalone repeated component boxes from repeated geometry clusters.
def standalone_repeated_components(shapes):
    repeated = []
    for cluster in geometry_clusters(
        [
            item
            for item in shapes
            if not looks_like_footer_text(item["text"])
            and word_count(item["text"]) >= 3
            and item["height"] >= 45
            and item["width"] >= 160
        ]
    ):
        if len(cluster) < 3:
            continue
        tops = [item["top"] for item in cluster]
        lefts = [item["left"] for item in cluster]
        spans_rows = max(tops) - min(tops) > max(item["height"] for item in cluster) * 0.8
        spans_columns = max(lefts) - min(lefts) > max(item["width"] for item in cluster) * 0.8
        if spans_rows or spans_columns:
            repeated.extend(cluster)

    by_id = {item["id"]: item for item in repeated}
    return sorted(by_id.values(), key=lambda item: (item["top"], item["left"]))


# Chooses the main title text shape on a title slide.
def choose_title_shape(shapes):
    # Scores title-slide shapes by text scale, alignment, height, and width.
    def score(item):
        lower = item["lower"]
        slide_left = min(shape["left"] for shape in shapes)
        slide_right = max(shape["left"] + shape["width"] for shape in shapes)
        slide_width = max(slide_right - slide_left, 1)
        item_center_x = item["left"] + item["width"] / 2
        slide_center_x = slide_left + slide_width / 2
        centeredness = 1 - min(abs(item_center_x - slide_center_x) / (slide_width / 2), 1)
        lines = max(line_count(item["text"]), 1)
        font_size = inferred_font_size(item)
        value = font_size * 35

        if "subtitle" in lower:
            value -= 700
        if "presentation title" in lower:
            value += 900
        elif "title" in lower and "subtitle" not in lower:
            value += 500
        if word_count(item["text"]) <= 6 and lines <= 2:
            value += 300
        if word_count(item["text"]) >= 12 or lines >= 3:
            value -= 450
        value += centeredness * 220
        value += min(item["height"], 220) * 1.2
        value += min(item["width"] / slide_width, 1.0) * 120
        if "category" in lower or "section label" in lower:
            value -= 500
        if "presenter" in lower or "date" in lower:
            value -= 500
        value -= item["top"] / 100
        return value

    return max(shapes, key=score) if shapes else None


# Chooses the subtitle/description text shape on a title slide.
def choose_subtitle_shape(shapes, title_shape):
    title_top = title_shape["top"] if title_shape else -1

    # Scores subtitle candidates by descriptive text, size, and position below the title.
    def score(item):
        lower = item["lower"]
        if looks_like_footer_text(item["text"]):
            return -10_000
        value = item["area"] / 20000
        if "subtitle" in lower or "description" in lower:
            value += 500
        if item["top"] > title_top:
            value += 100
        if "presenter" in lower or "date" in lower:
            value -= 500
        return value

    candidates = [item for item in shapes if item is not title_shape]
    return max(candidates, key=score) if candidates else None


# Chooses the main title text shape on a content slide.
def choose_content_title_shape(shapes, repeated_component_ids=None):
    repeated_component_ids = set(repeated_component_ids or [])
    candidates = [
        item for item in shapes
        if item["id"] not in repeated_component_ids
    ] or shapes
    top_shapes = candidates[:5]

    # Scores content-title candidates among the highest text shapes.
    def score(item):
        lower = item["lower"]
        value = item["area"] / 10000
        if "content slide title" in lower:
            value += 1000
        elif "title" in lower:
            value += 600
        if is_short_eyebrow(item):
            value -= 900
        if word_count(item["text"]) >= 3:
            value += 250
        if item["height"] >= 45:
            value += 250
        if "section label" in lower or "category" in lower:
            value -= 700
        if len(item["text"].splitlines()) > 2:
            value -= 500
        value -= item["top"] / 10
        return value

    return max(top_shapes or candidates, key=score) if candidates else None


# Chooses the best single body text shape while avoiding titles and repeated components.
def choose_content_body_shape(shapes, title_shape, extra_component_ids=None):
    row_pairs = row_component_pairs(shapes)
    component_ids = {heading["id"] for heading, _ in row_pairs}
    component_ids.update(description["id"] for _, description in row_pairs)
    component_ids.update(extra_component_ids or [])
    drop_cap_item, drop_cap_partner = drop_cap_body_pair(shapes)

    # Scores body candidates by usable text-region confidence while avoiding components.
    def score(item):
        value = item["area"] / 5000
        if item is drop_cap_partner:
            value += 1600
        if item is drop_cap_item:
            value -= 1600
        if item["id"] in component_ids:
            value -= 2000
        value += body_region_confidence(item) * 1000
        if item is title_shape:
            value -= 2000
        if item["top"] <= (title_shape["top"] if title_shape else 0):
            value -= 500
        return value

    candidates = [item for item in shapes if item is not title_shape]
    return max(candidates, key=score) if candidates else None


# Detects editable title/subtitle roles for a title sample slide.
def detect_title_slide_style(shapes) -> TitleSlideStyle:
    title_shape = choose_title_shape(shapes)
    subtitle_shape = choose_subtitle_shape(shapes, title_shape)
    protected_ids = {
        item["id"]
        for item in (title_shape, subtitle_shape)
        if item is not None
    }
    return TitleSlideStyle(
        title_id=title_shape["id"] if title_shape else None,
        subtitle_id=subtitle_shape["id"] if subtitle_shape else None,
        deletable_ids={item["id"] for item in shapes if item["id"] not in protected_ids},
    )


# Detects editable title/body/component roles for a content sample slide.
def detect_content_slide_style(shapes) -> ContentSlideStyle:
    row_pairs = row_component_pairs(shapes)
    paired_components = [heading for heading, _ in row_pairs]
    paired_component_ids = {item["id"] for item in paired_components}
    standalone_components = [
        item for item in standalone_repeated_components(shapes)
        if item["id"] not in paired_component_ids
    ]
    repeated_components = paired_components + standalone_components
    repeated_component_ids = {item["id"] for item in repeated_components}
    title_shape = choose_content_title_shape(shapes, repeated_component_ids=repeated_component_ids)
    body_shape = choose_content_body_shape(shapes, title_shape, extra_component_ids=repeated_component_ids)
    drop_cap_item, drop_cap_partner = drop_cap_body_pair(shapes)
    body_candidates = sorted(
        [
            item
            for item in shapes
            if item is not title_shape and body_region_confidence(item) >= 0.35
        ],
        key=body_region_confidence,
        reverse=True,
    )
    body_confidence = body_region_confidence(body_shape) if body_shape else 0.0

    use_body_region = len(repeated_components) < 2 and body_shape is not None and (
        looks_like_body_region(body_shape) or body_shape is drop_cap_partner
    )
    use_repeated_components = len(repeated_components) >= 2 and not use_body_region

    protected_ids = {
        item["id"]
        for item in (title_shape, body_shape, drop_cap_item, drop_cap_partner)
        if item is not None
    }
    component_ids = {item["id"] for item in repeated_components}
    repeated_component_id_set = {component["id"] for component in repeated_components}
    component_description_ids = {
        description["id"]
        for heading, description in row_pairs
        if heading["id"] in repeated_component_id_set
    }
    footer_ids = {item["id"] for item in shapes if looks_like_footer_text(item["text"])}

    deletable_text_ids = {
        item["id"]
        for item in shapes
        if item["id"] not in protected_ids
        and item["id"] not in component_ids
        and item["id"] not in component_description_ids
        and item["id"] not in footer_ids
    }

    return ContentSlideStyle(
        title_id=title_shape["id"] if title_shape else None,
        body_id=body_shape["id"] if body_shape else None,
        drop_cap_id=drop_cap_item["id"] if drop_cap_item else None,
        drop_cap_body_id=drop_cap_partner["id"] if drop_cap_partner else None,
        repeated_component_ids=[item["id"] for item in repeated_components],
        repeated_component_description_ids=[
            description["id"]
            for heading, description in row_pairs
            if heading["id"] in repeated_component_id_set
        ],
        use_body_region=use_body_region,
        use_repeated_components=use_repeated_components,
        deletable_text_ids=deletable_text_ids,
        footer_ids=footer_ids,
        body_confidence=body_confidence,
        body_candidate_ids=[item["id"] for item in body_candidates],
    )
