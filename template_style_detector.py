from __future__ import annotations

from dataclasses import dataclass


@dataclass
class TitleSlideStyle:
    title_id: int | None
    subtitle_id: int | None
    deletable_ids: set[int]


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


def normalized_label(text: str) -> str:
    return " ".join((text or "").lower().replace("\r", "\n").split())


def word_count(text: str) -> int:
    return len([word for word in (text or "").replace("\r", "\n").split() if word.strip()])


def line_count(text: str) -> int:
    return len([line for line in (text or "").replace("\r", "\n").splitlines() if line.strip()])


def is_short_eyebrow(item) -> bool:
    text = normalized_label(item["text"])
    if not text:
        return False
    if word_count(text) <= 4 and item["height"] < 40:
        return True
    return any(word in text for word in ("chapter", "section", "part", "issue")) and word_count(text) <= 5


def is_drop_cap_item(item) -> bool:
    text = item["text"].strip()
    return len(text) == 1 and text.isalpha() and item["width"] <= 120 and item["height"] >= 60


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


def drop_cap_body_pair(shapes):
    for item in shapes:
        if not is_drop_cap_item(item):
            continue
        partner = drop_cap_body_partner(item, shapes)
        if partner is not None:
            return item, partner
    return None, None


def looks_like_body_region(item) -> bool:
    lower = item["lower"]
    if "key point" in lower or "bullet" in lower:
        return True
    return line_count(item["text"]) >= 4 and item["area"] > 1_000_000_000_000


def looks_like_footer_text(text: str) -> bool:
    label = normalized_label(text)
    if not label:
        return False
    if any(word in label for word in ("company", "confidential", "copyright")):
        return True
    return label.replace(" ", "").isdigit() and len(label.replace(" ", "")) <= 3


def is_generic_component_label(text: str) -> bool:
    label = normalized_label(text)
    if not label:
        return False
    if len(label.split()) > 4:
        return False
    if any(char.isdigit() for char in label):
        return False
    if any(word in label for word in ("company", "confidential", "date", "presenter")):
        return False
    if any(word in label for word in ("description", "supporting", "detail", "context")):
        return False
    return label not in {"title", "subtitle", "section label", "category"}


def repeated_component_items(shapes):
    repeated_words = ("highlight", "metric", "insight", "callout")
    explicit = [item for item in shapes if any(word in item["lower"] for word in repeated_words)]

    label_counts = {}
    for item in shapes:
        label = normalized_label(item["text"])
        if is_generic_component_label(label):
            label_counts[label] = label_counts.get(label, 0) + 1

    repeated_labels = {label for label, count in label_counts.items() if count >= 2}
    generic = [item for item in shapes if normalized_label(item["text"]) in repeated_labels]
    by_id = {item["id"]: item for item in explicit + generic}
    return sorted(by_id.values(), key=lambda item: (item["top"], item["left"]))


def row_component_pairs(shapes):
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
            if abs(description["width"] - heading["width"]) > max(90, heading["width"] * 0.2):
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


def choose_title_shape(shapes):
    def score(item):
        lower = item["lower"]
        value = item["area"] / 10000
        if "presentation title" in lower:
            value += 900
        elif "title" in lower:
            value += 500
        if "category" in lower or "section label" in lower:
            value -= 500
        if "presenter" in lower or "date" in lower:
            value -= 500
        value -= item["top"] / 100
        return value

    return max(shapes, key=score) if shapes else None


def choose_subtitle_shape(shapes, title_shape):
    title_top = title_shape["top"] if title_shape else -1

    def score(item):
        lower = item["lower"]
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


def choose_content_title_shape(shapes):
    top_shapes = shapes[:5]

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

    return max(top_shapes or shapes, key=score) if shapes else None


def choose_content_body_shape(shapes, title_shape):
    row_pairs = row_component_pairs(shapes)
    component_ids = {item["id"] for item in repeated_component_items(shapes)}
    component_ids.update(heading["id"] for heading, _ in row_pairs)
    component_ids.update(description["id"] for _, description in row_pairs)
    drop_cap_item, drop_cap_partner = drop_cap_body_pair(shapes)

    def score(item):
        lower = item["lower"]
        value = item["area"] / 5000
        if item is drop_cap_partner:
            value += 1600
        if item is drop_cap_item:
            value -= 1600
        if item["id"] in component_ids:
            value -= 2000
        if "key point" in lower or "add a sentence" in lower:
            value += 1000
        if item is title_shape:
            value -= 2000
        if item["top"] <= (title_shape["top"] if title_shape else 0):
            value -= 500
        return value

    candidates = [item for item in shapes if item is not title_shape]
    return max(candidates, key=score) if candidates else None


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


def detect_content_slide_style(shapes) -> ContentSlideStyle:
    title_shape = choose_content_title_shape(shapes)
    row_pairs = row_component_pairs(shapes)
    repeated_components = repeated_component_items(shapes)
    if row_pairs and not repeated_components:
        repeated_components = [heading for heading, _ in row_pairs]
    body_shape = choose_content_body_shape(shapes, title_shape)
    drop_cap_item, drop_cap_partner = drop_cap_body_pair(shapes)

    use_body_region = body_shape is not None and (
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
    )
