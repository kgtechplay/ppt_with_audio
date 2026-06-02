from __future__ import annotations

from pathlib import Path

from pptx import Presentation

from template_layout_detector import TemplateLayoutProfile
from template_style_detector import detect_content_slide_style, detect_title_slide_style


class _DebugLogger:
    def __init__(self, log_path: Path | None):
        self.log_path = log_path
        if self.log_path:
            self.log_path.parent.mkdir(parents=True, exist_ok=True)
            self.log_path.write_text("", encoding="utf-8")

    def write(self, message: str) -> None:
        line = f"[template-com] {message}"
        print(line, flush=True)
        if self.log_path:
            with self.log_path.open("a", encoding="utf-8") as file:
                file.write(line + "\n")


def _preview(text: str, limit: int = 90) -> str:
    text = " ".join((text or "").split())
    if len(text) <= limit:
        return text
    return text[: limit - 3] + "..."


def _item_summary(item) -> str:
    return (
        f"id={item.get('id')} parents={item.get('parent_ids', ())} "
        f"box=({item.get('left'):.0f},{item.get('top'):.0f},"
        f"{item.get('width'):.0f},{item.get('height'):.0f}) "
        f"text='{_preview(item.get('text', ''))}'"
    )


PLACEHOLDER_WORDS = (
    "title",
    "subtitle",
    "description",
    "presenter",
    "date",
    "category",
    "section label",
    "key point",
    "highlight",
    "metric",
    "insight",
    "callout",
    "short card description",
    "supporting",
    "goes right here",
    "add a sentence",
)

REPEATED_COMPONENT_WORDS = (
    "highlight",
    "metric",
    "insight",
    "callout",
)


def _shape_bounds(shape) -> tuple[float, float, float, float] | None:
    try:
        return (
            float(shape.Left),
            float(shape.Top),
            float(shape.Width),
            float(shape.Height),
        )
    except Exception:
        return None


def _shape_id(shape) -> int | None:
    try:
        return int(shape.Id)
    except Exception:
        return None


def _iter_shape_tree(shapes_collection, parent_ids=()):
    for index in range(1, shapes_collection.Count + 1):
        shape = shapes_collection(index)
        shape_id = _shape_id(shape)
        if shape_id is None:
            continue

        yield shape, parent_ids

        try:
            group_items = shape.GroupItems
            group_count = group_items.Count
        except Exception:
            group_count = 0

        if group_count:
            yield from _iter_shape_tree(group_items, parent_ids + (shape_id,))


def _text(shape) -> str:
    try:
        if not shape.HasTextFrame:
            return ""
        if not shape.TextFrame.HasText:
            return ""
        return shape.TextFrame.TextRange.Text.strip()
    except Exception:
        return ""


def _text_shapes(slide):
    shapes = []
    for shape, parent_ids in _iter_shape_tree(slide.Shapes):
        text = _text(shape)
        if not text:
            continue
        bounds = _shape_bounds(shape)
        if bounds is None:
            continue
        left, top, width, height = bounds
        shape_id = _shape_id(shape)
        shapes.append(
            {
                "shape": shape,
                "id": shape_id,
                "parent_ids": parent_ids,
                "text": text,
                "lower": text.lower(),
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "area": width * height,
            }
        )
    return sorted(shapes, key=lambda item: (item["top"], item["left"]))


def _all_shape_items(slide):
    items = []
    for shape, parent_ids in _iter_shape_tree(slide.Shapes):
        bounds = _shape_bounds(shape)
        shape_id = _shape_id(shape)
        if bounds is None or shape_id is None:
            continue
        left, top, width, height = bounds
        items.append(
            {
                "shape": shape,
                "id": shape_id,
                "parent_ids": parent_ids,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "right": left + width,
                "bottom": top + height,
                "area": width * height,
                "cx": left + width / 2,
                "cy": top + height / 2,
                "text": _text(shape),
            }
        )
    return items


def _slide_by_id(presentation, slide_id):
    return presentation.Slides.FindBySlideID(slide_id)


def _duplicate_to_end(presentation, source_slide_id):
    source = _slide_by_id(presentation, source_slide_id)
    duplicate_range = source.Duplicate()
    duplicate = duplicate_range.Item(1)
    duplicate.MoveTo(presentation.Slides.Count)
    return presentation.Slides(presentation.Slides.Count)


def _set_shape_text(shape, text: str) -> None:
    text_range = shape.TextFrame.TextRange
    text_range.Text = text


def _clear_shape_text(shape) -> None:
    try:
        shape.TextFrame.TextRange.Text = ""
    except Exception:
        pass


def _delete_shape(shape) -> None:
    try:
        shape.Delete()
    except Exception:
        pass


def _slide_area(slide) -> float:
    try:
        setup = slide.Parent.PageSetup
        return float(setup.SlideWidth) * float(setup.SlideHeight)
    except Exception:
        return 720 * 540


def _contains_point(item, x, y) -> bool:
    return item["left"] <= x <= item["right"] and item["top"] <= y <= item["bottom"]


def _overlaps(a, b) -> bool:
    return not (
        a["right"] < b["left"]
        or a["left"] > b["right"]
        or a["bottom"] < b["top"]
        or a["top"] > b["bottom"]
    )


def _expanded(item, padding) -> dict:
    return {
        **item,
        "left": item["left"] - padding,
        "top": item["top"] - padding,
        "right": item["right"] + padding,
        "bottom": item["bottom"] + padding,
    }


def _delete_unused_component(slide, text_item, protected_ids=None, logger=None) -> None:
    protected_ids = set(protected_ids or [])
    shape_items = _all_shape_items(slide)
    text_id = int(text_item["id"])
    protected_parent_ids = {
        parent_id
        for item in shape_items
        if item["id"] in protected_ids
        for parent_id in item.get("parent_ids", ())
    }
    text_box = {
        "left": text_item["left"],
        "top": text_item["top"],
        "right": text_item["left"] + text_item["width"],
        "bottom": text_item["top"] + text_item["height"],
        "area": text_item["area"],
        "cx": text_item["left"] + text_item["width"] / 2,
        "cy": text_item["top"] + text_item["height"] / 2,
    }
    slide_area = _slide_area(slide)

    containers = [
        item
        for item in shape_items
        if item["id"] != text_id
        and item["id"] not in protected_parent_ids
        and not item["text"]
        and item["area"] > text_box["area"] * 1.4
        and item["area"] < slide_area * 0.45
        and _contains_point(item, text_box["cx"], text_box["cy"])
    ]

    if containers:
        container = min(containers, key=lambda item: item["area"])
        delete_region = _expanded(container, 14)
        delete_ids = {
            item["id"]
            for item in shape_items
            if item["area"] < slide_area * 0.45
            and item["id"] not in protected_parent_ids
            and (
                _contains_point(delete_region, item["cx"], item["cy"])
                or _overlaps(delete_region, item)
            )
        }
    else:
        container = None
        delete_region = _expanded(text_box, 36)
        delete_ids = {
            item["id"]
            for item in shape_items
            if item["id"] == text_id
            or (
                item["area"] < slide_area * 0.15
                and item["id"] not in protected_parent_ids
                and _contains_point(delete_region, item["cx"], item["cy"])
            )
        }

    delete_ids = delete_ids - protected_ids - protected_parent_ids

    if logger:
        container_text = _item_summary(container) if container else "none"
        logger.write(
            "delete unused component from "
            f"{_item_summary(text_item)} | container={container_text} | "
            f"protected={sorted(protected_ids)} protected_parents={sorted(protected_parent_ids)} "
            f"delete_ids={sorted(delete_ids)}"
        )

    for item in reversed(shape_items):
        if item["id"] in delete_ids:
            _delete_shape(item["shape"])
    return delete_ids


def _component_container(slide, text_item):
    shape_items = _all_shape_items(slide)
    text_box = {
        "left": text_item["left"],
        "top": text_item["top"],
        "right": text_item["left"] + text_item["width"],
        "bottom": text_item["top"] + text_item["height"],
        "area": text_item["area"],
        "cx": text_item["left"] + text_item["width"] / 2,
        "cy": text_item["top"] + text_item["height"] / 2,
    }
    slide_area = _slide_area(slide)
    containers = [
        item
        for item in shape_items
        if item["id"] != text_item["id"]
        and not item["text"]
        and item["area"] > text_box["area"] * 1.4
        and item["area"] < slide_area * 0.45
        and _contains_point(item, text_box["cx"], text_box["cy"])
    ]
    return min(containers, key=lambda item: item["area"]) if containers else None


def _normalized_label(text: str) -> str:
    return " ".join(text.lower().replace("\r", "\n").split())


def _is_generic_component_label(text: str) -> bool:
    label = _normalized_label(text)
    if not label:
        return False
    words = label.split()
    if len(words) > 4:
        return False
    if any(char.isdigit() for char in label):
        return False
    if any(word in label for word in ("company", "confidential", "date", "presenter")):
        return False
    if any(word in label for word in ("description", "supporting", "detail", "context")):
        return False
    if label in {"title", "subtitle", "section label", "category"}:
        return False
    return True


def _repeated_component_items(shapes):
    explicit_components = [
        item
        for item in shapes
        if any(word in item["lower"] for word in REPEATED_COMPONENT_WORDS)
    ]

    label_counts = {}
    for item in shapes:
        label = _normalized_label(item["text"])
        if _is_generic_component_label(label):
            label_counts[label] = label_counts.get(label, 0) + 1

    repeated_labels = {
        label
        for label, count in label_counts.items()
        if count >= 2
    }

    generic_components = [
        item
        for item in shapes
        if _normalized_label(item["text"]) in repeated_labels
    ]

    components_by_id = {
        item["id"]: item
        for item in explicit_components + generic_components
    }
    return sorted(components_by_id.values(), key=lambda item: (item["top"], item["left"]))


def _fill_repeated_components(slide, components, values, protected_ids, logger=None):
    used_ids = set()
    deleted_ids = set()

    for item, value in zip(components, values):
        _set_shape_text(item["shape"], value)
        if logger:
            logger.write(f"fill repeated component {_item_summary(item)} with '{_preview(value)}'")
        protected_ids.add(item["id"])
        used_ids.add(item["id"])

    for item in components[len(values):]:
        if item["id"] in deleted_ids:
            continue
        description_item = description_by_component_id.get(item["id"])
        if description_item is not None and description_item["id"] not in deleted_ids:
            if logger:
                logger.write(
                    f"delete unused component description {_item_summary(description_item)}"
                )
            _delete_shape(description_item["shape"])
            deleted_ids.add(description_item["id"])
        deleted_ids.update(_delete_unused_component(slide, item, protected_ids=protected_ids, logger=logger))

    return used_ids, deleted_ids


def _line_count(text: str) -> int:
    return len([line for line in text.replace("\r", "\n").splitlines() if line.strip()])


def _word_count(text: str) -> int:
    return len([word for word in text.replace("\r", "\n").split() if word.strip()])


def _is_short_eyebrow(item) -> bool:
    text = _normalized_label(item["text"])
    if not text:
        return False
    if _word_count(text) <= 4 and item["height"] < 40:
        return True
    if any(word in text for word in ("chapter", "section", "part", "issue")) and _word_count(text) <= 5:
        return True
    return False


def _is_drop_cap_item(item) -> bool:
    text = item["text"].strip()
    return len(text) == 1 and text.isalpha() and item["width"] <= 120 and item["height"] >= 60


def _drop_cap_body_partner(drop_cap_item, shapes):
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
        if _word_count(item["text"]) < 8:
            continue
        candidates.append(item)

    return min(candidates, key=lambda item: abs(item["left"] - drop_right)) if candidates else None


def _drop_cap_body_pair(shapes):
    for item in shapes:
        if not _is_drop_cap_item(item):
            continue
        partner = _drop_cap_body_partner(item, shapes)
        if partner is not None:
            return item, partner
    return None, None


def _looks_like_body_region(item) -> bool:
    lower = item["lower"]
    if "key point" in lower or "bullet" in lower:
        return True
    if _line_count(item["text"]) >= 4 and item["area"] > 1_000_000_000_000:
        return True
    return False


def _looks_like_template_text(text: str) -> bool:
    lower = text.lower()
    return any(word in lower for word in PLACEHOLDER_WORDS)


def _looks_like_footer_text(text: str) -> bool:
    label = _normalized_label(text)
    if not label:
        return False
    if any(word in label for word in ("company", "confidential", "copyright")):
        return True
    if label.replace(" ", "").isdigit() and len(label.replace(" ", "")) <= 3:
        return True
    return False


def _should_delete_unmapped_content_text(slide, item) -> bool:
    if _looks_like_footer_text(item["text"]):
        return False

    return True


def _choose_title_shape(shapes):
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


def _choose_subtitle_shape(shapes, title_shape):
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


def _choose_content_title_shape(shapes):
    top_shapes = shapes[:5]

    def score(item):
        lower = item["lower"]
        value = item["area"] / 10000
        if "content slide title" in lower:
            value += 1000
        elif "title" in lower:
            value += 600
        if _is_short_eyebrow(item):
            value -= 900
        if _word_count(item["text"]) >= 3:
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


def _choose_content_body_shape(shapes, title_shape):
    component_ids = {item["id"] for item in _repeated_component_items(shapes)}
    drop_cap_item, drop_cap_partner = _drop_cap_body_pair(shapes)

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


def _format_body_text(bullets: list[str]) -> str:
    return "\r".join(bullets)


def _set_body_text(slide, body_shape, bullets: list[str], shapes, protected_ids, logger=None):
    drop_cap_item, drop_cap_partner = _drop_cap_body_pair(shapes)
    body_text = _format_body_text(bullets)

    if body_shape is drop_cap_partner and drop_cap_item is not None:
        if bullets:
            first_text = " ".join(bullets[0].split())
            first_letter = first_text[:1]
            remainder = first_text[1:].lstrip()
            remaining_lines = [remainder] if remainder else []
            remaining_lines.extend(bullets[1:])
            _set_shape_text(drop_cap_item["shape"], first_letter)
            _set_shape_text(drop_cap_partner["shape"], "\r".join(remaining_lines))
            protected_ids.add(drop_cap_item["id"])
            protected_ids.add(drop_cap_partner["id"])
            if logger:
                logger.write(
                    "set drop-cap body pair: "
                    f"drop_cap={_item_summary(drop_cap_item)} -> '{_preview(first_letter)}', "
                    f"body={_item_summary(drop_cap_partner)} with {len(remaining_lines)} line(s)"
                )
            return

    _set_shape_text(body_shape["shape"], body_text)
    protected_ids.add(body_shape["id"])
    if logger:
        logger.write(
            f"set body shape id={body_shape['id']} with {len(bullets)} bullet(s): "
            f"{[_preview(bullet, 50) for bullet in bullets]}"
        )


def _protect_component_region(slide, component_item, protected_ids):
    container = _component_container(slide, component_item)
    if container is None:
        protected_ids.add(component_item["id"])
        return

    shape_items = _all_shape_items(slide)
    region = _expanded(container, 14)
    for item in shape_items:
        if (
            _contains_point(region, item["cx"], item["cy"])
            or _overlaps(region, item)
        ):
            protected_ids.add(item["id"])


def _component_description_item(slide, component_item, all_text_items):
    container = _component_container(slide, component_item)
    if container is None:
        return None
    region = _expanded(container, 8)
    candidates = [
        item
        for item in all_text_items
        if item["id"] != component_item["id"]
        and _contains_point(region, item["left"] + item["width"] / 2, item["top"] + item["height"] / 2)
        and item["top"] >= component_item["top"]
    ]
    return max(candidates, key=lambda item: item["top"]) if candidates else None


def _split_component_value(value: str) -> tuple[str, str]:
    value = " ".join(value.split())
    if not value:
        return "", ""
    separators = [": ", " - ", ". "]
    for separator in separators:
        if separator in value:
            head, tail = value.split(separator, 1)
            if head and tail:
                return head.strip(), tail.strip()
    words = value.split()
    if len(words) <= 6:
        return value, ""
    return " ".join(words[:6]), " ".join(words[6:])


def _fill_repeated_card_components(
    slide,
    components,
    values,
    protected_ids,
    all_text_items,
    logger=None,
    description_by_component_id=None,
):
    used_ids = set()
    deleted_ids = set()
    description_by_component_id = description_by_component_id or {}

    for item, value in zip(components, values):
        title_text, description_text = _split_component_value(value)
        description_item = description_by_component_id.get(item["id"])
        if description_item is None:
            description_item = _component_description_item(slide, item, all_text_items)

        _set_shape_text(item["shape"], title_text)
        if logger:
            logger.write(
                f"fill card title {_item_summary(item)} with '{_preview(title_text)}'"
            )
        protected_ids.add(item["id"])
        used_ids.add(item["id"])

        if description_item is not None:
            if description_text:
                _set_shape_text(description_item["shape"], description_text)
                if logger:
                    logger.write(
                        "fill card description "
                        f"{_item_summary(description_item)} with '{_preview(description_text)}'"
                    )
            else:
                _clear_shape_text(description_item["shape"])
                if logger:
                    logger.write(
                        f"clear empty card description {_item_summary(description_item)}"
                    )
            protected_ids.add(description_item["id"])
            used_ids.add(description_item["id"])

        _protect_component_region(slide, item, protected_ids)

    for item in components[len(values):]:
        if item["id"] in deleted_ids:
            continue
        deleted_ids.update(_delete_unused_component(slide, item, protected_ids=protected_ids, logger=logger))

    return used_ids, deleted_ids


def _delete_repeated_components(slide, components, protected_ids, logger=None):
    deleted_ids = set()
    for item in components:
        if item["id"] in deleted_ids:
            continue
        deleted_ids.update(_delete_unused_component(slide, item, protected_ids=protected_ids, logger=logger))
    return deleted_ids


def _replace_title_slide_text(slide, title: str, logger=None) -> None:
    shapes = _text_shapes(slide)
    shape_by_id = {item["id"]: item for item in shapes}
    style = detect_title_slide_style(shapes)
    title_shape = shape_by_id.get(style.title_id)
    subtitle_shape = shape_by_id.get(style.subtitle_id)
    protected_ids = set()

    if logger:
        logger.write(f"title slide text shape count={len(shapes)}")
        for item in shapes:
            logger.write(f"title candidate: {_item_summary(item)}")
        logger.write(
            "title mapping: "
            f"title={_item_summary(title_shape) if title_shape else 'none'} | "
            f"subtitle={_item_summary(subtitle_shape) if subtitle_shape else 'none'}"
        )

    for item in shapes:
        if item is title_shape:
            _set_shape_text(item["shape"], title)
            if logger:
                logger.write(f"set title shape id={item['id']} to '{_preview(title)}'")
            protected_ids.add(item["id"])
        elif item is subtitle_shape:
            _set_shape_text(item["shape"], "Generated with OpenAI")
            if logger:
                logger.write(f"set subtitle shape id={item['id']} to 'Generated with OpenAI'")
            protected_ids.add(item["id"])

    deleted_ids = set()
    for item in shapes:
        if item["id"] in deleted_ids:
            continue
        if item is not title_shape and item is not subtitle_shape and _looks_like_template_text(item["text"]):
            deleted_ids.update(_delete_unused_component(slide, item, protected_ids=protected_ids, logger=logger))


def _replace_content_slide_text(slide, title: str, bullets: list[str], logger=None) -> None:
    shapes = _text_shapes(slide)
    shape_by_id = {item["id"]: item for item in shapes}
    style = detect_content_slide_style(shapes)
    title_shape = shape_by_id.get(style.title_id)
    repeated_components = [
        shape_by_id[shape_id]
        for shape_id in style.repeated_component_ids
        if shape_id in shape_by_id
    ]
    body_shape = shape_by_id.get(style.body_id)
    description_by_component_id = {
        component_id: shape_by_id[description_id]
        for component_id, description_id in zip(
            style.repeated_component_ids,
            style.repeated_component_description_ids,
        )
        if component_id in shape_by_id and description_id in shape_by_id
    }
    use_body_region = style.use_body_region
    use_repeated_components = style.use_repeated_components
    protected_ids = set()
    deleted_ids = set()

    if logger:
        logger.write(f"content slide '{_preview(title)}' text shape count={len(shapes)}")
        for item in shapes:
            logger.write(f"content candidate: {_item_summary(item)}")
        logger.write(
            "content mapping: "
            f"title={_item_summary(title_shape) if title_shape else 'none'} | "
            f"body={_item_summary(body_shape) if body_shape else 'none'} | "
            f"use_body_region={use_body_region} | "
            f"repeated_count={len(repeated_components)} | "
            f"use_repeated_components={use_repeated_components}"
        )
        logger.write(
            "style model: "
            f"drop_cap_id={style.drop_cap_id}, drop_cap_body_id={style.drop_cap_body_id}, "
            f"repeated_description_ids={style.repeated_component_description_ids}, "
            f"deletable_text_ids={sorted(style.deletable_text_ids)}, "
            f"footer_ids={sorted(style.footer_ids)}"
        )
        for item in repeated_components:
            logger.write(f"repeated component candidate: {_item_summary(item)}")

    for item in shapes:
        if item is title_shape:
            _set_shape_text(item["shape"], title)
            if logger:
                logger.write(f"set content title shape id={item['id']} to '{_preview(title)}'")
            protected_ids.add(item["id"])
        elif use_body_region and item is body_shape:
            _set_body_text(slide, body_shape, bullets, shapes, protected_ids, logger=logger)

    filled_component_ids = set()
    if use_repeated_components:
        filled_component_ids, component_deleted_ids = _fill_repeated_card_components(
            slide,
            repeated_components,
            bullets,
            protected_ids,
            shapes,
            logger=logger,
            description_by_component_id=description_by_component_id,
        )
        deleted_ids.update(component_deleted_ids)
    elif repeated_components:
        if logger:
            logger.write("body region selected; deleting unused repeated components")
        deleted_ids.update(_delete_repeated_components(slide, repeated_components, protected_ids, logger=logger))

    for item in shapes:
        if item["id"] in deleted_ids:
            continue
        if item["id"] in filled_component_ids:
            continue
        if item["id"] in protected_ids:
            continue
        if item is title_shape or item is body_shape:
            continue

        if _looks_like_template_text(item["text"]):
            deleted_ids.update(_delete_unused_component(slide, item, protected_ids=protected_ids, logger=logger))
        elif item["id"] in style.deletable_text_ids or _should_delete_unmapped_content_text(slide, item):
            if logger:
                logger.write(f"delete unmapped content text shape {_item_summary(item)}")
            _delete_shape(item["shape"])
            deleted_ids.add(item["id"])
        elif logger:
            logger.write(f"preserve unmapped text shape {_item_summary(item)}")


def _write_notes_with_python_pptx(output_file: Path, plan) -> None:
    prs = Presentation(output_file)

    if len(prs.slides) >= 1:
        prs.slides[0].notes_slide.notes_text_frame.text = ""

    for index, slide_data in enumerate(plan["slides"], start=1):
        if index >= len(prs.slides):
            break
        prs.slides[index].notes_slide.notes_text_frame.text = slide_data.get(
            "speaker_notes",
            "",
        )

    prs.save(output_file)


def build_presentation_from_template_com(
    plan,
    output_file,
    profile: TemplateLayoutProfile,
    debug_log_path=None,
):
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("PowerPoint COM template building requires pywin32.") from exc

    if not profile.title_slide_index or not profile.content_slide_index:
        raise ValueError("COM template builder requires detected title and content sample slides.")

    output_file = Path(output_file).resolve()
    output_file.parent.mkdir(parents=True, exist_ok=True)
    debug_log_path = (
        Path(debug_log_path)
        if debug_log_path
        else output_file.with_name(f"{output_file.stem}_template_debug.log")
    )
    logger = _DebugLogger(debug_log_path)
    logger.write("starting COM template build")
    logger.write(f"template_path={profile.template_path.resolve()}")
    logger.write(f"output_file={output_file}")
    logger.write(
        "profile: "
        f"title_layout={profile.title_layout_index} ({profile.title_layout_name}), "
        f"content_layout={profile.content_layout_index} ({profile.content_layout_name}), "
        f"title_slide={profile.title_slide_index}, content_slide={profile.content_slide_index}, "
        f"confidence={profile.confidence:.2f}, reason={profile.selection_reason}"
    )
    logger.write(f"plan title='{_preview(plan.get('title', ''))}'")
    logger.write(f"plan content slide count={len(plan.get('slides', []))}")

    pythoncom.CoInitialize()
    powerpoint = win32.DispatchEx("PowerPoint.Application")
    try:
        powerpoint.DisplayAlerts = 0
    except Exception:
        pass

    presentation = None
    try:
        presentation = powerpoint.Presentations.Open(str(profile.template_path.resolve()))
        logger.write(f"opened template in PowerPoint with {presentation.Slides.Count} slide(s)")
        original_slide_ids = [
            presentation.Slides(index).SlideID
            for index in range(1, presentation.Slides.Count + 1)
        ]
        title_source_id = presentation.Slides(profile.title_slide_index).SlideID
        content_source_id = presentation.Slides(profile.content_slide_index).SlideID
        logger.write(
            f"source ids: title_source_id={title_source_id}, "
            f"content_source_id={content_source_id}, original_slide_ids={original_slide_ids}"
        )

        keep_ids = []

        title_slide = _duplicate_to_end(presentation, title_source_id)
        keep_ids.append(title_slide.SlideID)
        logger.write(
            f"duplicated title sample slide {profile.title_slide_index} "
            f"to generated slide id={title_slide.SlideID}"
        )
        _replace_title_slide_text(title_slide, plan["title"], logger=logger)

        for slide_index, slide_data in enumerate(plan["slides"], start=1):
            content_slide = _duplicate_to_end(presentation, content_source_id)
            keep_ids.append(content_slide.SlideID)
            logger.write(
                f"duplicated content sample slide {profile.content_slide_index} "
                f"to generated content slide {slide_index} id={content_slide.SlideID}"
            )
            _replace_content_slide_text(
                content_slide,
                slide_data["title"],
                slide_data.get("bullets", []),
                logger=logger,
            )

        for slide_id in reversed(original_slide_ids):
            logger.write(f"deleting original template slide id={slide_id}")
            _slide_by_id(presentation, slide_id).Delete()

        logger.write(f"saving generated presentation to {output_file}")
        presentation.SaveAs(str(output_file))
    except Exception as exc:
        logger.write(f"COM template build failed: {type(exc).__name__}: {exc}")
        raise
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        try:
            powerpoint.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

    logger.write("writing speaker notes with python-pptx")
    _write_notes_with_python_pptx(output_file, plan)
    logger.write("COM template build complete")
    return output_file
