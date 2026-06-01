import json
import sys
from pathlib import Path


def slide_to_markdown(slide):
    slide_type = slide.get("slide_type", "content")

    title = slide.get("title", "")
    subtitle = slide.get("subtitle", "")

    bullets = slide.get("bullets", [])

    notes = slide.get("speaker_notes", "")

    background = slide.get("background_style", "")

    image = slide.get("image", "")

    md = []

    md.append("---")

    if slide_type == "cover":
        md.append("layout: cover")

    elif slide_type == "image-right":
        md.append("layout: image-right")

        if image:
            md.append(f"image: {image}")

    elif slide_type == "image-left":
        md.append("layout: image-left")

        if image:
            md.append(f"image: {image}")

    elif slide_type == "section":
        md.append("layout: center")

    elif slide_type == "quote":
        md.append("layout: center")

    if background:
        md.append(f"background: {background}")

    md.append("---")
    md.append("")

    md.append(f"# {title}")
    md.append("")

    if subtitle:
        md.append(subtitle)
        md.append("")

    for bullet in bullets:
        md.append(f"- {bullet}")

    md.append("")
    md.append("---")
    md.append("")

    md.append(notes)

    md.append("")

    return "\n".join(md)


def build_slidev(data):
    output = []

    output.append("---")
    output.append(f'title: "{data["presentation_title"]}"')
    output.append("theme: seriph")
    output.append("drawings: false")
    output.append("transition: fade")
    output.append("---")
    output.append("")

    for slide in data["slides"]:
        output.append(slide_to_markdown(slide))
        output.append("")

    return "\n".join(output)


def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("python create_slidev_md.py slides.json")
        sys.exit(1)

    json_file = Path(sys.argv[1])
    if not json_file.exists():
        print(f"File not found: {json_file}")
        sys.exit(1)

    with open(json_file, "r", encoding="utf-8") as f:
        slide_data = json.load(f)

    markdown = build_slidev(slide_data)

    output_file = json_file.with_suffix(".md")

    with open(output_file, "w", encoding="utf-8") as f:
        f.write(markdown)

    print(f"Created: {output_file.resolve()}")


if __name__ == "__main__":
    main()
