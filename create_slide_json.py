import json
import sys
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()
client = OpenAI()

SYSTEM_PROMPT = """
You are an expert presentation designer and content strategist.

Create a presentation slide plan as structured JSON.

Return ONLY valid JSON.
Do not include markdown.
Do not include explanations.
"""

SLIDE_JSON_SCHEMA = {
    "presentation_title": "string",
    "presentation_subtitle": "string",
    "audience": "string",
    "tone": "string",
    "slides": [
        {
            "slide_number": 1,
            "slide_type": "cover | agenda | content | image-right | image-left | quote | section | summary | closing",
            "title": "string",
            "subtitle": "string",
            "bullets": ["string"],
            "visual_prompt": "string",
            "background_style": "string",
            "speaker_notes": "string"
        }
    ]
}


def load_text_file(file_path: str) -> str:
    return Path(file_path).read_text(encoding="utf-8")


def generate_slide_json(content: str, number_of_slides: int = 8) -> dict:
    prompt = f"""
Create a polished presentation structure from the content below.

Required output JSON format:
{json.dumps(SLIDE_JSON_SCHEMA, indent=2)}

Instructions:
- Create exactly {number_of_slides} slides.
- Make the presentation executive-friendly.
- Use clear slide titles.
- Use 3 to 5 bullets per normal content slide.
- Use fewer bullets for cover, section, quote, and closing slides.
- Add a visual_prompt for each slide that can be used later for image generation.
- Add a background_style for each slide.
- Add speaker_notes of 80 to 130 words per slide.
- Keep the content faithful to the source material.
- Do not invent hard facts, statistics, or names not present in the content.

Source content:
{content}
"""

    response = client.responses.create(
        model="gpt-4.1-mini",
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
    )

    raw_json = response.output_text.strip()
    return json.loads(raw_json)


def save_json(data: dict, output_file: str):
    Path(output_file).write_text(
        json.dumps(data, indent=2, ensure_ascii=False),
        encoding="utf-8"
    )


def main():
    if len(sys.argv) < 2:
        print("Usage:")
        print("python create_slide_json.py content.txt")
        print("python create_slide_json.py content.txt 10")
        sys.exit(1)

    input_file = sys.argv[1]

    number_of_slides = 8
    if len(sys.argv) >= 3:
        number_of_slides = int(sys.argv[2])

    content = load_text_file(input_file)

    slide_json = generate_slide_json(
        content=content,
        number_of_slides=number_of_slides
    )

    output_file = Path(input_file).with_name(
        f"{Path(input_file).stem}_slides.json"
    )

    save_json(slide_json, output_file)

    print(f"Saved slide JSON: {output_file}")


if __name__ == "__main__":
    main()