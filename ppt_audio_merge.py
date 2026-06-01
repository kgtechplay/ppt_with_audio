"""Embed per-slide MP3 files into a PowerPoint deck using comtypes."""

import argparse
import re
import sys
from pathlib import Path

import comtypes.client


def output_path_for_ppt(ppt_path: Path) -> Path:
    """Same folder as input; name: '<stem> + audio.pptx'."""
    return ppt_path.with_name(f"{ppt_path.stem} + audio.pptx")


SLIDE_AUDIO_RE = re.compile(r"^slide_(\d+)\.mp3$", re.IGNORECASE)


def build_audio_index(audio_dir: Path) -> dict[int, Path]:
    """Map slide number -> MP3 for every slide_*.mp3 in the folder."""
    index: dict[int, Path] = {}
    for path in sorted(audio_dir.glob("*.mp3")):
        match = SLIDE_AUDIO_RE.match(path.name)
        if not match:
            continue
        slide_number = int(match.group(1))
        if slide_number in index:
            print(
                f"Warning: duplicate audio for slide {slide_number}, "
                f"keeping {index[slide_number].name}, ignoring {path.name}"
            )
            continue
        index[slide_number] = path.resolve()
    return index


def merge_audio_into_ppt(
    ppt_path: Path,
    audio_dir: Path,
    output_path: Path | None = None,
    play_on_entry: bool = True,
) -> Path:
    ppt_path = ppt_path.resolve()
    audio_dir = audio_dir.resolve()

    if not ppt_path.is_file():
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")
    if not audio_dir.is_dir():
        raise NotADirectoryError(f"Audio folder not found: {audio_dir}")

    output_path = (output_path or output_path_for_ppt(ppt_path)).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    audio_by_slide = build_audio_index(audio_dir)
    if not audio_by_slide:
        raise RuntimeError(
            f"No slide_*.mp3 files found in {audio_dir}. "
            "Expected names like slide_001.mp3, slide_02.mp3, etc."
        )
    print(f"Found {len(audio_by_slide)} audio file(s) in {audio_dir}")

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    # Do not set Visible=False; Office 365 often blocks hiding the window.
    presentation = None
    embedded = 0
    missing = []

    try:
        presentation = powerpoint.Presentations.Open(str(ppt_path))
        slide_count = presentation.Slides.Count
        for slide_number in range(1, slide_count + 1):
            audio_file = audio_by_slide.get(slide_number)
            if audio_file is None:
                missing.append(slide_number)
                continue

            slide = presentation.Slides(slide_number)
            shape = slide.Shapes.AddMediaObject2(
                FileName=str(audio_file),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=20,
                Top=20,
                Width=40,
                Height=40,
            )
            shape.AnimationSettings.PlaySettings.PlayOnEntry = play_on_entry
            shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = True
            embedded += 1
            print(f"Slide {slide_number}: embedded {audio_file.name}")

        presentation.SaveAs(str(output_path))
    finally:
        if presentation is not None:
            presentation.Close()
        powerpoint.Quit()

    if missing:
        print(f"Warning: no audio for slide(s): {', '.join(map(str, missing))}")
    if embedded == 0:
        raise RuntimeError(
            f"No audio files matched in {audio_dir}. "
            "Expected names like slide_001.mp3 or slide_01.mp3."
        )

    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Embed MP3 files from a folder into each slide of a PowerPoint file."
    )
    parser.add_argument("ppt_path", help="Input .pptx file")
    parser.add_argument("audio_dir", help="Folder containing slide_*.mp3 files")
    parser.add_argument(
        "--output",
        "-o",
        help='Output .pptx (default: same folder as input, "<name> + audio.pptx")',
    )
    parser.add_argument(
        "--no-play-on-entry",
        action="store_true",
        help="Do not auto-play audio when the slide appears",
    )
    args = parser.parse_args()

    ppt_path = Path(args.ppt_path)
    audio_dir = Path(args.audio_dir)
    output_path = Path(args.output) if args.output else None

    try:
        result = merge_audio_into_ppt(
            ppt_path=ppt_path,
            audio_dir=audio_dir,
            output_path=output_path,
            play_on_entry=not args.no_play_on_entry,
        )
    except (FileNotFoundError, NotADirectoryError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Created: {result}")


if __name__ == "__main__":
    main()
