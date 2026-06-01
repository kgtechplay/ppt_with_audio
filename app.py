import os
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

from create_presentation import run_create_presentation
from elevenlabs_voice_to_ppt import run_elevenlabs_audio
from ppt_audio_merge import merge_audio_into_ppt
from speaker_notes import run_speaker_notes

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"

AUDIO_SUBDIR = "Audio"


def job_dir_for_source(source_filename: str) -> Path:
    """output/<input-file-stem>/"""
    job_dir = (OUTPUT_DIR / Path(source_filename).stem).resolve()
    job_dir.mkdir(parents=True, exist_ok=True)
    return job_dir


def save_upload(uploaded_file, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    target_path = target_dir / uploaded_file.name
    target_path.write_bytes(uploaded_file.getbuffer())
    return target_path


def make_progress_handler(progress_bar, status_text, step_offset=0, step_weight=1.0):
    def handler(step, total, message):
        overall = step_offset + (step / total) * step_weight
        progress_bar.progress(min(overall, 1.0))
        status_text.write(message)

    return handler


def render_download(path: Path, label: str):
    if path.exists():
        st.download_button(
            label=label,
            data=path.read_bytes(),
            file_name=path.name,
            mime="application/octet-stream",
        )


def run_content_pipeline(
    content_path: Path,
    topic: str,
    num_slides: int,
    progress_bar,
    status,
) -> dict[str, Path]:
    """Content .txt -> PPT with notes -> Audio/ -> merged PPT."""
    job_dir = job_dir_for_source(content_path.name)
    ppt_path = job_dir / f"{content_path.stem}.pptx"
    audio_dir = job_dir / AUDIO_SUBDIR

    status.write("Step 1/3: Creating presentation with speaker notes")
    progress_handler = make_progress_handler(progress_bar, status, 0.0, 1 / 3)

    notes_path = run_create_presentation(
        content_path=content_path,
        output_path=ppt_path,
        topic=topic,
        num_slides=num_slides,
        progress_callback=progress_handler,
    )
    status.write(f"Saved: `{notes_path.relative_to(BASE_DIR)}`")

    status.write("Step 2/3: Generating narration (ElevenLabs)")
    progress_handler = make_progress_handler(progress_bar, status, 1 / 3, 1 / 3)

    run_elevenlabs_audio(
        input_pptx=notes_path,
        audio_dir=audio_dir,
        progress_callback=progress_handler,
    )
    status.write(f"Saved audio: `{audio_dir.relative_to(BASE_DIR)}/`")

    status.write("Step 3/3: Merging audio into PowerPoint")
    progress_handler = make_progress_handler(progress_bar, status, 2 / 3, 1 / 3)

    merged_name = f"{ppt_path.stem} + audio.pptx"
    merged_path = job_dir / merged_name

    merge_audio_into_ppt(
        ppt_path=notes_path,
        audio_dir=audio_dir,
        output_path=merged_path,
    )
    status.write(f"Saved: `{merged_path.relative_to(BASE_DIR)}`")

    return {
        "job_dir": job_dir,
        "presentation": notes_path,
        "audio_dir": audio_dir,
        "merged": merged_path,
    }


def run_existing_ppt_pipeline(
    ppt_path: Path,
    context_path: Path | None,
    context_text: str | None,
    progress_bar,
    status,
) -> dict[str, Path]:
    """Existing PPT -> notes -> Audio/ -> merged PPT (same folder layout)."""
    job_dir = job_dir_for_source(ppt_path.name)
    notes_path = job_dir / f"{ppt_path.stem}_with_notes.pptx"
    audio_dir = job_dir / AUDIO_SUBDIR

    status.write("Step 1/3: Generating speaker notes")
    progress_handler = make_progress_handler(progress_bar, status, 0.0, 1 / 3)

    notes_path = run_speaker_notes(
        input_path=ppt_path,
        context_path=context_path,
        context_text=context_text,
        output_path=notes_path,
        progress_callback=progress_handler,
    )
    status.write(f"Saved: `{notes_path.relative_to(BASE_DIR)}`")

    status.write("Step 2/3: Generating narration (ElevenLabs)")
    progress_handler = make_progress_handler(progress_bar, status, 1 / 3, 1 / 3)

    run_elevenlabs_audio(
        input_pptx=notes_path,
        audio_dir=audio_dir,
        progress_callback=progress_handler,
    )
    status.write(f"Saved audio: `{audio_dir.relative_to(BASE_DIR)}/`")

    status.write("Step 3/3: Merging audio into PowerPoint")
    progress_handler = make_progress_handler(progress_bar, status, 2 / 3, 1 / 3)

    merged_path = job_dir / f"{notes_path.stem} + audio.pptx"
    merge_audio_into_ppt(
        ppt_path=notes_path,
        audio_dir=audio_dir,
        output_path=merged_path,
    )
    status.write(f"Saved: `{merged_path.relative_to(BASE_DIR)}`")

    return {
        "job_dir": job_dir,
        "presentation": notes_path,
        "audio_dir": audio_dir,
        "merged": merged_path,
    }


st.set_page_config(
    page_title="Voice PPT Generator",
    page_icon="🎙️",
    layout="centered",
)

st.title("Voice PPT Generator")
st.caption(
    "From content or PowerPoint → speaker notes → ElevenLabs audio → merged deck. "
    f"Outputs go to `{OUTPUT_DIR.name}/<input-file-name>/`."
)

mode = st.radio(
    "What do you have?",
    [
        "Content file (text)",
        "PowerPoint without speaker notes",
    ],
)

st.divider()

ppt_upload = None
content_upload = None
context_upload = None

if mode == "Content file (text)":
    st.subheader("Content")
    content_upload = st.file_uploader(
        "Content file (.txt)",
        type=["txt"],
    )
    topic = st.text_input(
        "Presentation topic",
        value="Business Presentation",
    )
    num_slides = st.number_input(
        "Number of slides",
        min_value=2,
        max_value=30,
        value=8,
        help="Includes 1 title slide plus the remaining content slides.",
    )
    context_text = ""
else:
    st.subheader("Upload your presentation")
    ppt_upload = st.file_uploader(
        "PowerPoint file (.pptx)",
        type=["pptx"],
    )
    st.subheader("Optional context for speaker notes")
    context_upload = st.file_uploader(
        "Context file (.txt)",
        type=["txt"],
    )
    context_text = st.text_area(
        "Or paste context here",
        placeholder="Audience, goals, tone...",
        height=120,
    )
    topic = ""
    num_slides = 0

generate = st.button("Generate presentation with audio", type="primary")

if generate:
    if mode == "Content file (text)" and not content_upload:
        st.error("Please upload a content file.")
        st.stop()

    if mode == "PowerPoint without speaker notes" and not ppt_upload:
        st.error("Please upload a PowerPoint file.")
        st.stop()

    progress_bar = st.progress(0)
    status_text = st.empty()
    log_area = st.container()
    job_dir = None

    try:
        with st.status("Running pipeline...", expanded=True) as status:
            if mode == "Content file (text)":
                job_dir = job_dir_for_source(content_upload.name)
                content_path = save_upload(content_upload, job_dir)
                results = run_content_pipeline(
                    content_path=content_path,
                    topic=topic,
                    num_slides=int(num_slides),
                    progress_bar=progress_bar,
                    status=status,
                )
            else:
                job_dir = job_dir_for_source(ppt_upload.name)
                ppt_path = save_upload(ppt_upload, job_dir)
                context_path = None
                if context_upload:
                    context_path = save_upload(context_upload, job_dir)
                context_value = context_text.strip() or None

                results = run_existing_ppt_pipeline(
                    ppt_path=ppt_path,
                    context_path=context_path,
                    context_text=context_value,
                    progress_bar=progress_bar,
                    status=status,
                )

            status.update(label="Pipeline complete", state="complete")

        progress_bar.progress(1.0)
        status_text.success("All done.")

        st.subheader("Generated files")
        col1, col2 = st.columns(2)

        with col1:
            render_download(
                results["presentation"],
                "Download presentation with speaker notes",
            )

        with col2:
            render_download(
                results["merged"],
                "Download narrated presentation",
            )

        audio_dir = results["audio_dir"]
        audio_paths = sorted(audio_dir.glob("*.mp3")) if audio_dir.exists() else []
        if audio_paths:
            with st.expander(f"Audio files ({len(audio_paths)})"):
                for audio_path in audio_paths:
                    st.audio(audio_path.read_bytes(), format="audio/mp3")
                    render_download(audio_path, f"Download {audio_path.name}")

        with log_area:
            st.markdown("**Output folder**")
            st.code(str(results["job_dir"]))
            st.caption(
                f"Contains the .pptx with speaker notes, "
                f"`{AUDIO_SUBDIR}/slide_*.mp3`, and the merged `* + audio.pptx`."
            )

    except Exception as exc:
        progress_bar.empty()
        status_text.error(f"Something went wrong: {exc}")
        if job_dir is not None and job_dir.exists():
            st.warning(f"Partial outputs may be in: `{job_dir}`")
        st.stop()

st.divider()
st.markdown(
    "Ensure `OPENAI_API_KEY`, `ELEVENLABS_API_KEY`, and `ELEVENLABS_VOICE_ID` "
    "are set in your `.env` file. Merging requires PowerPoint on Windows (`comtypes`)."
)
