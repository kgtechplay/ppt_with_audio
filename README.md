# ppt_with_audio

Generate PowerPoint presentations with AI speaker notes, ElevenLabs narration, and embedded slide audio.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
```

Add your API keys to `.env` (see [.env.example](.env.example)).

## Web UI

```powershell
streamlit run app.py
```

Outputs are written to `output/<input-file-name>/` (presentation, `Audio/` MP3s, and merged `* + audio.pptx`).

The web UI can optionally accept a PowerPoint template. Template mode detects reusable title and content sample slides, duplicates those slides through PowerPoint COM, replaces mapped text, and writes template detection/debug logs under the job output folder. When slide scoring is ambiguous, the detector uses OpenAI to choose the best reusable sample slide and falls back to heuristic scoring if the AI call fails.

## CLI pipeline

```powershell
python create_presentation.py content.txt 8
python elevenlabs_voice_to_ppt.py "output/content/content.pptx" --audio-only --audio-dir "output/content/Audio"
python ppt_audio_merge.py "output/content/content.pptx" "output/content/Audio"
```

## Requirements

- Windows + Microsoft PowerPoint (for `ppt_audio_merge.py` via COM)
- Windows + Microsoft PowerPoint for template-based deck generation via COM
- OpenAI and ElevenLabs API keys
