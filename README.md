# arXiv Flask App

A Flask web app for searching arXiv papers, processing PDF content, and generating Word documents.

## Features

- Search arXiv papers from a browser interface.
- Download and parse paper PDFs.
- Generate formatted Word documents from selected paper content.
- Optional local translation support through an Ollama-compatible API.

## Requirements

- Python 3.10+
- Dependencies listed in `requirements.txt`

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
python run_app.py
```

Then open `http://127.0.0.1:5000`.

## Optional Configuration

The app reads these environment variables when translation features are used:

- `OLLAMA_API_URL`, default: `http://127.0.0.1:11434/api/generate`
- `OLLAMA_MODEL`, default: `deepseek-r1:8b`
