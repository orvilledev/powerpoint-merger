# PowerPoint Merger

A Streamlit web application for merging multiple PowerPoint presentations while applying consistent formatting.

## Features

- Merge multiple PowerPoint (.pptx) files into a single presentation
- Extract text content from slides
- Apply consistent formatting:
  - Black background for all slides
  - Arial font, 72pt size
  - Centered text (horizontal and vertical)
  - Yellow text with bold for all-caps slides
  - White text for other slides
- Preserves original text case

## Requirements

- Python 3.7+
- streamlit
- python-pptx

## Installation

```bash
pip install streamlit python-pptx
```

## Usage

```bash
streamlit run app.py
```

Then open your browser to `http://localhost:8501` and:
1. Upload one or more PowerPoint files
2. Click "Merge PowerPoints"
3. Download the merged presentation

## How It Works

The app extracts text content from each slide in the uploaded presentations and creates new slides with:
- Standardized formatting (72pt Arial, centered)
- Black backgrounds
- Color coding based on text case (yellow for all-caps, white for others)

