# Bennet PDF Manager

Inspired by the sacred knowledge trove of Mr.Bennet's library from Jane Austin's _Pride and Prejudice_, Bennet is a desktop PDF manager built with PyQt6 and [pypdf](https://pypi.org/project/pypdf/). It scans a chosen "home" directory for PDF files, indexes them, and lets you quickly search, preview, and edit their metadata.

## Features

- Recursive scan of a home directory to index all PDF files.
- Search across title, author, year, subject, notes, and indexed contents.
- Heuristic extraction of publication year and author names from PDF metadata and first-page text.
- Inline metadata editor for title, author, subject, year, and personal notes.
- Writes updated metadata back into the PDF (including an embedded `BennetNotes` field).
- Cached index (`.bennet_pdf_cache.json`) for fast startup and rescans.
- Text preview of the first few pages of each PDF.
- Open PDFs in the system default viewer or reveal them in the file manager.

## Quick Start

Within bin/dist you can find a pre-compiled Bennet.exe. If you are running 64-bit Windows, it should just work. Otherwise, you can run the Python or build the executable yourself with the instructions provided below.

On first launch, click **"Choose Home…"** to select the directory that contains your PDFs. Bennet will index all PDFs under that directory and build a searchable table of records. You can:

- Filter the list by typing into the search box.
- Click a row to view and edit metadata.
- Click **Save** to write updated metadata back into the PDF.
- Double-click a row or use the context menu to open the PDF or show it in your file manager.

## Python Usage

Requirements:

- Python 3.9+
- [PyQt6](https://pypi.org/project/PyQt6/)
- [pypdf](https://pypi.org/project/pypdf/)

Install dependencies:

```bash
pip install PyQt6 pypdf
```

Running the application

```bash
python bennet.py
```

## Building a Windows executable

You can build a standalone `.exe` using PyInstaller:

```bash
python -m PyInstaller --name "Bennet" --windowed --onefile bennet.py
```

This will produce a self-contained executable named `Bennet.exe` that you can run on Windows without a separate Python installation.

## Windows Smart App Control Issue

If the .exe is blocked from running by Windows Smart App Control and there is no way to "Run Anyway", there is no other alternative than to either turn off Windows Smart App Control or to download and run the Python/Build the executable yourself.
