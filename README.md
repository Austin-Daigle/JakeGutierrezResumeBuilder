# JakeGutierrezResumeBuilder

This is a Python (Tkinter) desktop application that makes building professional resumes in the **Jake Gutierrez format** easy with a fast, structured GUI.

Original template reference:
- https://www.overleaf.com/latex/templates/jakes-resume/syzfjbzwjncs

---

## Download

> TODO: Add official download links here.

- [Download vX.Y.Z (Windows 10+)](TODO)
- [Download vX.Y.Z (macOS)](TODO)

---

## Screenshots

> TODO: Add images / GIFs here.

---

## Table of Contents

- [Key Features](#key-features)
- [Quick Start (Workflow)](#quick-start-workflow)
- [How the App Thinks (Data Model)](#how-the-app-thinks-data-model)
- [Detailed Usage Guide](#detailed-usage-guide)
  - [Projects (.json files)](#projects-json-files)
  - [Header](#header)
  - [Sections](#sections)
  - [Entries](#entries)
  - [Bullets](#bullets)
  - [Rich Text Editing](#rich-text-editing)
  - [Spellcheck](#spellcheck)
  - [Undo / Redo](#undo--redo)
  - [Preview](#preview)
  - [Export](#export)
- [Building & Packaging](#building--packaging)
  - [Windows: Build a .exe](#windows-build-a-exe)
  - [macOS: Build a .app](#macos-build-a-app)
- [Dependencies](#dependencies)
- [Troubleshooting](#troubleshooting)
- [Version History](#version-history)

---

## Key Features

### Core editing

- **Header editor** for contact info.
- **Sections list** (drag-and-drop reorder).
- **Entries list** per section (drag-and-drop reorder).
- **Entry editor dialogs** for editing a selected entry’s fields.
- **Bullet editor** with drag-and-drop reordering.

### Productivity

- **Undo / Redo** support for common edits and reorders.
- **Ctrl+S Save** for projects.
- **Tooltips** across the UI.
- **Preview** before exporting.

### Rich text support

- Rich text formatting where supported:
  - **bold**
  - *italic*
  - underline
  - strikethrough
  - text color
  - highlight color

### Spellcheck

- Optional spellcheck with:
  - red underline on misspellings
  - right-click suggestions
  - per-project **Ignore All** list
  - a manager dialog to edit ignored words

### Output formats

- **Export LaTeX (.tex)**
- **Export Word (.docx)**
- **Export PDF (.pdf)**

---

## Quick Start (Workflow)

1. **Create a project**
   - `File -> Save project as .json...`
2. **Fill out Header**
   - Add contact info in the Header panel.
3. **Create sections**
   - Add sections (Education, Experience, Projects, Skills, etc.).
   - Drag sections to reorder.
4. **Add entries**
   - Select a section.
   - Add entries under it.
   - Drag entries to reorder.
5. **Add bullets (Experience / Projects)**
   - Open an entry editor.
   - Add bullets and reorder bullets by dragging.
6. **Spellcheck (optional)**
   - `Settings -> Enable Spellcheck`
   - Right-click underlined words for suggestions.
   - Use **Ignore All** for words you want to ignore in this project.
   - Manage ignored words: `Settings -> Edit Document Spellcheck Data...`
7. **Preview**
   - Click `Preview` to verify ordering, spacing, and formatting.
8. **Export**
   - `File -> Export LaTeX / Word / PDF`

---

## How the App Thinks (Data Model)

The application maintains an internal in-memory state that mirrors what gets saved to disk as a JSON project file.

### High-level structure

- **Header**: contact fields (name, email, links, etc.)
- **Sections**: ordered list of sections
- Each **Section** contains:
  - metadata (title, kind)
  - ordered list of **Entries**
- Each **Entry** contains:
  - structured fields (depending on section kind)
  - bullets/body content (depending on kind)

### Why this matters

- **Preview and Export do not read the UI directly**.
- They render from the underlying structured state.
- Saving writes this structured state to your `.json` project file.

---

## Detailed Usage Guide

### Projects (.json files)

A “project” is a JSON file containing your entire resume state.

- **Save** writes current in-memory state to disk.
- **Load** replaces current in-memory state with the file contents.
- The app also stores project-specific spellcheck data in the JSON.

Recommended workflow:
- Create a project file first (so Ctrl+S works immediately).
- Save frequently.

---

### Header

The Header area contains your contact details (name, phone, email, LinkedIn, GitHub, etc.).

Behavior:
- Header changes update the internal state.
- Save persists them to the project JSON.

---

### Sections

A section is a top-level resume block (e.g., Education, Experience, Projects, Skills).

Key behaviors:
- The **order of sections** is the order used in Preview and Export.
- You can **drag-and-drop** sections to reorder.
- Deleting a section removes its entries from the project.

Section title editing:
- Section titles can be edited in the UI.
- Section titles support rich text (bold by default).

---

### Entries

Entries belong to a section (e.g., one job inside Experience, one school inside Education).

Key behaviors:
- The **order of entries** controls output order.
- Entry fields vary depending on section kind.
- Entry editing happens in an editor dialog.

---

### Bullets

Bullets are typically used for Experience and Projects.

Key behaviors:
- Bullets are managed as a list.
- Bullet order matters and is reflected in Preview and Export.
- Bullets support rich text editing where applicable.

---

### Rich Text Editing

Some fields use a rich text editor that supports formatting.

Supported formatting options:
- bold / italic / underline / strikethrough
- text color
- highlight color

Rich text is preserved across:
- Preview
- LaTeX export
- Word export
- PDF export

---

### Spellcheck

Spellcheck is designed to act like a modern editor:

- Misspelled words are underlined in red.
- Right-click a misspelled word:
  - choose a suggestion to replace it
  - or choose **Ignore All** to ignore that word in this project

Ignore list persistence:
- Ignored words are stored per project (inside the project JSON).

Edit Document Spellcheck Data:
- `Settings -> Edit Document Spellcheck Data...` opens a manager dialog where you can add/remove ignored words.

---

### Undo / Redo

Undo/Redo exists to support rapid iteration while editing:

- `Edit -> Undo`
- `Edit -> Redo`

Some changes (like typing) may be grouped for usability.

---

### Preview

Preview renders the internal data state into a display view so you can confirm:

- ordering
- spacing
- consistency
- formatting outcome

Preview is the recommended checkpoint before export.

---

### Export

The app supports multiple export formats:

- **LaTeX (.tex)**: good for version control and high-quality typography.
- **Word (.docx)**: good for recruiter workflows that prefer Word.
- **PDF (.pdf)**: good for final submission.

Exports are based on the internal structured state, not the raw UI.

---

## Building & Packaging

> This section is for developers who want to produce distributable builds.

### Windows: Build a .exe

This repo includes a PyInstaller spec file:

- `ResumeTemplateMaker.spec`

Build command:

```bash
python -m PyInstaller --noconfirm --clean ResumeTemplateMaker.spec
```

Output:
- `dist/JakeGResumeBuilder v.1.0.2.exe`

---

### macOS: Build a .app

The macOS development folder contains the macOS-specific entry script and build helpers:

- `macOSAppDev/JakeGResumeBuilder_GUI_v.1.0.2_macOS.py`
- `macOSAppDev/ResumeTemplateMaker_macOS.spec`
- `macOSAppDev/build_macos_app.sh`

Build command on a Mac:

```bash
chmod +x macOSAppDev/build_macos_app.sh
./macOSAppDev/build_macos_app.sh
```

Output:
- `macOSAppDev/dist/JakeGResumeBuilder v.1.0.2.app`

Note:
- macOS Gatekeeper may warn for unsigned apps. For wider distribution, use code signing + notarization.

---

## Dependencies

Core:
- Python 3
- Tkinter (ships with most Python distributions)

Optional feature dependencies (required for specific features when running from source):

- Spellcheck: `pyspellchecker`
  - Install: `pip install pyspellchecker`
- Word export: `python-docx`
  - Install: `pip install python-docx`
- PDF export: `reportlab`
  - Install: `pip install reportlab`

If a dependency is missing, the app shows an error message with the exact install command.

---

## Troubleshooting

### Spellcheck doesn’t work

- Ensure `pyspellchecker` is installed.
- Enable it via `Settings -> Enable Spellcheck`.

### Export fails

- Word export requires `python-docx`.
- PDF export requires `reportlab`.
- The error dialog will tell you exactly what’s missing.

### My ignored words disappeared

- Ignored words are stored in the project JSON.
- Make sure you saved your project after editing the ignore list.

---

## Version History

> You can expand this section as releases are published.

### v1.0.2

- Rich section title editing (formatted titles stored and rendered in preview/exports)
- Rich text improvements (including strikethrough support)
- UI tooltips added across the app
- macOS-specific variant added to address initial layout squishing

### v1.0.1

- TODO

---
