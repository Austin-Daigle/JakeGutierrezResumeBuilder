# Jake Gutierrez Resume Builder v1.0.2

JakeGutierrezResumeBuilder is a lightweight desktop app for building a professional resume in the **Jake Gutierrez / Jake’s Resume** format using a fast, structured, form-driven GUI.

Instead of fighting layout in a word processor, you edit **structured resume content** (Header + Sections + Entries + Bullets), preview it instantly, and export to production-ready formats (LaTeX Code, Microsoft Word, .Pdf, .Json).

Original template reference:
- https://www.overleaf.com/latex/templates/jakes-resume/syzfjbzwjncs

---

## Download
Select your version and download. It is recommended that you read the notes below (this is required for the macOS version).  
The source code for the Windows/Linux version and the macOS version are also available below for those who want to audit, reverse-engineer, or build their own version of the `.exe` or `.app`.

- **Windows (10+ recommended)**  
  - [Download v1.0.2 (.exe)](https://drive.google.com/file/d/1oQDn_C-4Tl8OqMC0SB6Xrhm7xRhjj2mc/view?usp=share_link)
- **macOS**  
  - [Download v1.0.2 (.app)](https://github.com/Austin-Daigle/JakeGutierrezResumeBuilder/blob/main/macOS/JakeGResumeBuilder%20v.1.0.2.app.zip)
- **Linux**  
  - [Download v1.0.2 (.py)](https://github.com/Austin-Daigle/JakeGutierrezResumeBuilder/blob/main/JakeGResumeBuilder_GUI_v.1.0.2.py)

**Download Execution Help Notes:**

**For Windows users:** This program may alert Windows Defender as being from an unknown developer.  
Select **“More info”** and then **“Run anyway.”**

**For macOS users:** Click on the link and select **“View raw”** or **“Download,”** then unzip the file.  
Apple Gatekeeper may display an error stating that “this file is damaged” and recommend moving it to the Trash.  
This happens because Apple follows a “walled garden” philosophy and requires developers to pay for and pass an audit to obtain a code-signing certificate that allows Gatekeeper to verify the app’s legitimacy.

**Fix for macOS:** Open your command terminal (you must have admin rights to perform this properly) and enter the following command:  
`xattr -d com.apple.quarantine <Program directory path here>`  
**Pro tip:** Dragging the file from the folder into the terminal automatically copies the directory path.  
This will remove the quarantine flag from the program, allowing the `.app` executable to run.

**For Linux:** Run the program from your command line or IDE.

**Source code:**

[Original Python code (as seen in the Windows/Linux version)](https://github.com/Austin-Daigle/JakeGutierrezResumeBuilder/blob/main/JakeGResumeBuilder_GUI_v.1.0.2.py)


[macOS optimized Python code](https://github.com/Austin-Daigle/JakeGutierrezResumeBuilder/blob/main/macOS/JakeGResumeBuilder_GUI_v.1.0.2_macOS.py)

---

## Table of Contents
- <a href="#download">Download</a>
- <a href="#screenshots">Screenshots</a>
- <a href="#key-features">Key Features</a>
- <a href="#quick-start">Quick Start</a>
- <a href="#detailed-usage-guide">Detailed Usage Guide</a>
  - <a href="#projects-json">Projects (.json)</a>
  - <a href="#header">Header</a>
  - <a href="#sections">Sections</a>
  - <a href="#entries">Entries</a>
  - <a href="#bullets">Bullets</a>
  - <a href="#rich-text-formatting">Rich Text Formatting</a>
  - <a href="#spellcheck">Spellcheck</a>
  - <a href="#undo--redo">Undo / Redo</a>
  - <a href="#preview">Preview</a>
  - <a href="#export">Export</a>
- <a href="#dependencies">Dependencies</a>
- <a href="#build-from-source">Build From Source</a>
  - <a href="#windows-exe">Windows (.exe)</a>
  - <a href="#macos-app">macOS (.app)</a>
- <a href="#version-history">Version History</a>
- <a href="#troubleshooting">Troubleshooting</a>

---

## Screenshots

Add screenshots / GIFs here.

---

## Key Features

### Core editing
- **Header editor** for contact info (name, phone, email, LinkedIn/GitHub links).
- **Sections list** with drag-and-drop reordering.
- **Entries list** per section with drag-and-drop reordering.
- **Entry editor dialogs** for editing structured fields.
- **Bullets editor** with drag-and-drop ordering.

### Rich text where it matters
- Rich text editing support for bullet/body content.
- **Section titles** support rich formatting while remaining backward-compatible with plain-text projects.

### Productivity
- **Undo / Redo** for common edits and reorders.
- **Save + Save As** project workflow.
- **Preview** your resume before exporting.
- Helpful UI tooltips across common actions.

### Spellcheck (optional)
- Red underline spellcheck for supported text inputs.
- Right-click suggestions.
- **Ignore All** per project.
- **Edit Document Spellcheck Data...** to manage ignored words.

### Export formats
- **LaTeX (.tex)**
- **Word (.docx)**
- **PDF (.pdf)**

---

## Quick Start

1. **Create a project**
   - `File -> Save project as .json...`

2. **Fill out Header**
   - Enter contact details and links.

3. **Add / reorder Sections**
   - Add sections (Education, Experience, Projects, Skills, etc.).
   - Drag sections to reorder.

4. **Add / reorder Entries**
   - Select a section.
   - Add entries under it.
   - Drag entries to reorder.

5. **Add bullets (Experience / Projects)**
   - Open an entry editor.
   - Add bullets and drag to reorder.

6. **Preview**
   - Click `Preview` to verify layout and consistency.

7. **Export**
   - `File -> Export LaTeX / Word / PDF`

---

## Detailed Usage Guide

### Projects (.json)

A project is a single `.json` file containing your entire resume state.

- **Save** writes your current in-memory state to disk.
- **Load** replaces your current in-memory state with the selected file.

Recommended workflow:
- Create a project file first so `Ctrl+S` works immediately.
- Save often.

---

### Header

The Header contains top-of-resume contact info (name, phone, email, and link fields). Header changes update internal state and are included in Preview/Export.

---

### Sections

A section is a top-level resume block (e.g., Education, Experience, Projects, Skills).

Key behaviors:
- Section order controls output order.
- Drag-and-drop reorders sections instantly.
- Deleting a section removes its entries from the project.

---

### Entries

Entries belong to a section (e.g., a job within Experience, a school within Education).

Key behaviors:
- Entry order controls output order.
- Fields vary by section type.
- Editing happens in a focused editor dialog.

---

### Bullets

Bullets are primarily used in Experience and Projects.

Key behaviors:
- Bullet order matters.
- Drag-and-drop reordering is supported.
- Rich text formatting is supported where applicable.

---

### Rich Text Formatting

Rich text formatting is stored in the project file as structured segments (not as raw HTML).

Supported formatting (where available):
- Bold
- Italic
- Underline
- Strikethrough
- Text color
- Highlight color

---

### Spellcheck

Spellcheck is designed to feel like a modern editor:

- Misspelled words are underlined in bright red.
- Right-click a misspelled word to:
  - choose a suggestion
  - choose **Ignore All** to ignore that word in the current project

Ignored words are stored per-project in the project JSON.

---

### Undo / Redo

Undo/Redo supports rapid iteration while editing:
- `Edit -> Undo`
- `Edit -> Redo`

---

### Preview

Preview renders from the internal structured state (not directly from UI widgets). It’s the recommended checkpoint before export.

---

### Export

Exports are generated from the same structured state used for Preview:

- **LaTeX (.tex)**: best for typography and version control.
- **Word (.docx)**: useful for recruiters who prefer Word.
- **PDF (.pdf)**: ideal for final submission.

---

## Dependencies

This app is a single-file Tkinter GUI script.

Optional feature dependencies:
- Spellcheck: `pyspellchecker`
- Word export: `python-docx`
- PDF export: `reportlab`

If an optional dependency is missing, the app shows an error dialog with the exact install command.

---

## Build From Source

### Windows (.exe)

This repo includes a PyInstaller spec file for Windows:

```bash
python -m PyInstaller --noconfirm --clean ResumeTemplateMaker.spec
```

Output:
- `dist/JakeGResumeBuilder v.1.0.2.exe`

---

### macOS (.app)

The macOS-specific app script and build tooling are located in `macOSAppDev/`.

```bash
chmod +x macOSAppDev/build_macos_app.sh
./macOSAppDev/build_macos_app.sh
```

Output:
- `macOSAppDev/dist/JakeGResumeBuilder v.1.0.2.app`

Note: macOS Gatekeeper may warn on unsigned apps. For broad distribution you typically need code signing and notarization.

---

## Version History

### v1.0.2
- Rich-text section title editing (with bold-by-default rendering).
- Strikethrough support across preview and export formats.
- Expanded tooltips across the UI for discoverability.
- macOS-specific variant (`macOSAppDev/`) to address startup layout behavior.

### v1.0.1
- Stabilized Windows packaging and core editing workflow.
- Export support for LaTeX / Word / PDF.
- Spellcheck support (optional dependency).

### v1.0.0
- Initial public release.
- Header / Sections / Entries / Bullets editing with preview/export workflow.

---

## Troubleshooting

### Spellcheck doesn’t work
- Install `pyspellchecker`.
- Enable it via `Settings -> Enable Spellcheck`.

### Export fails
- Word export requires `python-docx`.
- PDF export requires `reportlab`.
- The error dialog will tell you what’s missing.

### Ignored words disappeared
- Ignored words are stored inside the project JSON.
- Save your project after editing spellcheck data.
