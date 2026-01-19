# JakeGutierrezResumeBuilder
This is a Python script that makes building professional resumes in the Jake Gutierrez format easy with a simple GUI.
Here is the link to the original page and formatting:
[https://www.overleaf.com/latex/templates/jakes-resume/syzfjbzwjncs](https://www.overleaf.com/latex/templates/jakes-resume/syzfjbzwjncs)

[Download Here: (Windows 10+)](https://drive.google.com/file/d/1JFJU74b_RAGoFmr8yhBAooWwMuXCD5mS/view?usp=sharing)


[Download Here: (Python Script)](https://github.com/Austin-Daigle/JakeGutierrezResumeBuilder/blob/main/JakeCResumeBuilder_GUI_v.1.0.py)


<img width="550" height="350" alt="image" src="https://github.com/user-attachments/assets/fcabd604-c84d-4750-95e5-62d36be7a35c" />



# Download Link (.exe):
[Resume Template Maker.exe (Windows Only)](https://drive.google.com/file/d/1tvza5xIbJ_dcYEn6eF16h2ZAI7yxP-6T/view?usp=sharing)

# Resume Template Maker (Tkinter)
A lightweight desktop application for building a resume from structured data (Header + Sections + Entries + Bullets), previewing it, and exporting it to multiple formats. The UI is optimized for fast editing, reordering content via drag-and-drop, and maintaining consistent formatting across entries.



This project is designed to feel like a “resume content IDE”:
- You edit structured fields instead of fighting layout/editing in Word.
<img width="550" height="350" alt="image" src="https://github.com/user-attachments/assets/dc52e248-ab63-4d68-bbe2-7165fde5b6e1" />

- You reorder content instantly (sections, entries, bullets).
 <img width="550" height="350" alt="image" src="https://github.com/user-attachments/assets/b518965b-dcb5-4fb2-baef-6392eedec6e2" />

- You can preview and export when ready.
- Edit Spellcheck status and word filtering rules
<img width="550" height="350" alt="image" src="https://github.com/user-attachments/assets/ea551907-2d5d-4d27-8302-83291fd699b0" />

- Live Previews show progress while editing.
<img width="550" height="350" alt="image" src="https://github.com/user-attachments/assets/bb821f19-15be-4ebc-9f57-f1a21899bb3c" />




---
## Table of Contents
- [Key Features](#key-features)
- [How the App Thinks (Data Model)](#how-the-app-thinks-data-model)
- [Quick Start (Workflow)](#quick-start-workflow)
- [Detailed Usage Guide](#detailed-usage-guide)
  - [Projects (.json files)](#projects-json-files)
  - [Header](#header)
  - [Sections](#sections)
  - [Entries](#entries)
  - [Bullets](#bullets)
  - [Spellcheck](#spellcheck)
  - [Undo / Redo](#undo--redo)
  - [Preview](#preview)
  - [Export](#export)
- [Help Menu](#help-menu)
- [Dependencies](#dependencies)
- [Troubleshooting](#troubleshooting)
- [Tips for Resume Writing](#tips-for-resume-writing)
---
## Key Features
### Core editing
- **Header editor** for contact info.
- **Sections list** (drag-and-drop reorder).
- **Entries list** per section (drag-and-drop reorder).
- **Entry editor dialogs** for editing a selected entry’s fields.
- **Bullet editor** with drag-and-drop reordering and rich text support.
### Productivity
- **Undo / Redo** support for common edits and reorders.
- **Ctrl+S Save** (safe: only triggers on actual Ctrl+S).
- **Preview** before exporting.
### Spellcheck
- **Per-word red underline** spellcheck for Text-based inputs.
- Right-click suggestions for misspelled words.
- **Ignore All** to ignore a word for the current project.
- **Edit Document Spellcheck Data...** to manage ignored words (add/delete/apply).
- Spellcheck ignore list persists with the project.
### Output formats
- **Export LaTeX (.tex)**
- **Export Word (.docx)**
- **Export PDF (.pdf)**
---
## How the App Thinks (Data Model)
The application maintains an internal in-memory state that mirrors what gets saved to disk as a JSON project file.
### High-level structure
- **Header**: contact fields (name, email, links, etc.)
- **Sections**: ordered list of sections
- Each **Section** contains:
  - metadata (name/type/kind)
  - ordered list of **Entries**
- Each **Entry** contains:
  - structured fields (depending on section type)
  - optional bullet list / body content (depending on type)
### Why this matters
- **Preview and Export do not read the UI directly**.
- They render from the underlying structured state.
- Saving writes this structured state to your `.json` project file.
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
5. **Add bullets (Experience/Projects)**
   - Open an entry editor.
   - Add bullets and reorder bullets by dragging.
6. **Spellcheck**
   - `Settings -> Enable Spellcheck`
   - Right-click underlined words for suggestions.
   - Use **Ignore All** for words you want to ignore in this project.
   - Manage ignored words: `Settings -> Edit Document Spellcheck Data...`
7. **Preview**
   - Click `Preview` to verify structure and spacing.
8. **Export**
   - `File -> Export LaTeX / Word / PDF`
---
## Detailed Usage Guide
### Projects (.json files)
A “project” is a JSON file containing your entire resume state.
- **Save** writes current in-memory state to disk.
- **Load** replaces current in-memory state with the file contents.
- The app also saves project-specific spellcheck data in the JSON.
**Recommended workflow**
- Create a project file first (so Ctrl+S works silently).
- Save frequently.
---
### Header
The Header area contains your contact details (name, phone, email, LinkedIn, GitHub, etc.).
**Behavior**
- Header changes update the internal state.
- Save persists them to the project JSON.
---
### Sections
A section is a top-level resume block (e.g., Education, Experience, Projects, Skills).
**Key behaviors**
- The **order of sections** is the order used in Preview and Export.
- You can **drag-and-drop** sections to reorder.
- Deleting a section removes its entries from the project data.
---
### Entries
Entries belong to a section (e.g., one job inside Experience, one school inside Education).
**Key behaviors**
- The **order of entries** controls output order.
- Entry fields vary depending on section kind.
- Entry editing happens in an editor dialog.
---
### Bullets
Bullets are typically used for Experience and Projects.
**Key behaviors**
- Bullets are managed as a list.
- Bullet order matters and is reflected in Preview and Export.
- Bullets support rich text editing where applicable.
---
### Spellcheck
Spellcheck is designed to act like a modern editor:
- Misspelled words are underlined in **bright red**.
- Right-click a misspelled word:
  - choose a suggestion to replace it
  - or choose **Ignore All** to ignore that word in this project
#### Ignore list persistence
Ignored words are stored per project as a list in the JSON (e.g. `spellcheck_ignore_all`).
#### Edit Document Spellcheck Data
`Settings -> Edit Document Spellcheck Data...` opens a manager dialog:
- Selectable list of ignored words
- **Add**: add one or more words (space-separated)
- **Delete**: remove selected word
- **Apply**: apply changes immediately to current spellcheck settings
This is useful for:
- names
- acronyms
- tech terms (e.g., “NextJS”, “PostgreSQL”)
- company names
---
### Undo / Redo
Undo/Redo exists to support rapid iteration while editing.
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
- **Word (.docx)**: good for recruiters who prefer Word.
- **PDF (.pdf)**: good for final submission.
Exports are based on the internal structured state, not the raw UI.
---
## Help Menu
The app includes a top ribbon menu:
- `Help -> Quick Start Guide`
- `Help -> Detailed Help`
These open scrollable “rich text” dialogs.
Additionally, the program writes copies of the help content to:
- `Program Help Docs/Quick Start Guide.txt`
- `Program Help Docs/Detailed Help.txt`
This folder is created automatically under the application workspace directory.
---
## Dependencies
Core:
- Python (Tkinter-based GUI)
Optional / feature dependencies:
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
- Word export requires `python-docx`
- PDF export requires `reportlab`
- The error dialog will tell you exactly what’s missing.
### My ignored words disappeared
- Ignored words are stored in the project JSON.
- Make sure you saved your project after editing the ignore list.
---
## Tips for Resume Writing
- Prefer measurable impact: “Reduced build time by 35%” beats “Improved build process.”
- Keep bullets parallel (start with a strong verb).
- Keep Skills scannable: group tools and avoid filler.
- Use Preview often to catch layout issues before exporting.
---
