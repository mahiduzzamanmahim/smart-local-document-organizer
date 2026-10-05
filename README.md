# Smart Document Organizer v2 📁

A local, privacy-focused desktop application built with Python and Tkinter that automatically categorizes and organizes messy files into structured folders based on filename heuristics and deep text extraction.

---

## Key Features

- **Local & Privacy-Focused:** All file parsing and classification happen locally on your device—no data ever leaves your machine.
- **Smart Hybrid Classification:** Combines quick regex-based filename scoring with dynamic content extraction for `.pdf`, `.docx`, `.xlsx`, and plain text files.
- **Non-Blocking UI:** Runs file scanning and movement on a background `threading.Thread` with `queue.Queue` messaging to keep the GUI smooth and responsive.
- **Atomic Undo History:** Persists all file operations to a local `organization_history.json` log, allowing complete rollback and restoration of moved files.
- **Collision Prevention:** Uses safe file movement routines to increment filenames automatically if a target file already exists, preventing accidental overwrites.

---

## Supported Categories & File Formats

### Default Categories
- 📑 **Invoices:** Invoices, receipts, tax bills, payment receipts
- 📄 **Resumes:** CVs, resumes, employment history
- 🔬 **Research:** Academic papers, DOIs, methodologies, literature reviews
- 📚 **Study:** Lecture notes, assignments, coursework, exam prep
- 💳 **Finance:** Bank statements, account reports, salary slips
- ⚖️ **Legal:** Contracts, agreements, affidavits, petitions
- 📂 **Documents:** General text documents and reports
- 🖼️ **Images:** `.jpg`, `.png`, `.webp`, `.svg`, and other image formats
- 📦 **Miscellaneous:** Fallback directory for unclassified files

### Supported Document Types for Content Inspection
- PDF Files (`.pdf`)
- Word Documents (`.docx`)
- Excel Spreadsheets (`.xlsx`)
- Plain Text & Markdown (`.txt`, `.md`, `.rtf`)

---

## Getting Started

### Prerequisites
Make sure you have **Python 3.8+** installed on your computer.

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/mahiduzzamanmahim/smart-local-document-organizer.git
   cd smart-local-document-organizer
