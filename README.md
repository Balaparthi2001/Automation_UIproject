# Automation Control Panel

A Flask-based web application for managing and running automation scripts. Provides a modern UI to upload, edit, run, and manage Python automation scripts with real-time terminal output streaming.

## Features

- **Web Dashboard** - Modern UI with dark theme to manage all automations
- **Script Management** - Upload, edit, update, and delete Python automation scripts
- **Real-time Execution** - Run scripts with SSE (Server-Sent Events) streaming terminal output
- **Categorized Organization** - Scripts organized under labels (Label_1 through Label_4)
- **Code Editor** - Built-in editor with run-edited and run-base options
- **Admin Control** - Password-protected admin access for upload/update/delete operations

## Prerequisites

- **Python 3.10+**
- **System Dependencies** (for specific automations only):
  - Tesseract-OCR (for OCR-based scripts)
  - Poppler (for PDF processing)

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/Balaparthi2001/Automation_UIproject.git
cd Automation_UIproject

# 2. Create and activate virtual environment
python -m venv .venv
# Windows:
.venv\Scripts\activate
# Linux/Mac:
source .venv/bin/activate

# 3. Install Python dependencies
pip install -r requirements_auto.txt

# 4. Run the application
python app.py
```

The app will automatically open at `http://127.0.0.1:5000`.

## Project Structure

```
Automation_UIproject/
├── app.py                        # Flask application (backend API)
├── requirements_auto.txt         # Python dependencies
├── .gitignore                    # Git ignore rules
├── README.md                     # This file
├── automations/
│   ├── Label_1/
│   │   ├── pageReplace_baseon_pgnum.py      # Excel-driven PDF page replacement
│   │   ├── PageReplace_byBookmark_new.py    # Bookmark-based PDF page replacement
│   │   └── pdfverifyline_addsent.py         # Batch add text to PDFs
│   ├── Label_3/
│   │   ├── Robocopy.py                      # Excel → Robocopy GUI tool
│   │   └── Stamp_findpdf.py                 # Red border box detection in PDFs
│   └── Label_4/
│       ├── Stamp_findpdf_drawingNumber.py   # DN & SHT extraction with OCR
│       └── SubSysno_based_descrptionChange.py  # Annotation → Excel matching
└── static/
    ├── index.html                # Frontend HTML
    ├── styles.css                # CSS styling
    └── app.js                    # Frontend JavaScript logic
```

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| `GET` | `/` | Serve the web dashboard |
| `GET` | `/files` | List all automation scripts |
| `GET` | `/files/tree` | List scripts organized by category |
| `GET` | `/files/content?path=` | Get file content |
| `POST` | `/upload` | Upload a new script |
| `POST` | `/update` | Update an existing script |
| `DELETE` | `/files?path=` | Delete a script |
| `GET` | `/run/stream?file=&scope=` | Run a script (SSE stream) |
| `POST` | `/run/preview/start` | Start a preview run |
| `POST` | `/run/stop` | Stop a running script |

## Dependencies

- **flask** - Web framework
- **pandas** / **openpyxl** / **xlrd** - Excel processing
- **PyMuPDF (fitz)** - PDF manipulation
- **PyPDF2** - PDF bookmark reading
- **opencv-python** - Computer vision (red box detection)
- **pytesseract** - OCR text extraction
- **pdf2image** - PDF to image conversion
- **Pillow** - Image processing
- **numpy** - Numerical computing

## Usage

1. **Run** - Click on a label, then click "▶ Run" on any script to execute it
2. **Edit** - Click "📝 Edit" to view code, make temporary edits, and run the edited version
3. **Upload** - Upload new `.py` scripts (requires admin password)
4. **Update** - Replace existing scripts with new versions (requires admin password)
5. **Delete** - Remove scripts (requires admin password)

Default admin password: `admin@123` (change in `static/app.js`)