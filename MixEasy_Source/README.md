# MixEasy Source Code - Modular Architecture

## Overview
MixEasy is a Vietnamese educational tool for creating, mixing/shuffling exam questions,
formatting Word documents, and generating various output formats (PDF, HTML, PowerPoint).

**Version:** V2026.04.25  
**Python:** 3.12  
**Platform:** Windows (uses win32com for Word/PowerPoint automation)

## Project Structure

```
MixEasy_Source/
├── main.py                    # Application entry point (GUI)
├── config.py                  # App configuration (version, dates, links)
├── requirements.txt           # Python dependencies
│
├── core/                      # Core utilities
│   └── functions.py           # Shared Word/document utility functions
│
├── registration/              # License & activation
│   ├── dangki_pc.py           # PC-based registration (HDD serial)
│   └── dangki_usb.py          # USB-based registration (USB serial)
│
├── mixing/                    # Exam mixing/shuffling engine
│   ├── mixeasy.py             # Main mixing orchestrator (win32com)
│   ├── mixeasy_file_mau.py    # Template file handling (python-docx)
│   ├── mix_spr.py             # Vietnamese mixing via Spire.Doc
│   ├── mix_spr_en.py          # English/foreign lang mixing via Spire.Doc
│   ├── mix_spr_func.py        # Shared Spire.Doc mixing functions
│   ├── mix_docx.py            # Vietnamese mixing via python-docx
│   ├── mix_docx_en.py         # English mixing via python-docx
│   ├── mix_docx_func.py       # Shared python-docx mixing functions
│   └── mix_docx_check_data.py # Data validation before mixing
│
├── formatting/                # Document standardization
│   ├── chuan_hoa.py           # Main formatting (win32com)
│   ├── chuan_hoa_docx.py      # Formatting via python-docx
│   └── chuan_hoa_spr.py       # Formatting via Spire.Doc
│
├── answers/                   # Answer sheet processing
│   ├── dapan_2025.py          # Answer table generation
│   └── dapan_danh_dau.py      # Answer marking/highlighting
│
├── tools/                     # Standalone tools
│   ├── tool_by_docx.py        # Document tools (python-docx based)
│   ├── tool_by_spr.py         # Document tools (Spire.Doc based)
│   ├── rename.py              # Batch file renaming
│   └── qr_code.py             # QR code generation
│
├── conversion/                # Format converters
│   ├── word_to_pp.py          # Word → PowerPoint
│   ├── pdf.py                 # Word → PDF
│   ├── omml_to_latex.py       # Office Math (OMML) → LaTeX
│   ├── latex_to_omml.py       # LaTeX → OMML
│   ├── pdf_to_word.py         # PDF → Word (via olmocr)
│   └── word_to_html.py        # Word → HTML
│
├── html_generators/           # HTML output generators
│   ├── trac_nghiem_online.py      # Online quiz HTML
│   ├── trac_nghiem_online_html.py # Online quiz HTML template
│   ├── trac_nghiem_offline.py     # Offline quiz HTML
│   ├── trac_nghiem_offline_html.py# Offline quiz HTML template
│   ├── vong_quay.py               # Spinning wheel HTML
│   ├── bang_diem.py               # Grade board HTML (React)
│   └── dong_ho_dem_nguoc.py       # Countdown timer
│
└── gist/                      # GitHub Gist integration
    ├── gist_manager.py        # Gist CRUD operations
    ├── shortlink_huy.py       # Shortlink manager (Huy)
    └── shortlink_mixeasy.py   # Shortlink manager (MixEasy)
```

## Module Dependencies

### Dependency Layers (bottom = no internal deps, top = most deps)

```
Layer 4 (Entry):  main.py
                    ↓
Layer 3 (Orchestrators):
  mixing/mixeasy.py ← formatting/chuan_hoa.py, tools/tool_by_docx.py,
                       mixing/mix_docx_check_data.py, mixing/mixeasy_file_mau.py
  html_generators/trac_nghiem_*.py ← conversion/omml_to_latex.py,
                                      conversion/word_to_html.py, mixing/mixeasy.py
                    ↓
Layer 2 (Mid-level):
  formatting/chuan_hoa.py     ← core/functions.py, tools/tool_by_docx.py
  mixing/mix_spr.py           ← formatting/chuan_hoa_spr.py, mixing/mix_spr_func.py
  mixing/mix_docx.py          ← formatting/chuan_hoa_docx.py
  mixing/mixeasy_file_mau.py  ← mixing/mix_docx_func.py, mixing/mix_docx_en.py
  formatting/chuan_hoa_spr.py ← mixing/mix_spr_func.py
  formatting/chuan_hoa_docx.py← tools/tool_by_docx.py
  answers/dapan_*.py          ← core/functions.py
  conversion/pdf.py           ← core/functions.py
  conversion/word_to_pp.py    ← core/functions.py
                    ↓
Layer 1 (Foundation - no internal deps):
  core/functions.py
  tools/tool_by_docx.py, tools/tool_by_spr.py, tools/rename.py, tools/qr_code.py
  mixing/mix_spr_func.py, mixing/mix_docx_func.py, mixing/mix_docx_check_data.py
  conversion/omml_to_latex.py, conversion/latex_to_omml.py, conversion/pdf_to_word.py
  html_generators/vong_quay.py, html_generators/bang_diem.py
  html_generators/dong_ho_dem_nguoc.py
  html_generators/trac_nghiem_*_html.py (templates)
  registration/dangki_pc.py, registration/dangki_usb.py
  gist/*
```

## Key Dependencies (External)

| Package | Usage |
|---------|-------|
| `pywin32` (win32com) | Word/PowerPoint COM automation |
| `spire.doc` | Document processing (Spire.Doc for Python) |
| `python-docx` | .docx file manipulation |
| `openpyxl` | Excel file handling |
| `lxml` | XML processing |
| `qrcode` | QR code generation |
| `wmi` | Windows hardware info (HWID) |
| `Pillow` (PIL) | Image processing |
| `requests` | HTTP requests (Gist API) |
| `mammoth` | Word to HTML conversion |
| `officemath2latex` | OMML to LaTeX conversion |
| `latex2mathml` | LaTeX to MathML conversion |
| `mathml2omml` | MathML to OMML conversion |

## Notes

- Many functions have `# WARNING: Decompyle incomplete` markers where the
  decompiler (Decompyle++) could not fully reconstruct the Python 3.12 bytecode.
- The application uses two document processing backends:
  - **win32com** (COM automation) for operations requiring the actual Word application
  - **python-docx / Spire.Doc** for offline document manipulation
- Registration uses HDD/USB serial numbers as hardware IDs.
