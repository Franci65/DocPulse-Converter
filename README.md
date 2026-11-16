# DocPulse Converter

**DocPulse Converter** is an ultra-fast, 100% native C++20 / Win32 portable desktop utility that converts **Excel Binary Workbooks (`.xlsb`)** and **Word Documents (`.docx`)** — along with images, web pages, and plain-text/code files — into clean PDF documents without requiring Microsoft Office or an embedded browser.

> *"Strong essences are kept in small bottles — pure Win32, zero bloat, true portability."*

---

## Supported Formats

### Primary Document Engines
* **`.xlsb` (Microsoft Excel Binary Workbook):** Native BIFF12 binary parser (`sharedStrings.bin`, `workbook.bin`, `sheet*.bin`) rendering every worksheet into aligned A4 Landscape tables.
* **`.docx` (Microsoft Word OpenXML):** Native XML extractor preserving paragraphs, tabs, and table cells.

### Bonus Formats
* **Images (via Windows Imaging Component):** `.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, `.tif`, `.tiff`, `.webp`
* **Web & Rich Text:** `.html`, `.htm`, `.rtf`
* **Data & Source Code:** `.csv`, `.tsv`, `.json`, `.xml`, `.md`, `.txt`, `.log`, `.ini`, `.cpp`, `.h`, `.py`, `.js`

---

## Key Features

* **Zero External Dependencies:** Single standalone executable (~1.18 MB), native Dark Mode UI, zero installation, and zero registry modifications.
* **Smart Output Naming:** Automatically suggests `<OriginalFileName>.pdf` in the Save dialog.
* **Drag & Drop + File Picker:** Drop any supported file directly into the window or use the file browser.

---

## Building from Source

### Prerequisites
* Windows 10 / 11 (x64)
* [w64devkit](https://github.com/skeeto/w64devkit) (GCC C++20 / MinGW-w64)

### Build Steps
Run `build.bat` to compile `DocPulseConverter.exe` with Skylake AVX2 optimizations:

```bat
build.bat
```

---

## Repository Structure

* `DocPulseConverter.cpp` — Native Win32 Dark Mode GUI, Drag & Drop handler, WIC image pipeline, and custom PDF 1.4 generator
* `XlsbReader.hpp` — Standalone C++20 BIFF12 binary parser for `.xlsb` workbooks
* `OfficeXmlReader.hpp` — Standalone C++20 parser for `.docx`, `.html`, and `.rtf` documents
* `build.bat` — One-click build script for `w64devkit`
* `LICENSE` — MIT License
