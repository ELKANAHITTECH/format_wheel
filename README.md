# format_wheel
# Format Wheel

**Author:** Damilare Adekeye

A GTA-style **radial file converter for Windows**. Select files in Windows Explorer, hold a hotkey, aim at a format on the wheel, and release to convert your files.

- Hold the hotkey to dim the screen and display a weapon-wheel-style ring of file formats around your cursor.
- Aim with your mouse and release the hotkey, click, or press Enter to convert.
- Convert between PDF, DOCX, XLSX, CSV, TXT, ZIP, TIFF, BMP, WEBP, GIF, ICO, PNG, and JPG.
- Runs quietly in the Windows system tray.

> **Platform:** Windows 10 and Windows 11. The conversion engine is written in Python, while Explorer integration, automatic startup, and Microsoft Office integration are Windows-specific.

---

## Table of Contents

1. [Features](#features)
2. [Quick Start](#quick-start)
3. [Using the Wheel](#using-the-wheel)
4. [Supported Conversions](#supported-conversions)
5. [Dependencies](#dependencies)
6. [Settings and Configuration](#settings-and-configuration)
7. [Building a Standalone Executable](#building-a-standalone-executable)
8. [Troubleshooting](#troubleshooting)
9. [How It Works](#how-it-works)
10. [Known Limitations](#known-limitations)
11. [Contributing](#contributing)
12. [Author](#author)
13. [License](#license)

---

## Features

### Radial Wheel Interface

- Full-screen dimming overlay with a vignette effect.
- Animated radial menu with glowing selection indicators.
- Hidden cursor and a visual aiming indicator.
- Unsupported conversions are greyed out and marked `UNAVAILABLE`.
- Mouse, keyboard, and scroll-wheel navigation.
- Animated notifications with a progress indicator.
- Click completed notifications to open the output folder.

### File Conversion Engine

- Background conversion processing to keep the interface responsive.
- Multiple conversion routes with fallbacks for missing libraries.
- Microsoft Word, Microsoft Excel, LibreOffice, and built-in rendering support where applicable.
- Multi-page PDF and TIFF handling.
- Image conversion with transparency, EXIF rotation, and animated GIF/WebP support.
- Support for password-protected PDFs where the relevant library and password handling permit it.
- Optional OCR support for scanned PDFs and images.
- Batch conversion of multiple files.
- ZIP archive creation for multiple files and folders.

### Productivity Features

- Reads selected files from Windows Explorer.
- Falls back to files copied to the clipboard.
- Optional conversion of clipboard images, including screenshots.
- Configurable global hotkey.
- Configurable output directory and theme colour.
- Optional startup with Windows.
- Single-instance protection.
- Optional copying of converted files to the clipboard.

---

## Quick Start

### 1. Requirements

- Windows 10 or Windows 11.
- Python 3.9 or newer.
- Internet connection for installing dependencies.

Download Python from [python.org](https://www.python.org/downloads/) and enable **Add Python to PATH** during installation.

### 2. Clone the Repository

Replace `YOUR_GITHUB_USERNAME` with your GitHub username or organisation.

```powershell
git clone https://github.com/YOUR_GITHUB_USERNAME/format-wheel.git
cd format-wheel
```

### 3. Create a Virtual Environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks virtual-environment activation, you can install packages using the environment's Python executable directly.

### 4. Install Dependencies

Install the core dependencies:

```powershell
python -m pip install --upgrade pip
python -m pip install PyQt6 pynput Pillow python-docx openpyxl pywin32
```

Install the recommended conversion dependencies:

```powershell
python -m pip install pymupdf pdf2docx reportlab pypdf xlrd
```

For OCR functionality, install the Python package:

```powershell
python -m pip install pytesseract
```

You must also install the Tesseract OCR application separately. See the [Tesseract installation guide for Windows](https://github.com/UB-Mannheim/tesseract/wiki).

### 5. Run Format Wheel

```powershell
python format_wheel.py
```

The application starts in the system tray. If you cannot see its icon, check the hidden-icons menu in the Windows taskbar.

To run without displaying a console window:

```powershell
pythonw format_wheel.py
```

---

## Using the Wheel

1. Open Windows Explorer and select one or more files. Select folders when creating ZIP archives.
2. Hold the configured activation hotkey. The default is `Ctrl+Shift+C`.
3. Move your mouse toward the desired output format.
4. Release the hotkey, click the selected format, or press Enter to begin conversion.

Converted files are saved to `Documents\FormatWheel` by default.

### Keyboard and Mouse Controls

| Input | Action |
|---|---|
| Hold hotkey, aim, release | Convert to the selected format |
| Left-click | Convert to the highlighted format |
| `Enter` or `Space` | Confirm the highlighted format |
| Arrow keys or `Tab` | Navigate available formats |
| Mouse scroll | Cycle through formats |
| `Esc` or `Backspace` | Cancel |
| Right-click | Cancel |
| Click the centre | Cancel |
| Press the hotkey again | Close the wheel |

You can also use the system-tray menu and select **Convert selected files...**.

### Input File Selection

Format Wheel searches for input files in this order:

1. Files selected in the active or most recent Windows Explorer window.
2. Files copied using `Ctrl+C`.
3. An image on the clipboard, such as a screenshot, if clipboard-image support is enabled.

If no supported input is found, the application displays a notification.

---

## Supported Conversions

The following table describes the intended conversion capabilities. Actual availability depends on the input file, installed dependencies, and available rendering engines.

| Source | Images | PDF | DOCX | XLSX | CSV | TXT | ZIP |
|---|---|---|---|---|---|---|---|
| Images | Yes | Yes | Embedded images | Embedded images | OCR required | OCR required | Yes |
| Text, code, Markdown, HTML | Rendered | Yes | Yes | Yes | Yes | Yes | Yes |
| CSV / TSV | Rendered | Yes | Table | Columns | Copy | Yes | Yes |
| DOCX / DOC | Yes | Yes | Yes | Tables | Tables | Yes | Yes |
| XLSX / XLS / XLSM | Yes | Yes | Tables | Yes | One file per sheet | Yes | Yes |
| PDF | Rendered pages | Copy | Rich conversion | Tables | Tables | Yes | Yes |
| Folders | No | No | No | No | No | No | Yes |

**Legend:** Yes means supported through an appropriate conversion route. OCR, Microsoft Office, LibreOffice, and optional Python packages may be required for certain operations.

### Notable Conversion Behaviour

- **PDF to DOCX:** Uses `pdf2docx` when available, with a PyMuPDF-based fallback. Layout reconstruction is best-effort and may not perfectly preserve complex documents.
- **PDF to images:** Renders actual PDF pages. Multi-page output is combined into one TIFF or saved as separate image files, depending on the target format.
- **DOCX/XLSX to PDF or images:** Uses Microsoft Office when available, then LibreOffice, then a simplified built-in renderer.
- **Images to ICO:** Creates a square icon with multiple standard resolutions.
- **XLSX to CSV:** Exports individual worksheets as separate CSV files using UTF-8 with a BOM.
- **Transparent images:** Transparency is flattened onto a white background when converting to formats that do not support it.
- **Animated GIF/WebP:** Supported routes may preserve animation. Other output formats may use only the first frame.
- **Same-format conversions:** A conversion such as PNG to PNG is disabled.
- **Multiple files to ZIP:** Creates a single timestamped archive.

---

## Dependencies

Format Wheel uses a combination of required Python packages and optional conversion engines.

### Core Dependencies

| Package | Purpose |
|---|---|
| `PyQt6` | User interface, radial wheel, notifications, and settings |
| `pynput` | Global keyboard and mouse input |
| `Pillow` | Image processing and conversion |
| `python-docx` | Microsoft Word document handling |
| `openpyxl` | Excel workbook handling |
| `pywin32` | Windows integration and Microsoft Office automation |

### Optional Dependencies

| Package or Application | Purpose |
|---|---|
| `pymupdf` | PDF rendering, text extraction, and document reconstruction |
| `pdf2docx` | PDF-to-DOCX layout conversion |
| `reportlab` | PDF generation |
| `pypdf` | PDF reading and fallback processing |
| `xlrd` | Reading legacy `.xls` workbooks |
| `pytesseract` | Python interface for OCR |
| Tesseract OCR | OCR engine required by `pytesseract` |
| Microsoft Word and Excel | Office automation and legacy document support |
| LibreOffice | Alternative document conversion engine |

### Install All Python Dependencies

```powershell
python -m pip install PyQt6 pynput Pillow python-docx openpyxl pywin32 pymupdf pdf2docx reportlab pypdf xlrd pytesseract
```

**Important:** Installing Python packages does not install Microsoft Office, LibreOffice, or the Tesseract OCR application. Install those separately if you need their functionality.

---

## Settings and Configuration

Open **Settings...** from the system-tray menu. Supported settings apply immediately.

| Setting | Default | Description |
|---|---|---|
| Activation hotkey | `ctrl+shift+c` | Global shortcut used to open the wheel |
| Output folder | `Documents\FormatWheel` | Destination for converted files |
| Theme colour | `#00e5ff` | Accent colour for the interface |
| Wheel radius | `250` | Outer wheel radius, in pixels |
| Screen dimming | `150` | Darkness of the overlay |
| Hold, aim, release | Enabled | Converts when the activation hotkey is released |
| Start with Windows | Disabled | Enables startup for the current user |
| Clipboard image | Enabled | Allows clipboard images to be converted |
| Copy result to clipboard | Disabled | Copies converted files to the clipboard |
| Open output folder | Disabled | Opens the destination after conversion |
| Quiet success | Disabled | Hides successful-conversion notifications |

### Choosing a Hotkey

On Windows 11, `Ctrl+Shift+C` is also used by Explorer's **Copy as path** command. If the shortcut conflicts with your workflow or another application, change it in Settings.

For example:

```text
ctrl+alt+w
ctrl+shift+f9
```

### Configuration Files

Application settings and diagnostics are stored under:

```text
%APPDATA%\FormatWheel\
```

| File | Purpose |
|---|---|
| `config.json` | Application settings |
| `format_wheel.log` | Errors and diagnostic information |

Close the application before manually editing `config.json`.

---

## Building a Standalone Executable

Format Wheel can be packaged as a Windows executable using [PyInstaller](https://pyinstaller.org/).

### 1. Install PyInstaller

```powershell
python -m pip install pyinstaller
```

### 2. Build the Executable

```powershell
pyinstaller --noconsole --onefile --name FormatWheel format_wheel.py
```

The resulting executable should be available at:

```text
dist\FormatWheel.exe
```

If a packaged feature fails because a dependency was not included, rebuild with the required PyInstaller collection options. For example:

```powershell
pyinstaller --noconsole --onefile --name FormatWheel --collect-all pymupdf format_wheel.py
```

Test the generated executable on a clean Windows environment to verify that all required libraries and external applications are available.

---

## Troubleshooting

### The Hotkey Does Nothing

- Confirm that Format Wheel is running and visible in the system tray.
- Check whether another application is using the same shortcut.
- Open Settings and try a different hotkey.
- Verify that `pynput` is installed.
- Try running with the necessary permissions if interacting with elevated applications.

### Nothing Is Selected

- Select files in Windows Explorer before opening the wheel.
- For Desktop files or selections that Explorer integration cannot read, select the files and press `Ctrl+C` before activating the wheel.
- Confirm that the selected files are accessible.

### A Format Is Greyed Out

- The selected conversion may not be supported for that input.
- A required Python dependency may be missing.
- Install the relevant optional packages and restart the application if necessary.

### PDF to DOCX Fails

Install the PDF processing dependencies:

```powershell
python -m pip install pymupdf pdf2docx pypdf python-docx
```

Restart Format Wheel and try again.

For scanned PDFs, install `pytesseract` and the Tesseract OCR application. Password-protected documents may also require the appropriate password-handling support.

### Other PDF Conversions Fail

Install the PDF-related packages:

```powershell
python -m pip install pymupdf pypdf reportlab
```

The appropriate package depends on whether the conversion requires PDF reading, rendering, or generation.

### DOCX or XLSX to PDF Looks Incorrect

- Install Microsoft Word or Excel if you need Microsoft Office automation.
- Alternatively, install [LibreOffice](https://www.libreoffice.org/).
- Without a supported external rendering engine, the application may use simplified built-in rendering.

### Legacy DOC or XLS Files Fail

- Legacy `.doc` files generally require Microsoft Word or LibreOffice.
- Legacy `.xls` files may require Excel, LibreOffice, or `xlrd`.

### Missing or Garbled Characters

The built-in renderer may use a monospaced font such as Consolas. Characters outside the font's coverage may not render correctly.

### General Errors

Open the application log:

```text
%APPDATA%\FormatWheel\format_wheel.log
```

When reporting an issue, include the error message, Windows version, Python version, installed optional packages, and the source and target formats.

---

## How It Works

The main implementation is contained in `format_wheel.py` and is organised into functional sections.

| Component | Responsibility |
|---|---|
| Configuration and registry | Application settings, format definitions, and wheel ordering |
| Helpers | Encoding detection, file identification, font handling, and OCR checks |
| `Doc` model | Intermediate document representation |
| Readers | Load text, CSV, HTML, DOCX, XLSX, PDF, and image files |
| Office bridges | Microsoft Office automation and LibreOffice integration |
| Writers | Generate DOCX, XLSX, CSV, TXT, ZIP, PDF, and image output |
| `convert_file()` | Selects the appropriate conversion route |
| `RadialWheel` | Wheel interface, animation, and user input |
| `Toast` and `SettingsDialog` | Notifications and configuration interface |
| `HotkeyManager` | Global hotkey registration and listening |
| Application integration | System tray, input selection, and background conversion |

The core conversion architecture follows:

**Source file → Intermediate document (`Doc`) → Target file**

To add a new input format, implement a reader that produces the intermediate document representation. To add a new output format, implement a writer and register the format in `FORMATS`.

---

## Known Limitations

- **Windows-focused integration:** Explorer selection, startup registration, and Office automation use Windows-specific APIs.
- **PDF layout reconstruction:** Multi-column PDFs, complex tables, forms, and unusual fonts may not be reproduced perfectly in DOCX.
- **External rendering dependency:** DOCX/XLSX conversion may lose formatting or embedded images when Microsoft Office and LibreOffice are unavailable.
- **Memory usage:** Large files are processed in memory and may require substantial RAM.
- **Explorer tabs:** Multiple Windows 11 Explorer tabs may complicate detection of the intended file selection.
- **Optional dependencies:** Some conversions require additional Python libraries or external applications.
- **OCR accuracy:** Scanned documents and images may produce incomplete or inaccurate text depending on image quality and OCR configuration.
- **Executable packaging:** External applications such as LibreOffice and Tesseract may need to be installed separately even when Format Wheel is packaged as an executable.

---

## Contributing

Contributions, bug reports, and feature suggestions are welcome.

1. Fork the repository.
2. Create a feature branch:

   ```powershell
   git checkout -b feature/my-change
   ```

3. Implement and test your changes.
4. Keep the code readable and maintain existing project conventions.
5. Submit a pull request explaining the change and its purpose.

When reporting a bug, include your Windows version, Python version, relevant dependencies, source and target formats, and the contents of `format_wheel.log` where appropriate.

### Potential Improvements

- Additional formats such as RTF, ODT, HEIC, and SVG.
- Better PDF column and table reconstruction.
- Configurable radial-wheel ordering.
- Improved cross-platform input selection.
- Expanded automated tests for conversion routes and dependency fallbacks.

---

## Author

**Damilare Adekeye**

Format Wheel is an independent file-conversion utility designed to make everyday file conversion faster through a radial interface inspired by video-game weapon wheels.

---

## License

This project is released under the **MIT License**.

Include a `LICENSE` file in the repository containing the standard MIT License text, with the copyright holder set to:

```text
Copyright (c) 2026 Damilare Adekeye
```

See the [MIT License](https://opensource.org/license/mit) for the full license text.

---

*Inspired by the radial weapon-selection interface in the Grand Theft Auto games. Format Wheel is an independent project and is not affiliated with, endorsed by, or sponsored by Rockstar Games or Take-Two Interactive.*
