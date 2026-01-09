# Research: Microsoft Office to PDF Conversion in Python
## For Hospital Environment - Senior User Application

**Research Date:** January 9, 2026
**Context:** Adding Office-to-PDF conversion to the PDF Locker application for senior users (70+) in hospital settings

---

## Executive Summary

For a hospital environment with Windows + Microsoft Office already installed and targeting senior users, the **recommended approach** is:

**Primary Solution: `docx2pdf` + `comtypes` (for Windows with Office installed)**
- Leverages existing Office installation
- MIT licensed (commercial-friendly)
- Simple API
- Excellent conversion quality (uses native Office)
- Easy deployment for end users

**Backup Solution: LibreOffice Headless (for environments without Office)**
- Free and open-source (MPL 2.0)
- No Office license required
- Cross-platform
- Good conversion quality
- Larger deployment footprint

---

## Detailed Analysis by Approach

### 1. `docx2pdf` Library

**License:** MIT License (OSI Approved)
- ✅ Commercial use: YES
- ✅ Modification: YES
- ✅ Distribution: YES
- ✅ Private use: YES

**How it works:**
- Uses Microsoft Office COM interface internally
- Requires Microsoft Office to be installed
- Windows primary, macOS support via AppleScript

**Code Example:**
```python
from docx2pdf import convert

# Single file
convert("input.docx", "output.pdf")

# Batch conversion
convert("input_folder/", "output_folder/")
```

**Pros:**
- ✅ Very simple API
- ✅ Excellent conversion quality (uses native Office)
- ✅ MIT licensed - no restrictions
- ✅ Handles Word documents well
- ✅ Small footprint (just a wrapper)
- ✅ Perfect for hospital with Office installed

**Cons:**
- ❌ Requires Microsoft Office installed
- ❌ Windows/macOS only (no Linux)
- ❌ Only handles Word documents (.docx, .doc)
- ❌ Doesn't support Excel or PowerPoint
- ❌ Subject to Office COM automation limitations

**Dependencies:**
- Windows: `comtypes` or `pywin32`
- macOS: `appscript`

**Best for:** Hospital Windows environment with Office installed, Word documents only

---

### 2. `comtypes` with Microsoft Office COM Interface

**License:** MIT License
- ✅ Commercial use: YES
- ✅ Full permissions for modification and distribution

**How it works:**
- Direct COM automation of Office applications
- Pure Python implementation (no C extensions)
- Requires Office installed on the system

**Code Example:**
```python
import comtypes.client

def convert_word_to_pdf(input_path, output_path):
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(input_path)
    doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF
    doc.Close()
    word.Quit()

def convert_excel_to_pdf(input_path, output_path):
    excel = comtypes.client.CreateObject('Excel.Application')
    excel.Visible = False
    wb = excel.Workbooks.Open(input_path)
    wb.ExportAsFixedFormat(0, output_path)  # 0 = PDF
    wb.Close()
    excel.Quit()

def convert_powerpoint_to_pdf(input_path, output_path):
    powerpoint = comtypes.client.CreateObject('Powerpoint.Application')
    presentation = powerpoint.Presentations.Open(input_path)
    presentation.SaveAs(output_path, 32)  # 32 = PDF
    presentation.Close()
    powerpoint.Quit()
```

**Pros:**
- ✅ MIT licensed - fully commercial-friendly
- ✅ Supports Word, Excel, PowerPoint
- ✅ Excellent conversion quality (native Office)
- ✅ Pure Python (no C extensions)
- ✅ More control than docx2pdf
- ✅ Can handle advanced Office features

**Cons:**
- ❌ Requires Microsoft Office installed
- ❌ Windows only
- ❌ **Microsoft does not support unattended automation** (see warnings below)
- ❌ Can be unstable in server/unattended scenarios
- ❌ Dialog boxes can cause hangs
- ❌ More complex code than docx2pdf

**CRITICAL WARNINGS from Microsoft:**
Microsoft officially states:
- "Does not currently recommend, and does not support, Automation of Microsoft Office applications from any unattended, non-interactive client application"
- Office may exhibit unstable behavior and/or deadlock
- Dialog boxes can cause application hangs
- Security risks with macro execution
- Not suitable for high-volume server scenarios

**Best for:** Desktop applications with user present, full Office format support needed

---

### 3. `pywin32` Alternative

**License:** Python Software Foundation License (complex, multiple components)
- Generally permissive but more complex licensing than MIT

**Comparison with comtypes:**
- Similar functionality to comtypes
- Better dispatch-based COM support
- Not pure Python (includes C extensions)
- More complex licensing situation
- Both have same Office automation limitations

**Recommendation:** Use `comtypes` instead due to simpler licensing and pure Python implementation.

---

### 4. LibreOffice Command-Line Conversion

**License:** MPL 2.0 (Mozilla Public License) / LGPLv3 dual license
- ✅ Commercial use: YES - completely free
- ✅ Use in business/hospital: YES
- ✅ No licensing fees ever

**How it works:**
- LibreOffice runs in "headless" mode (no GUI)
- Executes via command line
- Python subprocess calls the command

**Code Example:**
```python
import subprocess
from pathlib import Path

def convert_to_pdf_libreoffice(input_path, output_dir):
    """Convert Office file to PDF using LibreOffice"""
    cmd = [
        'soffice',
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', str(output_dir),
        str(input_path)
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise Exception(f"Conversion failed: {result.stderr}")

    # Output file has same name with .pdf extension
    output_file = Path(output_dir) / f"{Path(input_path).stem}.pdf"
    return output_file
```

**Pros:**
- ✅ Completely free - no license costs
- ✅ Cross-platform (Windows, Mac, Linux)
- ✅ Supports Word, Excel, PowerPoint, and more
- ✅ No Office installation required
- ✅ Stable for unattended/server use
- ✅ Good conversion quality
- ✅ No Microsoft COM limitations
- ✅ Can be bundled with application

**Cons:**
- ❌ Requires LibreOffice installation (~400MB)
- ❌ Conversion quality slightly lower than native Office
- ❌ May have formatting differences vs. Office
- ❌ Slower startup time (launches full LibreOffice engine)
- ❌ Complex fonts may not render identically

**Python Wrappers:**
- `unoconv` (GPL-2.0 license) - command-line tool
- `unoserver` (MIT License) - newer, server-based
- Direct subprocess calls (no additional dependencies)

**Best for:**
- Environments without Office
- Cross-platform needs
- Server/unattended scenarios
- Cost-sensitive deployments

---

### 5. `python-docx` + `reportlab` Approach

**Licenses:**
- `python-docx`: MIT License ✅
- `reportlab`: BSD License (open-source) / Commercial license available ✅

**How it works:**
- Read Word document with python-docx
- Parse content and structure
- Generate PDF with reportlab
- Manual layout recreation

**Code Concept:**
```python
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def docx_to_pdf_manual(input_path, output_path):
    doc = Document(input_path)
    pdf = canvas.Canvas(output_path, pagesize=letter)

    y_position = 750
    for paragraph in doc.paragraphs:
        pdf.drawString(72, y_position, paragraph.text)
        y_position -= 20

    pdf.save()
```

**Pros:**
- ✅ Both libraries MIT/BSD licensed
- ✅ No Office required
- ✅ Cross-platform
- ✅ Full control over PDF generation
- ✅ Can customize output extensively

**Cons:**
- ❌ **MAJOR**: Extremely complex to implement properly
- ❌ Must manually handle: fonts, styles, tables, images, headers, footers
- ❌ Layout preservation is very difficult
- ❌ Poor conversion quality without extensive work
- ❌ Thousands of lines of code for full fidelity
- ❌ Ongoing maintenance burden
- ❌ Not practical for production use

**Verdict:** NOT RECOMMENDED for hospital application. Too complex, poor quality.

---

### 6. `openpyxl` + PDF Generation (Excel)

**License:** MIT License ✅

**Reality Check:**
- `openpyxl` reads/writes Excel files
- **Does NOT have built-in PDF export**
- Would need reportlab or similar for PDF generation
- Same problems as python-docx + reportlab approach

**Alternative Approach:**
```python
import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table

def excel_to_pdf_manual(input_path, output_path):
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    # Extract data
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(list(row))

    # Create PDF
    doc = SimpleDocTemplate(output_path, pagesize=letter)
    table = Table(data)
    doc.build([table])
```

**Cons:**
- ❌ No formatting preservation
- ❌ No formulas, charts, or complex features
- ❌ Only basic table data
- ❌ Poor quality output

**Verdict:** NOT RECOMMENDED. Use comtypes or LibreOffice instead.

---

### 7. `python-pptx` Approach (PowerPoint)

**License:** MIT License ✅

**Reality Check:**
- `python-pptx` reads/writes PowerPoint files
- **Does NOT have PDF export capability**
- Would need external PDF generation

**Options:**
1. Use comtypes to automate PowerPoint (see section 2)
2. Use LibreOffice headless
3. Use commercial library (Aspose.Slides)

**Verdict:** NOT RECOMMENDED as standalone. Use comtypes or LibreOffice.

---

### 8. Commercial Solutions

#### Aspose (Words/Cells/Slides for Python)

**License:** Commercial - Requires paid license
- Free trial available
- Per-developer licensing
- Costs: ~$1,000+ per developer

**Pros:**
- ✅ Professional quality
- ✅ No Office required
- ✅ Cross-platform
- ✅ Excellent conversion quality
- ✅ All Office formats
- ✅ Good documentation

**Cons:**
- ❌ Expensive
- ❌ Ongoing license costs
- ❌ Not suitable for hospital budget constraints

**Verdict:** Overkill for this application. Free solutions are adequate.

#### Apryse SDK (PDFTron)

**License:** Commercial - Requires paid license

**Pros:**
- ✅ No Office required
- ✅ High quality
- ✅ Cross-platform

**Cons:**
- ❌ Very expensive (enterprise pricing)
- ❌ Complex licensing

**Verdict:** Not suitable for senior user application.

---

## Comparison Matrix

| Solution | License | Office Required | Cross-Platform | Quality | Complexity | Cost | Best For |
|----------|---------|----------------|----------------|---------|------------|------|----------|
| **docx2pdf** | MIT ✅ | YES (Word) | Win/Mac | Excellent | Very Low | Free | Hospital + Office |
| **comtypes** | MIT ✅ | YES (All) | Windows only | Excellent | Medium | Free | Full Office support |
| **pywin32** | PSF | YES (All) | Windows only | Excellent | Medium | Free | Alternative to comtypes |
| **LibreOffice** | MPL 2.0 ✅ | NO | All | Good | Low | Free | No Office installed |
| **python-docx + reportlab** | MIT/BSD ✅ | NO | All | Poor | Very High | Free | NOT RECOMMENDED |
| **openpyxl + PDF** | MIT ✅ | NO | All | Poor | Very High | Free | NOT RECOMMENDED |
| **python-pptx** | MIT ✅ | NO | All | N/A | N/A | Free | NOT RECOMMENDED |
| **Aspose** | Commercial ❌ | NO | All | Excellent | Low | $1000+ | Enterprise only |
| **Apryse** | Commercial ❌ | NO | All | Excellent | Medium | $$$$$ | Enterprise only |

---

## Recommended Implementation for Hospital Environment

### Scenario A: Hospital with Microsoft Office Installed (Most Common)

**Recommended Stack:**
1. **Primary:** `docx2pdf` (for Word documents)
2. **Extended:** `comtypes` (for Excel and PowerPoint)

**Implementation Strategy:**
```python
# requirements.txt
docx2pdf>=0.8.1
comtypes>=1.4.0

# conversion.py
from pathlib import Path
from docx2pdf import convert as docx2pdf_convert
import comtypes.client

class OfficeConverter:
    """Simple Office to PDF converter for hospital environment"""

    def convert_to_pdf(self, input_path, output_path=None):
        """
        Convert Office file to PDF

        Args:
            input_path: Path to Office file
            output_path: Output PDF path (optional)
        """
        input_path = Path(input_path)

        if output_path is None:
            output_path = input_path.with_suffix('.pdf')

        extension = input_path.suffix.lower()

        if extension in ['.docx', '.doc']:
            # Use docx2pdf for Word (simplest)
            docx2pdf_convert(str(input_path), str(output_path))

        elif extension in ['.xlsx', '.xls']:
            # Use comtypes for Excel
            self._convert_excel(input_path, output_path)

        elif extension in ['.pptx', '.ppt']:
            # Use comtypes for PowerPoint
            self._convert_powerpoint(input_path, output_path)

        else:
            raise ValueError(f"Unsupported file type: {extension}")

        return output_path

    def _convert_excel(self, input_path, output_path):
        """Convert Excel to PDF using comtypes"""
        excel = comtypes.client.CreateObject('Excel.Application')
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(str(input_path.absolute()))
            wb.ExportAsFixedFormat(0, str(output_path.absolute()))
            wb.Close()
        finally:
            excel.Quit()

    def _convert_powerpoint(self, input_path, output_path):
        """Convert PowerPoint to PDF using comtypes"""
        powerpoint = comtypes.client.CreateObject('Powerpoint.Application')
        try:
            presentation = powerpoint.Presentations.Open(str(input_path.absolute()))
            presentation.SaveAs(str(output_path.absolute()), 32)  # 32 = PDF
            presentation.Close()
        finally:
            powerpoint.Quit()
```

**Pros of this approach:**
- Leverages existing Office installation
- MIT licensed throughout
- Excellent conversion quality
- Simple deployment (small dependencies)
- Familiar output for users

**Considerations:**
- Warn users if dialogs appear
- Handle COM errors gracefully
- Test thoroughly with user present
- Consider fallback to LibreOffice

---

### Scenario B: Hospital without Microsoft Office (Less Common)

**Recommended:** LibreOffice Headless

**Implementation:**
```python
# requirements.txt
# No additional Python packages needed!
# Just ensure LibreOffice is installed

# conversion.py
import subprocess
from pathlib import Path

class LibreOfficeConverter:
    """Office to PDF converter using LibreOffice"""

    def __init__(self, soffice_path='soffice'):
        """
        Args:
            soffice_path: Path to LibreOffice executable
                         Windows: 'C:/Program Files/LibreOffice/program/soffice.exe'
                         Mac: '/Applications/LibreOffice.app/Contents/MacOS/soffice'
                         Linux: 'soffice' or 'libreoffice'
        """
        self.soffice_path = soffice_path

    def convert_to_pdf(self, input_path, output_dir=None):
        """
        Convert Office file to PDF using LibreOffice

        Args:
            input_path: Path to Office file
            output_dir: Output directory (defaults to same as input)

        Returns:
            Path to generated PDF file
        """
        input_path = Path(input_path)

        if output_dir is None:
            output_dir = input_path.parent

        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        # Build command
        cmd = [
            self.soffice_path,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', str(output_dir),
            str(input_path)
        ]

        # Execute conversion
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120  # 2 minute timeout
        )

        if result.returncode != 0:
            raise Exception(
                f"LibreOffice conversion failed: {result.stderr}"
            )

        # Calculate output path
        output_file = output_dir / f"{input_path.stem}.pdf"

        if not output_file.exists():
            raise Exception(f"PDF not created: {output_file}")

        return output_file
```

**Deployment Notes:**
- Bundle LibreOffice Portable (400MB)
- Or require LibreOffice installation
- Include in installer script
- Add to PATH

---

## Integration with PDF Locker Application

### Recommended UI Flow for Senior Users

```
┌─────────────────────────────────────┐
│  Step 1: Choose File Type           │
│                                      │
│  ☐ PDF file (already PDF)           │
│  ☐ Word document (.docx, .doc)      │
│  ☐ Excel spreadsheet (.xlsx, .xls)  │
│  ☐ PowerPoint (.pptx, .ppt)         │
│                                      │
│  [Next]                              │
└─────────────────────────────────────┘

┌─────────────────────────────────────┐
│  Step 2: Select File                 │
│                                      │
│  [Large File Selection Button]       │
│                                      │
│  Selected: report.docx               │
│                                      │
│  [Back] [Next]                       │
└─────────────────────────────────────┘

┌─────────────────────────────────────┐
│  Step 2.5: Converting... (if Office) │
│                                      │
│  Converting Word file to PDF...      │
│  [Progress Bar]                      │
│                                      │
│  This may take a moment...           │
└─────────────────────────────────────┘

┌─────────────────────────────────────┐
│  Step 3: Set Password                │
│  (Existing UI)                       │
└─────────────────────────────────────┘
```

### Code Integration Example

```python
# pdf_locker_with_office.py

class PDFLockerWithOfficeSupport(PDFLockerApp):
    """Extended PDF Locker with Office conversion support"""

    def __init__(self):
        super().__init__()

        # Initialize converter
        try:
            # Try Office-based conversion first
            from office_converter import OfficeConverter
            self.converter = OfficeConverter()
            self.conversion_method = "Microsoft Office"
        except:
            # Fall back to LibreOffice
            from libreoffice_converter import LibreOfficeConverter
            self.converter = LibreOfficeConverter()
            self.conversion_method = "LibreOffice"

    def _select_files(self):
        """Enhanced file selection with Office support"""
        files = filedialog.askopenfilenames(
            title="ファイルを選んでください",
            filetypes=[
                ("すべての対応ファイル", "*.pdf;*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt"),
                ("PDFファイル", "*.pdf"),
                ("Wordファイル", "*.docx;*.doc"),
                ("Excelファイル", "*.xlsx;*.xls"),
                ("PowerPointファイル", "*.pptx;*.ppt"),
            ]
        )

        if files:
            for file in files:
                # Convert Office files to PDF first
                if not file.lower().endswith('.pdf'):
                    file = self._convert_to_pdf_with_progress(file)

                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.file_listbox.insert(tk.END, Path(file).name)

    def _convert_to_pdf_with_progress(self, office_file):
        """Convert Office file to PDF with user feedback"""
        file_name = Path(office_file).name

        # Show progress dialog
        progress_dialog = tk.Toplevel(self.root)
        progress_dialog.title("変換中")
        progress_dialog.geometry("400x150")

        tk.Label(
            progress_dialog,
            text=f"{file_name} をPDFに変換しています...",
            font=("Yu Gothic UI", 14)
        ).pack(pady=20)

        progress_bar = ttk.Progressbar(
            progress_dialog,
            mode='indeterminate',
            length=300
        )
        progress_bar.pack(pady=10)
        progress_bar.start()

        status_label = tk.Label(
            progress_dialog,
            text="しばらくお待ちください",
            font=("Yu Gothic UI", 12)
        )
        status_label.pack(pady=10)

        # Convert in background thread
        pdf_path = None
        error = None

        def convert():
            nonlocal pdf_path, error
            try:
                temp_dir = Path.home() / "Desktop" / "一時変換ファイル"
                temp_dir.mkdir(exist_ok=True)

                pdf_path = self.converter.convert_to_pdf(
                    office_file,
                    temp_dir / f"{Path(office_file).stem}.pdf"
                )

            except Exception as e:
                error = str(e)

        import threading
        thread = threading.Thread(target=convert, daemon=True)
        thread.start()

        # Wait for completion
        def check_completion():
            if thread.is_alive():
                self.root.after(100, check_completion)
            else:
                progress_bar.stop()
                progress_dialog.destroy()

                if error:
                    messagebox.showerror(
                        "変換エラー",
                        f"ファイルをPDFに変換できませんでした:\n{error}\n\n"
                        f"元のファイル形式で続行できません。"
                    )
                else:
                    messagebox.showinfo(
                        "変換完了",
                        f"PDFに変換しました:\n{Path(pdf_path).name}"
                    )

        self.root.after(100, check_completion)

        return pdf_path if not error else None
```

---

## Deployment Considerations

### Windows Deployment (Recommended)

**Option 1: With Microsoft Office (Simplest)**
```bash
# requirements.txt
pypdf[crypto]>=4.0.0
docx2pdf>=0.8.1
comtypes>=1.4.0
tkinterdnd2>=0.3.0
pyinstaller>=6.0.0

# Build executable
python -m PyInstaller pdf_locker_with_office.py --onefile --windowed
```

**Package size:** ~40-50MB (small increase)
**Requirements:** Windows + Microsoft Office installed

**Option 2: With LibreOffice Portable**
```
Application_Folder/
├── PDF_Locker.exe
├── LibreOfficePortable/
│   └── App/
│       └── libreoffice/
│           └── program/
│               └── soffice.exe
```

**Package size:** ~450-500MB total
**Requirements:** Windows only (LibreOffice bundled)

### Testing Requirements

Before deployment:
1. ✅ Test with actual hospital documents
2. ✅ Test with user present (senior users)
3. ✅ Verify no dialog boxes appear
4. ✅ Test error handling
5. ✅ Check conversion quality
6. ✅ Test file size limits
7. ✅ Verify output folder creation
8. ✅ Test with Japanese filenames

---

## Risk Assessment for Hospital Environment

### Microsoft Office COM Approach (docx2pdf + comtypes)

**RISKS:**
- ⚠️ **Medium**: Microsoft doesn't officially support unattended automation
- ⚠️ **Low**: Dialog boxes could appear and confuse users
- ⚠️ **Low**: Office updates might affect stability
- ⚠️ **Very Low**: File locks if Office is already open

**MITIGATIONS:**
- ✅ Application runs with user present (not server)
- ✅ Clear error messages in Japanese
- ✅ User can close dialogs if they appear
- ✅ Tested thoroughly before deployment
- ✅ Fall back to manual PDF save instructions

**OVERALL RISK:** **LOW** for desktop application with user present

### LibreOffice Approach

**RISKS:**
- ⚠️ **Low**: Conversion quality differences
- ⚠️ **Low**: Formatting issues with complex documents
- ⚠️ **Very Low**: LibreOffice updates

**MITIGATIONS:**
- ✅ Bundle specific LibreOffice version
- ✅ Test with sample hospital documents
- ✅ Provide preview before password locking
- ✅ Allow user to regenerate if issues

**OVERALL RISK:** **VERY LOW**

---

## Final Recommendation

### For Your Hospital PDF Locker Application:

**Phase 1: MVP (Minimum Viable Product)**
- Start with **docx2pdf only** (Word documents)
- MIT licensed, simple, reliable
- 95% of hospital documents are likely Word
- Easy to implement and test

**Phase 2: Extended Support (if needed)**
- Add **comtypes** for Excel and PowerPoint
- Still MIT licensed
- Covers remaining Office formats

**Phase 3: Fallback (optional)**
- Add **LibreOffice** option for environments without Office
- MPL 2.0 licensed (free)
- Ensures broadest compatibility

### Implementation Priority

```python
# Recommended implementation order:

1. ✅ Enhance UI to detect Office file types
2. ✅ Add docx2pdf integration (Word only)
3. ✅ Test thoroughly with senior users
4. ✅ Deploy and gather feedback
5. ⏸️ Add Excel/PowerPoint support if requested
6. ⏸️ Add LibreOffice fallback if needed
```

### Why This Approach?

1. **Incremental**: Start simple, add features as needed
2. **Low Risk**: MIT licensed, tested libraries
3. **User-Friendly**: Works with existing Office installation
4. **Quality**: Native Office conversion = perfect fidelity
5. **Deployment**: Small footprint, easy to distribute
6. **Support**: Well-documented, active communities
7. **Cost**: Completely free
8. **Hospital-Appropriate**: Leverages existing IT infrastructure

---

## License Compliance Summary

All recommended solutions are **commercial-use friendly**:

| Component | License | Commercial OK | Attribution Required | Source Disclosure Required |
|-----------|---------|---------------|---------------------|---------------------------|
| docx2pdf | MIT | ✅ YES | ⚠️ YES (in docs) | ❌ NO |
| comtypes | MIT | ✅ YES | ⚠️ YES (in docs) | ❌ NO |
| LibreOffice | MPL 2.0 | ✅ YES | ⚠️ YES (if modified) | ⚠️ Only if modified |
| pypdf | MIT | ✅ YES | ⚠️ YES (in docs) | ❌ NO |

**To comply:**
1. Include license files in your distribution
2. Add acknowledgments in documentation
3. No need to disclose your own source code
4. Free to use, modify, and distribute

---

## Additional Resources

### Documentation Links
- [docx2pdf on PyPI](https://pypi.org/project/docx2pdf/)
- [comtypes documentation](https://pythonhosted.org/comtypes/)
- [LibreOffice Headless Conversion](https://ask.libreoffice.org/t/convert-files-to-pdf-a-on-command-line-headless-mode/821)
- [Microsoft Office COM Automation Warnings](https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa)

### Testing Resources
- Sample hospital document templates
- Japanese character encoding tests
- Senior user interface guidelines

---

## Questions for Stakeholders

Before implementation, confirm:

1. ✅ Is Microsoft Office installed on target hospital computers?
2. ✅ What Office version? (2016, 2019, 2021, 365?)
3. ✅ What file types are most common? (Word? Excel? PowerPoint?)
4. ✅ Average file sizes?
5. ✅ Internet connectivity available for installation?
6. ✅ IT department approval process?
7. ✅ User training requirements?

---

**Prepared by:** Claude Code AI Assistant
**Date:** January 9, 2026
**Status:** Research Complete - Ready for Implementation
