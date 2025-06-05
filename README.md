# ShugeZhuanZhuan  - Format Converter

## Project Overview

`ShugeZhuanZhuan` (which translates to "Book Grid Converter" or "Document Converter") is a powerful, Python-based Graphical User Interface (GUI) tool designed for versatile document and e-book format conversions. I developed this utility out of frustration, as I couldn't find a comprehensive, free solution online that handles the specific range of formats I needed.

## Key Features & Highlights

* **Extensive Format Support:** Supports conversions between a wide array of popular document and e-book formats including:
    * **PDF**
    * **DOCX (Microsoft Word)**
    * **EPUB**
    * **AZW3**
    * **MOBI**
* **Robust Batch Processing:** Capable of processing dozens, or even hundreds, of files in a single operation, significantly boosting your conversion efficiency.
* **Multi-Target Conversion (Conditional):**
    * If all input files are of the **same original format**, you can convert them simultaneously to **multiple selected target formats** (e.g., convert 100 PDFs to both DOCX and EPUB in one go).
    * If your input files include **mixed original formats** (e.g., a batch containing both PDF and DOCX files), you must select **only one target format** for the entire batch (e.g., convert all mixed files to DOCX).
* **Intuitive Drag-and-Drop Interface:** Simplifies file selection with drag-and-drop functionality, making the conversion process straightforward.
* **Addresses Market Gaps:** Fills a niche for a free and comprehensive tool capable of handling these specific format conversions efficiently.

## Important Note

**This repository provides the source code only.**

`ShugeZhuanZhuan` is an open-source project. This repository does not provide pre-compiled executable software. Users are expected to have a Python environment set up and follow the instructions below to run the application from source.

## Getting Started (How to Run the Code)

### 1. Prerequisites (Windows)

Ensure you have **Python 3.8 or newer** installed on your system. It's highly recommended to **check "Add python.exe to PATH"** during Python installation for easier command-line access.

Additionally, for full conversion functionality, you might need these external applications:

* **Microsoft Word:** Required for accurate DOCX to PDF conversions, as the tool leverages Word's COM automation.
* **Calibre:** For all e-book format conversions (EPUB, AZW3, MOBI, and also PDF/DOCX to/from these formats). `ShugeZhuanZhuan` utilizes Calibre's powerful `ebook-convert` command-line tool. After installing Calibre, **ensure that its installation directory (specifically the `Calibre2` folder containing `ebook-convert.exe`) is added to your system's `PATH` environment variable.**

### 2. Download the Source Code

You can obtain the project's source code using one of these methods:

* **Using Git (Recommended):**
    ```bash
    git clone [https://github.com/YOUR_USERNAME/FormatConverter.git](https://github.com/YOUR_USERNAME/FormatConverter.git)
    cd FormatConverter
    ```
    (Replace `YOUR_USERNAME` with your actual GitHub username)
* **Direct ZIP Download:** Click the green "Code" button on the GitHub repository page, then select "Download ZIP". Extract the contents to your desired folder.

### 3. Install Python Dependencies

Open your Command Prompt (CMD) or PowerShell. Navigate to the folder where you downloaded the source code (e.g., `F:\ai\agent 2025.2.22\格式转换`). Then, run the following command to install all necessary Python libraries:

```bash
pip install tkinterdnd2 pdf2docx pywin32 docx2pdf
4. Run the Application
Once all dependencies are successfully installed, you can launch the ShugeZhuanZhuan application. In your Command Prompt (CMD) or PowerShell, navigate to the code directory and execute:

Bash

python FormatConverter.py
The graphical user interface (GUI) of the application should then appear.

Technical Stack
Python
Tkinter / TkinterDnD2 (for the GUI framework)
pdf2docx (for PDF to DOCX conversion)
pywin32 (for Windows-specific operations and DOCX to PDF conversion via Microsoft Word COM automation)
docx2pdf (for DOCX to PDF conversion, typically leveraging MS Word)
Calibre's ebook-convert command (for comprehensive e-book format conversions)
Contribution
If you encounter any bugs, have feature suggestions, or wish to contribute improvements, feel free to open an Issue or submit a Pull Request on this repository. Your contributions are welcome!

License
This project is open-sourced under the MIT License.
