DocGen Engine

A Professional Desktop Application for Automated Document Generation from Excel Data and Word Templates.

DocGen Engine is a fast, intuitive, and powerful document automation platform built with Python, Tkinter, DocxTemplate, and Excel processing libraries. It allows you to generate personalized DOCX and PDF documents using Word templates and bulk Excel data — all through a clean, modern GUI.

🚀 Features

Template-Based Document Generation using .docx placeholders

Excel Data Integration with .xlsx and .xls support

Smart Auto-Mapping between template fields and Excel columns

DOCX + PDF Output (Microsoft Word required for PDF conversion)

Batch Processing for unlimited document creation

Custom File Naming Rules including mobile numbers and variables

Live Preview of mapped fields before generation

Modern UI using glass-morphism and responsive Tkinter widgets

Integrated Logging for debugging and error reporting

📋 System Requirements

Windows 10 or Windows 11

Microsoft Word installed (for PDF export)

Python 3.8+ (only required for Methods 2 & 3)

🔧 Installation & Usage Guide

Choose from three methods, based on your needs:

✔️ Method 1: Download EXE → Run directly (no Python required)

✔️ Method 2: Clone repo → Install modules → Run Python script

✔️ Method 3: Clone repo → Build your own EXE using PyInstaller

🟦 Method 1 — Download the EXE & Run (Recommended for Non-Developers)

This is the simplest method. No Python installation required.

Step 1 — Download the Executable

Go to the GitHub repository

Navigate to Releases

Download:
DocGen Engine.exe or DocGen-Engine.zip

Step 2 — Extract (if zipped)

Right-click → Extract All

Step 3 — Run the Application

Double-click:

DocGen Engine.exe

Step 4 — Install Missing Dependencies (Only if prompted)

Some systems may require the following packages once:

pip install pandas docxtpl python-docx pywin32 tkinter


After installation, simply re-run the EXE.

🟩 Method 2 — Clone Repository & Run Python Source (For Developers)

Perfect for debugging, editing UI, or extending the tool.

Step 1 — Clone Repo
git clone https://github.com/vjaykr/DocGen-Engine.git
cd DocGen-Engine

Step 2 — (Optional) Create Virtual Environment
python -m venv .venv
.venv\Scripts\activate

Step 3 — Install Required Python Dependencies

Either use the requirements file:

pip install -r requirements.txt


Or install manually:

pip install pandas docxtpl python-docx pywin32 tkinter

Step 4 — Run the Application
python launcher.py


Keep the terminal open to see live logs, warnings, and error messages.

🟧 Method 3 — Clone Repo & Build EXE from Scratch (For Packaging / Deployment)

Use this if you want to package and distribute your own executable.

Step 1 — Clone Repo
git clone https://github.com/vjaykr/DocGen-Engine.git
cd DocGen-Engine

Step 2 — Install Dependencies
pip install -r requirements.txt
pip install pyinstaller

Step 3 — Build EXE

You can either use the helper script:

exe_build.bat


Or run PyInstaller manually:

pyinstaller --noconsole --onefile --name "DocGen Engine" launcher.py

Step 4 — Locate the Executable

Find your generated EXE in:

dist/DocGen Engine.exe

Step 5 — Distribute & Use

Zip the entire dist/ folder before sharing.

📖 How to Use DocGen Engine
Step 1 — Prepare Your Word Template

Use placeholders in double curly brackets:

Dear {{name}},
Your order {{order_id}} is confirmed.

Step 2 — Prepare Your Excel File

Use clear column headers

Each row = one generated document

Step 3 — Run DocGen Engine

Open the application

Select:

Word Template

Excel Data File

Output Folder

Click Scan Placeholders

Review or edit automatic field mappings

Preview data

Click Generate DOCX & PDF

✔️ Resulting files appear in:

SaralWorks_DOCX/
SaralWorks_PDF/

💡 Example
Template (contract.docx)
EMPLOYMENT CONTRACT

Employee Name: {{employee_name}}
Position: {{position}}
Salary: ${{salary}}
Start Date: {{start_date}}
Department: {{department}}

Excel (employees.xlsx)
employee_name	position	salary	start_date	department
John Smith	Developer	75000	2024-01-15	IT
Jane Doe	Designer	65000	2024-02-01	Marketing
Output

2 DOCX documents

2 PDF documents

Fully personalized

⚙️ Advanced Features
AI-Like Auto-Mapping

Exact match

Space/underscore/hyphen variations

Partial substring matching

Rich File Naming Options

Use any column in filename

Auto-include mobile numbers

Custom patterns supported

Robust Error Handling

Dependency checks

Template field validation

Real-time status logs

🛠️ Tech Stack

Python 3.12

Tkinter (GUI)

pandas (Excel processing)

docxtpl / python-docx (template rendering)

pywin32 (MS Word automation)

PyInstaller (EXE bundling)

📚 Project Architecture
DocGen-Engine/
│
├── launcher.py        # App entry point + dependency checks
├── ui/                # Tkinter GUI components
├── engine/            # Core template engine & document logic
├── utils/             # Helper modules
├── requirements.txt
└── exe_build.bat      # EXE build script

🤝 Contributing

Fork the repo

Create a feature branch

git checkout -b feature/MyFeature


Commit & push

Open a PR

📝 License

This project is distributed under the MIT License.

🐛 Issues & Support

Report bugs → Issues tab

Request features → Issues tab

Documentation → Wiki

📊 Changelog
v1.0.0

Initial release

Template-based document automation

Excel data integration

Auto-field mapping

PDF conversion

Modern UI

❤️ Made with Love for Document Automation

Let me know if you want:
✅ A logo for the project
✅ A GitHub Wiki version
✅ A shorter public README + longer documentation PDF
✅ A setup installer (.msi/.exe) instead of PyInstaller

I'm happy to prepare those!