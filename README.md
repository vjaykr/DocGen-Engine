# DocGen Engine

A professional desktop application for automated document generation from Excel data and Word templates. Built with Python and Tkinter, featuring an intuitive GUI for seamless document processing and PDF conversion.



## 🚀 Features

- **Template-Based Document Generation**: Use Word (.docx) templates with placeholder fields
- **Excel Data Integration**: Import data from Excel files (.xlsx, .xls) for bulk document creation
- **Intelligent Field Mapping**: Auto-map template placeholders to Excel columns with smart matching
- **Dual Output Format**: Generate both DOCX and PDF versions of documents
- **Batch Processing**: Process multiple records simultaneously
- **Custom File Naming**: Include mobile numbers or custom naming patterns
- **Real-time Preview**: Preview data mapping before generation
- **Professional UI**: Modern, responsive interface with glass morphism design
!
![DocGen Engine Interface](image.png)
## 📋 Requirements

- Windows 10/11
- Microsoft Word (for PDF conversion)
- Python 3.8+ (for development)

## 🔧 Installation

### Option 1: Download Executable (Recommended)
1. Download `DocGen Engine.exe` directly from this repository
2. Run:
   ```bash
   pip install pandas docxtpl python-docx pywin32 tkinter
   ```
3. Run the executable - it will automatically install required dependencies
4. Start using the application immediately

### Option 2: Build from Source
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/docgen-engine.git
   cd docgen-engine
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python launcher.py
   ```

4. Build executable (optional):
   ```bash
   exe_build.bat
   ```

## 📖 How to Use

### Step 1: Prepare Your Template
1. Create a Word document (.docx) with placeholder fields
2. Use double curly braces for placeholders: `{{field_name}}`
3. Example: `Dear {{name}}, your order {{order_id}} is ready.`

### Step 2: Prepare Your Data
1. Create an Excel file (.xlsx/.xls) with your data
2. Use column headers that match or relate to your template placeholders
3. Each row represents one document to be generated

### Step 3: Generate Documents
1. **Launch Application**: Run `DocGen Engine.exe`
2. **Select Files**:
   - Choose your Word template
   - Select your Excel data file
   - Set output folder (optional)
3. **Scan Placeholders**: Click "🔍 Scan Placeholders" to detect template fields
4. **Map Fields**: 
   - Use "✨ Auto Map Fields" for automatic mapping
   - Or manually map each placeholder to Excel columns
5. **Preview**: Click "👁️ Preview Data" to verify mappings
6. **Generate**: Click "🚀 Generate DOCX & PDF Files"

### Step 4: Access Results
- Find generated files in your output folder:
  - `SaralWorks_DOCX/` - Word documents
  - `SaralWorks_PDF/` - PDF documents

## 💡 Example Usage

### Template (contract.docx):
```
EMPLOYMENT CONTRACT

Employee Name: {{employee_name}}
Position: {{position}}
Salary: ${{salary}}
Start Date: {{start_date}}
Department: {{department}}
```

### Excel Data (employees.xlsx):
| employee_name | position | salary | start_date | department |
|---------------|----------|--------|------------|------------|
| John Smith | Developer | 75000 | 2024-01-15 | IT |
| Jane Doe | Designer | 65000 | 2024-02-01 | Marketing |

### Result:
- 2 DOCX files with personalized contracts
- 2 PDF files ready for distribution

## ⚙️ Advanced Features

### Auto-Mapping Intelligence
The application uses smart matching algorithms:
- **Exact Match**: Direct column name matching
- **Variation Handling**: Handles spaces, underscores, hyphens
- **Partial Matching**: Finds related fields using substring matching

### File Naming Options
- **Default**: Uses first column value as filename
- **Mobile Integration**: Includes mobile numbers in filenames
- **Custom Patterns**: Supports various naming conventions

### Error Handling
- **Dependency Management**: Auto-installs missing Python packages
- **File Validation**: Checks template and data file integrity
- **Progress Tracking**: Real-time status updates during generation

## 🛠️ Technical Details

### Built With
- **Python 3.12**: Core application logic
- **Tkinter**: Modern GUI framework
- **pandas**: Excel data processing
- **python-docx**: Word document manipulation
- **python-docx-template**: Template rendering
- **pywin32**: Microsoft Office integration

### Architecture
- **Launcher**: Dependency checking and installation
- **Main App**: Core application with GUI
- **Template Engine**: Document generation logic
- **File Handlers**: Excel and Word file processing

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🐛 Issues & Support

- **Bug Reports**: [Create an issue](../../issues)
- **Feature Requests**: [Request a feature](../../issues)
- **Documentation**: Check the [Wiki](../../wiki)

## 📊 Changelog

### v1.0.0
- Initial release
- Template-based document generation
- Excel data integration
- Auto-mapping functionality
- PDF conversion support
- Professional UI design

## 🙏 Acknowledgments

- Built for efficient document automation
- Designed for business and professional use
- Optimized for Windows environments
- Community-driven development

---

**Made with ❤️ for document automation**