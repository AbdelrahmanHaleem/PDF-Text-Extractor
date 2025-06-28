# Multi-Format Text Extractor - Batch Processing

A powerful text extraction tool that supports multiple file formats with both command-line and web interface options, including advanced batch processing capabilities.

## Features

- **Multi-Format Support**: Extract text from PDF, Word, PowerPoint, and Excel files
- **Batch Processing**: Process multiple files simultaneously
- **Web Interface**: Beautiful, modern web UI for easy file upload and text extraction
- **Command Line**: Simple command-line tool for batch processing
- Extract text from all pages/slides/sheets of documents
- Save extracted text to individual files or ZIP archives
- Preview of extracted text
- Error handling for common issues
- Progress tracking during extraction
- Drag and drop file upload
- Download extracted text files individually or as batch
- Format-specific metadata and statistics
- Detailed batch processing reports

## Supported File Formats

- **PDF Files** (.pdf) - Extract text from all pages
- **Word Documents** (.docx) - Extract text from paragraphs and tables
- **PowerPoint Presentations** (.pptx) - Extract text from slides and tables
- **Excel Spreadsheets** (.xlsx, .xls) - Extract data from all sheets

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

Or install dependencies individually:
```bash
pip install PyPDF2 Flask Werkzeug python-docx openpyxl python-pptx Pillow
```

## Usage

### Web Interface (Recommended)

1. Start the web server:
```bash
python app.py
```

2. Open your browser and go to: `http://localhost:5000`

3. Upload multiple files by:
   - Dragging and dropping multiple files onto the upload area
   - Clicking the upload area to browse and select multiple files
   - Files are displayed in a list with options to remove individual files

4. Click "Extract Text from All Files" to process all files

5. View batch results with individual file statistics and download options

### Command Line Interface

#### Single File Processing
```bash
python pdf_extractor.py "path/to/your/file.pdf"
python pdf_extractor.py "path/to/your/document.docx"
python pdf_extractor.py "path/to/your/presentation.pptx"
python pdf_extractor.py "path/to/your/spreadsheet.xlsx"
```

#### Batch Processing - Multiple Files
```bash
python batch_extractor.py file1.pdf file2.docx file3.pptx file4.xlsx
```

#### Batch Processing - Directory
```bash
python batch_extractor.py ./documents
```

#### Batch Processing - Specific File Types
```bash
python batch_extractor.py ./documents --patterns "*.pdf" "*.docx"
```

## Web Interface Features

- **Modern UI**: Beautiful gradient design with smooth animations
- **Multi-Format Support**: Handle PDF, Word, PowerPoint, and Excel files
- **Batch Upload**: Select and process multiple files at once
- **Drag & Drop**: Simply drag multiple files onto the upload area
- **File Management**: View selected files, remove individual files, clear all
- **Real-time Progress**: See the extraction progress in real-time
- **Batch Summary**: Overview of successful and failed files
- **Individual Results**: Detailed view of each processed file
- **Format-Specific Statistics**: View relevant metadata for each file type
- **Text Preview**: Preview the first 300 characters of extracted text
- **Download Options**: Download individual files or all as ZIP archive
- **Error Handling**: Clear indication of failed files with error messages

## Command Line Features

### Single File Processing (`pdf_extractor.py`)
- Process one file at a time
- Display format-specific metadata
- Save extracted text to individual files
- Preview of extracted text in terminal

### Batch Processing (`batch_extractor.py`)
- Process multiple files or entire directories
- Detailed progress reporting
- Comprehensive summary statistics
- JSON report generation
- Error handling for individual files
- Format breakdown statistics

## Format-Specific Features

### PDF Files
- Extract text from all pages
- Page count statistics
- Character and word count

### Word Documents (.docx)
- Extract text from paragraphs
- Extract text from tables
- Paragraph and table count statistics
- Character and word count

### PowerPoint Presentations (.pptx)
- Extract text from all slides
- Extract text from tables within slides
- Slide count statistics
- Character and word count

### Excel Spreadsheets (.xlsx, .xls)
- Extract data from all sheets
- Tabular data preservation
- Sheet count statistics
- Character and word count

## Output

### Web Interface
- **Batch Summary**: Total files, successful, failed
- **Individual Results**: Each file with its own statistics and preview
- **Download Options**: Individual files or ZIP archive
- **Error Reporting**: Clear indication of failed files

### Command Line
- **Single File**: Extracted text saved to `{original_filename}_extracted_text.txt`
- **Batch Processing**: Multiple files with detailed report
- **JSON Report**: Comprehensive batch processing report
- **Format-specific metadata display**
- **Preview of first 500 characters in terminal**
- **Total character and word count displayed**

## Example Output

### Web Interface Batch Results
```
Batch Processing Complete
┌─────────────┬─────────────┬─────────────┐
│ Total Files │ Successful  │   Failed    │
│     5       │     4       │     1       │
└─────────────┴─────────────┴─────────────┘

Individual File Results:
├─ document1.pdf (PDF, 3 pages, 1,250 chars)
├─ document2.docx (Word, 2 paragraphs, 890 chars)
├─ presentation.pptx (PowerPoint, 5 slides, 2,100 chars)
├─ spreadsheet.xlsx (Excel, 2 sheets, 3,450 chars)
└─ failed_file.txt (Failed: Unsupported format)
```

### Command Line Batch Processing
```
============================================================
BATCH PROCESSING SUMMARY
============================================================

✅ Successfully processed: 4 files
   Total characters extracted: 7,690
   Total words extracted: 1,284

   Format breakdown:
     PDF: 1 files
     Word Document: 1 files
     PowerPoint Presentation: 1 files
     Excel Spreadsheet: 1 files

❌ Failed to process: 1 files
   failed_file.txt: Unsupported file format

📊 Detailed report saved: batch_extraction_report_20250628_143022.json
```

## Error Handling

Both interfaces handle common errors:
- File not found
- Unsupported file formats
- Corrupted files
- Network errors (web interface)
- Individual file failures in batch processing
- General exceptions

## File Management

- Uploaded files are automatically cleaned up after processing
- Extracted text files can be downloaded individually or as batch
- Temporary files are managed automatically
- Maximum file size: 64MB (increased for batch processing)
- ZIP archive creation for batch downloads

## Requirements

- Python 3.6+
- PyPDF2 3.0.1
- Flask 2.3.3
- Werkzeug 2.3.7
- python-docx >=1.1.2
- openpyxl >=3.1.5
- python-pptx >=1.0.2

## Security Features

- File type validation (supported formats only)
- Secure filename handling
- Maximum file size limits (64MB)
- Automatic file cleanup
- Batch processing safety checks

## Browser Compatibility

The web interface works with all modern browsers:
- Chrome/Chromium
- Firefox
- Safari
- Edge

## Project Structure

```
Multi-Format Text Extractor/
├── app.py                 # Flask web application with batch support
├── pdf_extractor.py       # Single file command-line version
├── batch_extractor.py     # Batch processing command-line version
├── text_extractor.py      # Core text extraction module
├── requirements.txt       # Python dependencies
├── README.md             # Documentation
├── templates/
│   └── index.html        # Beautiful web interface with batch support
├── static/               # For any additional assets
└── uploads/              # Temporary file storage (auto-created)
```

## Advanced Usage

### Batch Processing with Custom Patterns
```bash
# Process only PDF files
python batch_extractor.py ./documents --patterns "*.pdf"

# Process specific file types
python batch_extractor.py ./documents --patterns "*.docx" "*.pptx"
```

### Web Interface Batch Features
- **Selective Download**: Download individual files or entire batch as ZIP
- **File Preview**: See file list before processing
- **Remove Files**: Remove individual files from batch before processing
- **Clear All**: Reset and start with new files
- **Progress Tracking**: Real-time progress for batch processing
