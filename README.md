# Multi-Format Text Extractor

A powerful text extraction tool that supports multiple file formats with both command-line and web interface options.

## Features

- **Multi-Format Support**: Extract text from PDF, Word, PowerPoint, and Excel files
- **Web Interface**: Beautiful, modern web UI for easy file upload and text extraction
- **Command Line**: Simple command-line tool for batch processing
- Extract text from all pages/slides/sheets of documents
- Save extracted text to a text file
- Preview of extracted text
- Error handling for common issues
- Progress tracking during extraction
- Drag and drop file upload
- Download extracted text files
- Format-specific metadata and statistics

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

3. Upload your file by:
   - Dragging and dropping the file onto the upload area
   - Clicking the upload area to browse for files

4. Click "Extract Text" to process your file

5. View the extracted text and download the text file

### Command Line Interface

#### Method 1: Run with default file
The script will automatically use available files in the current directory:

```bash
python pdf_extractor.py
```

#### Method 2: Specify a file
```bash
python pdf_extractor.py "path/to/your/file.pdf"
python pdf_extractor.py "path/to/your/document.docx"
python pdf_extractor.py "path/to/your/presentation.pptx"
python pdf_extractor.py "path/to/your/spreadsheet.xlsx"
```

## Web Interface Features

- **Modern UI**: Beautiful gradient design with smooth animations
- **Multi-Format Support**: Handle PDF, Word, PowerPoint, and Excel files
- **Drag & Drop**: Simply drag your file onto the upload area
- **Real-time Progress**: See the extraction progress in real-time
- **Format-Specific Statistics**: View relevant metadata for each file type
- **Text Preview**: Preview the first 1000 characters of extracted text
- **Download**: Download the complete extracted text as a .txt file
- **Copy to Clipboard**: Copy the extracted text directly to your clipboard
- **Multiple Files**: Process multiple files without refreshing the page

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
- Extracted text displayed in a preview area
- Format-specific statistics (pages, slides, sheets, etc.)
- Downloadable text file
- Copy to clipboard functionality

### Command Line
- Extracted text saved to `{original_filename}_extracted_text.txt`
- Format-specific metadata display
- Preview of first 500 characters in terminal
- Total character and word count displayed

## Example Output

### Web Interface
The web interface provides format-specific information:
- **PDF**: Pages, characters, words
- **Word**: Paragraphs, tables, characters, words
- **PowerPoint**: Slides, characters, words
- **Excel**: Sheets, characters, words

### Command Line
```
==================================================
FILE METADATA:
==================================================
File Type: Word Document
Characters: 15,420
Words: 2,584
Paragraphs: 45
Tables: 3
==================================================

TEXT PREVIEW (first 500 characters):
==================================================
[Extracted text preview here...]
==================================================

Total characters extracted: 15,420
Total words extracted: 2,584
```

## Error Handling

Both interfaces handle common errors:
- File not found
- Unsupported file formats
- Corrupted files
- Network errors (web interface)
- General exceptions

## File Management

- Uploaded files are automatically cleaned up after processing
- Extracted text files can be downloaded
- Temporary files are managed automatically
- Maximum file size: 32MB

## Requirements

- Python 3.6+
- PyPDF2 3.0.1
- Flask 2.3.3
- Werkzeug 2.3.7
- python-docx 0.8.11
- openpyxl 3.1.2
- python-pptx 0.6.21
- Pillow 10.0.1

## Security Features

- File type validation (supported formats only)
- Secure filename handling
- Maximum file size limits (32MB)
- Automatic file cleanup

## Browser Compatibility

The web interface works with all modern browsers:
- Chrome/Chromium
- Firefox
- Safari
- Edge

## Project Structure

```
Multi-Format Text Extractor/
├── app.py                 # Flask web application
├── pdf_extractor.py       # Command-line version
├── text_extractor.py      # Core text extraction module
├── requirements.txt       # Python dependencies
├── README.md             # Documentation
├── templates/
│   └── index.html        # Beautiful web interface
├── static/               # For any additional assets
└── uploads/              # Temporary file storage (auto-created)
```
