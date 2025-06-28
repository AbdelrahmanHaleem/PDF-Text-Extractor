from flask import Flask, render_template, request, jsonify, send_file
import os
from werkzeug.utils import secure_filename
import tempfile
import uuid
from text_extractor import TextExtractor

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = 'your-secret-key-here'

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Initialize text extractor
text_extractor = TextExtractor()

ALLOWED_EXTENSIONS = set(text_extractor.get_supported_formats())

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in [ext[1:] for ext in ALLOWED_EXTENSIONS]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file selected'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if file and allowed_file(file.filename):
        # Create unique filename
        unique_id = str(uuid.uuid4())
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{unique_id}_{filename}")
        
        try:
            file.save(file_path)
            
            # Extract text from file
            extracted_text, result = text_extractor.extract_text(file_path)
            
            if extracted_text is None:
                # Clean up file on error
                os.remove(file_path)
                return jsonify({'error': f'Failed to extract text: {result}'}), 400
            
            # Save extracted text to file
            text_file_path = file_path.replace(os.path.splitext(file_path)[1], '_extracted.txt')
            with open(text_file_path, 'w', encoding='utf-8') as f:
                f.write(extracted_text)
            
            # Prepare response with metadata
            response_data = {
                'success': True,
                'text': extracted_text,
                'characters': result['characters'],
                'words': result['words'],
                'text_file': text_file_path,
                'original_file': file_path,
                'file_type': result['file_type']
            }
            
            # Add format-specific metadata
            if result['file_type'] == 'PDF':
                response_data['pages'] = result['pages']
            elif result['file_type'] == 'Word Document':
                response_data['paragraphs'] = result['paragraphs']
                response_data['tables'] = result['tables']
            elif result['file_type'] == 'PowerPoint Presentation':
                response_data['slides'] = result['slides']
            elif result['file_type'] == 'Excel Spreadsheet':
                response_data['sheets'] = result['sheets']
            
            return jsonify(response_data)
            
        except Exception as e:
            # Clean up file on error
            if os.path.exists(file_path):
                os.remove(file_path)
            return jsonify({'error': f'Error processing file: {str(e)}'}), 500
    
    supported_formats = ', '.join([ext[1:].upper() for ext in ALLOWED_EXTENSIONS])
    return jsonify({'error': f'Invalid file type. Please upload: {supported_formats}'}), 400

@app.route('/download/<path:filename>')
def download_file(filename):
    try:
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return jsonify({'error': f'File not found: {str(e)}'}), 404

@app.route('/cleanup', methods=['POST'])
def cleanup_files():
    """Clean up uploaded files"""
    try:
        data = request.get_json()
        files_to_delete = data.get('files', [])
        
        for file_path in files_to_delete:
            if os.path.exists(file_path):
                os.remove(file_path)
        
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/supported-formats')
def get_supported_formats():
    """Get list of supported file formats"""
    formats = text_extractor.get_supported_formats()
    return jsonify({
        'formats': [ext[1:].upper() for ext in formats],
        'extensions': formats
    })

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000) 