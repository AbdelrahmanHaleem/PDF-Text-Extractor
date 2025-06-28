import os
import sys
from text_extractor import TextExtractor

def extract_text_from_file(file_path):
    """
    Extract text from a file using the TextExtractor class.
    
    Args:
        file_path (str): Path to the file
        
    Returns:
        tuple: (extracted_text, metadata_dict) or (None, error_message)
    """
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        # Initialize text extractor
        extractor = TextExtractor()
        
        # Check if file format is supported
        if not extractor.is_supported(file_path):
            supported_formats = ', '.join(extractor.get_supported_formats())
            raise ValueError(f"Unsupported file format. Supported formats: {supported_formats}")
        
        # Extract text
        extracted_text, result = extractor.extract_text(file_path)
        
        if extracted_text is None:
            return None, result
        
        return extracted_text, result
            
    except FileNotFoundError as e:
        print(f"Error: {e}")
        return None, None
    except ValueError as e:
        print(f"Error: {e}")
        return None, None
    except Exception as e:
        print(f"Unexpected error: {e}")
        return None, None

def save_text_to_file(text, output_path):
    """
    Save extracted text to a file.
    
    Args:
        text (str): Text to save
        output_path (str): Path for the output file
    """
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"Text saved to: {output_path}")
    except Exception as e:
        print(f"Error saving text to file: {e}")

def display_metadata(metadata):
    """
    Display file metadata in a formatted way.
    
    Args:
        metadata (dict): Metadata dictionary
    """
    print("\n" + "="*50)
    print("FILE METADATA:")
    print("="*50)
    print(f"File Type: {metadata['file_type']}")
    print(f"Characters: {metadata['characters']:,}")
    print(f"Words: {metadata['words']:,}")
    
    # Display format-specific metadata
    if metadata['file_type'] == 'PDF':
        print(f"Pages: {metadata['pages']}")
    elif metadata['file_type'] == 'Word Document':
        print(f"Paragraphs: {metadata['paragraphs']}")
        print(f"Tables: {metadata['tables']}")
    elif metadata['file_type'] == 'PowerPoint Presentation':
        print(f"Slides: {metadata['slides']}")
    elif metadata['file_type'] == 'Excel Spreadsheet':
        print(f"Sheets: {metadata['sheets']}")
    
    print("="*50)

def main():
    """
    Main function to run the text extractor.
    """
    # Check if file path is provided as command line argument
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        # Use the PDF file in the current directory if it exists
        default_files = [
            "Functional Requirements Document Phase 1.pdf",
            "Copy of User_Stories (1)(1).xlsx"
        ]
        
        file_path = None
        for default_file in default_files:
            if os.path.exists(default_file):
                file_path = default_file
                break
        
        if file_path is None:
            print("No file specified and no default files found.")
            print("Usage: python text_extractor.py <file_path>")
            print("Supported formats: PDF, DOCX, PPTX, XLSX, XLS")
            return
    
    print(f"Extracting text from: {file_path}")
    
    # Extract text from file
    extracted_text, metadata = extract_text_from_file(file_path)
    
    if extracted_text and metadata:
        # Create output filename
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = f"{base_name}_extracted_text.txt"
        
        # Save text to file
        save_text_to_file(extracted_text, output_path)
        
        # Display metadata
        display_metadata(metadata)
        
        # Print first 500 characters as preview
        print("\n" + "="*50)
        print("TEXT PREVIEW (first 500 characters):")
        print("="*50)
        print(extracted_text[:500] + "..." if len(extracted_text) > 500 else extracted_text)
        print("="*50)
        
        print(f"\nTotal characters extracted: {metadata['characters']:,}")
        print(f"Total words extracted: {metadata['words']:,}")
    else:
        print("Failed to extract text from file.")

if __name__ == "__main__":
    main() 