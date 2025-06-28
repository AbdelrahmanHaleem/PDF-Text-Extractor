import os
import sys
import glob
from text_extractor import TextExtractor
import json
from datetime import datetime

class BatchTextExtractor:
    """
    A batch text extractor that can process multiple files at once.
    """
    
    def __init__(self):
        self.extractor = TextExtractor()
        self.results = []
        self.failed_files = []
    
    def process_directory(self, directory_path, file_patterns=None):
        """
        Process all supported files in a directory.
        
        Args:
            directory_path (str): Path to the directory
            file_patterns (list): List of file patterns to match (e.g., ['*.pdf', '*.docx'])
        """
        if not os.path.exists(directory_path):
            print(f"Error: Directory not found: {directory_path}")
            return
        
        if file_patterns is None:
            file_patterns = ['*.pdf', '*.docx', '*.pptx', '*.xlsx', '*.xls']
        
        all_files = []
        for pattern in file_patterns:
            pattern_path = os.path.join(directory_path, pattern)
            all_files.extend(glob.glob(pattern_path))
        
        if not all_files:
            print(f"No supported files found in {directory_path}")
            return
        
        print(f"Found {len(all_files)} files to process:")
        for file in all_files:
            print(f"  - {os.path.basename(file)}")
        
        self.process_files(all_files)
    
    def process_files(self, file_paths):
        """
        Process a list of files.
        
        Args:
            file_paths (list): List of file paths to process
        """
        total_files = len(file_paths)
        print(f"\nProcessing {total_files} files...")
        
        for i, file_path in enumerate(file_paths, 1):
            print(f"\n[{i}/{total_files}] Processing: {os.path.basename(file_path)}")
            
            try:
                # Extract text
                extracted_text, metadata = self.extractor.extract_text(file_path)
                
                if extracted_text and metadata:
                    # Save extracted text
                    base_name = os.path.splitext(os.path.basename(file_path))[0]
                    output_path = f"{base_name}_extracted_text.txt"
                    
                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(extracted_text)
                    
                    # Store result
                    result = {
                        'file_path': file_path,
                        'file_name': os.path.basename(file_path),
                        'output_path': output_path,
                        'success': True,
                        'metadata': metadata,
                        'preview': extracted_text[:200] + "..." if len(extracted_text) > 200 else extracted_text
                    }
                    self.results.append(result)
                    
                    print(f"  ✅ Success: {metadata['characters']:,} characters, {metadata['words']:,} words")
                    
                else:
                    error_msg = metadata if metadata else "Unknown error"
                    self.failed_files.append({
                        'file_path': file_path,
                        'file_name': os.path.basename(file_path),
                        'error': error_msg
                    })
                    print(f"  ❌ Failed: {error_msg}")
                    
            except Exception as e:
                self.failed_files.append({
                    'file_path': file_path,
                    'file_name': os.path.basename(file_path),
                    'error': str(e)
                })
                print(f"  ❌ Error: {str(e)}")
        
        self.generate_summary_report()
    
    def generate_summary_report(self):
        """
        Generate a summary report of the batch processing.
        """
        print("\n" + "="*60)
        print("BATCH PROCESSING SUMMARY")
        print("="*60)
        
        # Success summary
        if self.results:
            print(f"\n✅ Successfully processed: {len(self.results)} files")
            total_chars = sum(r['metadata']['characters'] for r in self.results)
            total_words = sum(r['metadata']['words'] for r in self.results)
            print(f"   Total characters extracted: {total_chars:,}")
            print(f"   Total words extracted: {total_words:,}")
            
            # Format breakdown
            format_counts = {}
            for result in self.results:
                file_type = result['metadata']['file_type']
                format_counts[file_type] = format_counts.get(file_type, 0) + 1
            
            print("\n   Format breakdown:")
            for file_type, count in format_counts.items():
                print(f"     {file_type}: {count} files")
        
        # Failure summary
        if self.failed_files:
            print(f"\n❌ Failed to process: {len(self.failed_files)} files")
            for failed in self.failed_files:
                print(f"   {failed['file_name']}: {failed['error']}")
        
        # Generate detailed report
        self.save_detailed_report()
    
    def save_detailed_report(self):
        """
        Save a detailed JSON report of the batch processing.
        """
        report = {
            'timestamp': datetime.now().isoformat(),
            'summary': {
                'total_processed': len(self.results) + len(self.failed_files),
                'successful': len(self.results),
                'failed': len(self.failed_files)
            },
            'successful_files': self.results,
            'failed_files': self.failed_files
        }
        
        report_filename = f"batch_extraction_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        try:
            with open(report_filename, 'w', encoding='utf-8') as f:
                json.dump(report, f, indent=2, ensure_ascii=False)
            print(f"\n📊 Detailed report saved: {report_filename}")
        except Exception as e:
            print(f"\n⚠️  Could not save detailed report: {str(e)}")

def main():
    """
    Main function for batch text extraction.
    """
    if len(sys.argv) < 2:
        print("Batch Text Extractor")
        print("=" * 50)
        print("Usage:")
        print("  python batch_extractor.py <directory_path>")
        print("  python batch_extractor.py <file1> <file2> <file3> ...")
        print("\nExamples:")
        print("  python batch_extractor.py ./documents")
        print("  python batch_extractor.py file1.pdf file2.docx file3.pptx")
        print("\nSupported formats: PDF, DOCX, PPTX, XLSX, XLS")
        return
    
    batch_extractor = BatchTextExtractor()
    
    if len(sys.argv) == 2:
        # Single argument - treat as directory
        path = sys.argv[1]
        if os.path.isdir(path):
            batch_extractor.process_directory(path)
        elif os.path.isfile(path):
            batch_extractor.process_files([path])
        else:
            print(f"Error: Path not found: {path}")
    else:
        # Multiple arguments - treat as file list
        file_paths = sys.argv[1:]
        # Filter out non-existent files
        existing_files = [f for f in file_paths if os.path.exists(f)]
        non_existing_files = [f for f in file_paths if not os.path.exists(f)]
        
        if non_existing_files:
            print("Warning: The following files were not found:")
            for f in non_existing_files:
                print(f"  - {f}")
            print()
        
        if existing_files:
            batch_extractor.process_files(existing_files)
        else:
            print("No existing files found to process.")

if __name__ == "__main__":
    main() 