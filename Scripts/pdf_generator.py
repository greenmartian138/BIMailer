"""
PDF generation functionality for BIMailer automation system.
Handles converting PNG files to PDFs with original dimensions.
"""

from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import datetime
from utils import BIMailerLogger, get_png_files, get_file_creation_date, clean_filename
from config_manager import ConfigManager


class PDFGenerator:
    """Handles PDF generation from PNG files."""
    
    def __init__(self, config_manager: ConfigManager):
        self.config = config_manager
        self.logger = BIMailerLogger()
        
        # Directory paths - handle running from Scripts directory
        base_path = Path("..") if not Path("Input").exists() and Path("../Input").exists() else Path(".")
        
        self.input_dir = base_path / "Input"
        self.output_dir = base_path / "Output/PDFs"
        self.all_folder = self.input_dir / "ALL"
        
        # Ensure output directory exists
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def generate_pdf_for_folder(self, folder_name: str) -> Optional[Path]:
        """Generate PDF for a specific folder."""
        try:
            # Get PDF name from configuration
            pdf_name = self.config.get_pdf_name_for_folder(folder_name)
            if not pdf_name:
                self.logger.log_error(f"No PDF name configured for folder: {folder_name}")
                return None
            
            folder_path = self.input_dir / folder_name
            if not folder_path.exists():
                self.logger.log_error(f"Folder does not exist: {folder_path}")
                return None
            
            # Get PNG files from the specific folder
            folder_pngs = get_png_files(folder_path, self.config.processing_config.png_file_extensions)
            
            # Get PNG files from ALL folder (to be included in every PDF)
            all_pngs = get_png_files(self.all_folder, self.config.processing_config.png_file_extensions)
            
            # Combine PNG files, avoiding duplicates based on filename
            # Create a dictionary to track files by name, prioritizing ALL folder files
            unique_files = {}
            
            # Add ALL folder files first (these take priority)
            for png_file in all_pngs:
                unique_files[png_file.name] = png_file
            
            # Add folder-specific files only if they don't already exist
            for png_file in folder_pngs:
                if png_file.name not in unique_files:
                    unique_files[png_file.name] = png_file
            
            # Convert back to list, maintaining order (ALL files first, then unique folder files)
            all_png_files = list(unique_files.values())
            
            if not all_png_files:
                self.logger.log_file_operation(f"No PNG files found for folder: {folder_name}")
                return None
            
            # Generate PDF
            pdf_path = self._create_pdf(pdf_name, all_png_files, folder_name)
            
            if pdf_path:
                # Calculate actual counts after deduplication
                all_files_used = len([f for f in all_png_files if f in all_pngs])
                folder_files_used = len(all_png_files) - all_files_used
                
                self.logger.log_file_operation(
                    f"PDF generated successfully: {pdf_path} "
                    f"({all_files_used} ALL files + {folder_files_used} {folder_name} files, {len(all_png_files)} total unique files)"
                )
            
            return pdf_path
            
        except Exception as e:
            self.logger.log_error(f"Failed to generate PDF for folder {folder_name}", e)
            return None
    
    def _create_pdf(self, pdf_name: str, png_files: List[Path], folder_name: str) -> Optional[Path]:
        """Create PDF from list of PNG files."""
        try:
            # Clean PDF name for filename
            clean_pdf_name = clean_filename(pdf_name)
            timestamp = datetime.datetime.now().strftime(self.config.processing_config.timestamp_format)
            pdf_filename = f"{clean_pdf_name}_{timestamp}.pdf"
            pdf_path = self.output_dir / pdf_filename
            
            # Create PDF
            c = canvas.Canvas(str(pdf_path))
            
            file_info_list = []
            
            for png_file in png_files:
                try:
                    # Get image dimensions and DPI information
                    with Image.open(png_file) as img:
                        img_width, img_height = img.size
                        
                        # Get actual DPI from image metadata, fallback to 96 if not available
                        dpi = img.info.get('dpi', (96, 96))
                        if isinstance(dpi, tuple):
                            dpi_x, dpi_y = dpi
                        else:
                            dpi_x = dpi_y = dpi
                        
                        # Use the higher DPI value for better quality
                        actual_dpi = max(dpi_x, dpi_y)
                        
                        # For high-resolution images, maintain quality by using proper scaling
                        if actual_dpi > 150:
                            # High-res image: use actual DPI for precise scaling
                            page_width = img_width * 72 / actual_dpi
                            page_height = img_height * 72 / actual_dpi
                        else:
                            # Standard resolution: use 96 DPI assumption but with better scaling
                            page_width = img_width * 72 / 96
                            page_height = img_height * 72 / 96
                    
                    # Set page size to match image
                    c.setPageSize((page_width, page_height))
                    
                    # Create ImageReader for better quality control
                    img_reader = ImageReader(str(png_file))
                    
                    # Add image to PDF with maximum quality settings
                    # Use ImageReader for better compression control
                    c.drawImage(img_reader, 0, 0, 
                              width=page_width, height=page_height,
                              preserveAspectRatio=True, anchor='c')
                    
                    # Get file creation date
                    creation_date = get_file_creation_date(png_file)
                    file_info_list.append({
                        'filename': png_file.name,
                        'creation_date': creation_date,
                        'size': png_file.stat().st_size,
                        'dimensions': f"{img_width}x{img_height}"
                    })
                    
                    # Finish the page
                    c.showPage()
                    
                    self.logger.log_file_operation(
                        f"Added to PDF: {png_file.name} ({img_width}x{img_height})"
                    )
                    
                except Exception as e:
                    self.logger.log_error(f"Failed to process PNG file: {png_file}", e)
                    continue
            
            # Save PDF
            c.save()
            
            # Log PDF creation details
            pdf_size = pdf_path.stat().st_size
            self.logger.log_file_operation(
                f"PDF created: {pdf_path.name} "
                f"(Size: {pdf_size / 1024 / 1024:.2f} MB, Pages: {len(file_info_list)})"
            )
            
            # Store file information for email template
            self._store_pdf_metadata(pdf_path, pdf_name, folder_name, file_info_list)
            
            return pdf_path
            
        except Exception as e:
            self.logger.log_error(f"Failed to create PDF: {pdf_name}", e)
            return None
    
    def _store_pdf_metadata(self, pdf_path: Path, pdf_name: str, folder_name: str, file_info_list: List[Dict]):
        """Store PDF metadata for use in email templates."""
        try:
            metadata = {
                'pdf_path': str(pdf_path),
                'pdf_name': pdf_name,
                'folder_name': folder_name,
                'creation_timestamp': datetime.datetime.now().isoformat(),
                'file_count': len(file_info_list),
                'files': file_info_list
            }
            
            # Store metadata in a JSON file alongside the PDF
            metadata_path = pdf_path.with_suffix('.json')
            import json
            with open(metadata_path, 'w') as f:
                json.dump(metadata, f, indent=2, default=str)
            
            self.logger.log_file_operation(f"PDF metadata stored: {metadata_path}")
            
        except Exception as e:
            self.logger.log_error(f"Failed to store PDF metadata for {pdf_path}", e)
    
    def get_pdf_metadata(self, pdf_path: Path) -> Optional[Dict]:
        """Retrieve PDF metadata."""
        try:
            metadata_path = pdf_path.with_suffix('.json')
            if metadata_path.exists():
                import json
                with open(metadata_path, 'r') as f:
                    return json.load(f)
            return None
        except Exception as e:
            self.logger.log_error(f"Failed to load PDF metadata for {pdf_path}", e)
            return None
    
    def generate_pdfs_for_all_folders(self) -> Dict[str, Optional[Path]]:
        """Generate PDFs for all configured folders that have PNG files."""
        results = {}
        
        # Get all configured folder names
        folder_names = self.config.get_all_folder_names()
        
        # Filter out 'ALL' folder as it's used for global headers
        folder_names = [name for name in folder_names if name != 'ALL']
        
        self.logger.log_file_operation(f"Starting PDF generation for {len(folder_names)} folders")
        
        for folder_name in folder_names:
            self.logger.log_file_operation(f"Processing folder: {folder_name}")
            pdf_path = self.generate_pdf_for_folder(folder_name)
            results[folder_name] = pdf_path
        
        successful_pdfs = [path for path in results.values() if path is not None]
        self.logger.log_file_operation(
            f"PDF generation completed: {len(successful_pdfs)}/{len(folder_names)} successful"
        )
        
        return results
    
    def validate_pdf_size(self, pdf_path: Path) -> bool:
        """Validate PDF size against maximum attachment size."""
        try:
            pdf_size_mb = pdf_path.stat().st_size / 1024 / 1024
            max_size_mb = self.config.general_config.max_attachment_size_mb
            
            if pdf_size_mb > max_size_mb:
                self.logger.log_error(
                    f"PDF size ({pdf_size_mb:.2f} MB) exceeds maximum attachment size ({max_size_mb} MB): {pdf_path}"
                )
                return False
            
            return True
            
        except Exception as e:
            self.logger.log_error(f"Failed to validate PDF size: {pdf_path}", e)
            return False
    
    def get_folders_with_new_pngs(self) -> List[str]:
        """Get list of folders that have PNG files to process."""
        folders_with_pngs = []
        
        # Get all configured folder names
        folder_names = self.config.get_all_folder_names()
        
        # Filter out 'ALL' folder
        folder_names = [name for name in folder_names if name != 'ALL']
        
        for folder_name in folder_names:
            folder_path = self.input_dir / folder_name
            if folder_path.exists():
                png_files = get_png_files(folder_path, self.config.processing_config.png_file_extensions)
                if png_files:
                    folders_with_pngs.append(folder_name)
        
        return folders_with_pngs
    
    def cleanup_old_pdfs(self, keep_days: int = 7):
        """Clean up old PDF files from output directory."""
        try:
            cutoff_date = datetime.datetime.now() - datetime.timedelta(days=keep_days)
            
            pdf_files = list(self.output_dir.glob('*.pdf'))
            metadata_files = list(self.output_dir.glob('*.json'))
            
            cleaned_count = 0
            
            for file_path in pdf_files + metadata_files:
                file_date = datetime.datetime.fromtimestamp(file_path.stat().st_mtime)
                if file_date < cutoff_date:
                    file_path.unlink()
                    cleaned_count += 1
                    self.logger.log_file_operation(f"Cleaned up old file: {file_path}")
            
            if cleaned_count > 0:
                self.logger.log_file_operation(f"Cleaned up {cleaned_count} old files from output directory")
            
        except Exception as e:
            self.logger.log_error("Failed to cleanup old PDF files", e)


if __name__ == "__main__":
    # Test the PDF generator
    try:
        config_manager = ConfigManager()
        pdf_generator = PDFGenerator(config_manager)
        
        # Get folders with PNG files
        folders = pdf_generator.get_folders_with_new_pngs()
        print(f"Folders with PNG files: {folders}")
        
        # Generate PDFs for all folders
        if folders:
            results = pdf_generator.generate_pdfs_for_all_folders()
            print("PDF Generation Results:")
            for folder, pdf_path in results.items():
                if pdf_path:
                    print(f"  {folder}: {pdf_path}")
                else:
                    print(f"  {folder}: Failed")
        else:
            print("No folders with PNG files found")
            
    except Exception as e:
        print(f"Error testing PDF generator: {e}")
