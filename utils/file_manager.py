import os
import re
import sys
from typing import Optional

class FileManager:
    @staticmethod
    def get_resource_path(relative_path: str) -> str:
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            # Fallback to current directory or script directory
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        return os.path.join(base_path, relative_path)

    @staticmethod
    def sanitize_filename(name: str) -> str:
        """
        Removes invalid characters from a string to make it safe for folder and file names.
        """
        # Remove characters that are generally invalid in Windows/Linux file paths
        name = re.sub(r'[\\/*?:"<>|]', "", name)
        return name.strip()

    @staticmethod
    def generate_simple_output_path(output_dir: str, output_filename: str) -> str:
        """
        Generates the output path and ensures the folder structure exists.
        
        Args:
            output_dir: Directory where the file will be saved.
            output_filename: Name of the output file.
            
        Returns:
            The complete absolute file path for saving the output document.
        """
        if not output_dir or not output_filename:
            raise ValueError("Output Folder and File Name are required.")
            
        # Create directories if they do not exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Ensure pdf extension
        if not output_filename.lower().endswith('.pdf'):
            output_filename += '.pdf'
            
        return os.path.join(output_dir, output_filename)
