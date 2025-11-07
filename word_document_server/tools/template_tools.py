"""
Template management tools for Word Document Server.
"""

import os
import shutil
from typing import Optional
from docx import Document
from word_document_server.utils.file_utils import ensure_docx_extension


TEMPLATE_FILENAME = '.template.docx'
# Use the same storage directory as documents
TEMPLATE_DIR = os.getenv('DISK_PATH', os.getenv('DOCUMENTS_DIR', '/mnt/disk/documents'))


def get_template_path() -> str:
    """Get the path to the template file."""
    return os.path.join(TEMPLATE_DIR, TEMPLATE_FILENAME)


def template_exists() -> bool:
    """Check if a template exists."""
    return os.path.exists(get_template_path())


async def set_template_from_file(template_filename: str) -> str:
    """
    Set the template from an existing document file.
    
    Args:
        template_filename: Path to the template document file
    """
    template_filename = ensure_docx_extension(template_filename)
    
    if not os.path.exists(template_filename):
        return f"Template file {template_filename} does not exist"
    
    try:
        template_path = get_template_path()
        # Ensure template directory exists
        os.makedirs(os.path.dirname(template_path), exist_ok=True)
        
        # Copy the file to template location
        shutil.copy2(template_filename, template_path)
        
        return f"Template set successfully from {template_filename}"
    except Exception as e:
        return f"Failed to set template: {str(e)}"


async def get_template_info() -> str:
    """Get information about the current template."""
    if not template_exists():
        return "No template is set. Use set_template_from_file to upload a template."
    
    try:
        template_path = get_template_path()
        size = os.path.getsize(template_path) / 1024  # KB
        return f"Template exists: {template_path} ({size:.2f} KB)"
    except Exception as e:
        return f"Error getting template info: {str(e)}"


async def clear_template() -> str:
    """Remove the current template."""
    template_path = get_template_path()
    
    if not os.path.exists(template_path):
        return "No template to clear"
    
    try:
        os.remove(template_path)
        return "Template cleared successfully"
    except Exception as e:
        return f"Failed to clear template: {str(e)}"

