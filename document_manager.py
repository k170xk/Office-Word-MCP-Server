"""
Document manager that handles storage operations transparently.
Downloads documents before editing, uploads after saving.
"""

import os
import hashlib
import tempfile
from typing import Optional

from storage_adapter import get_storage_adapter


def _temp_file_for_storage_key(temp_dir: str, storage_key: str) -> str:
    digest = hashlib.sha256(storage_key.encode("utf-8")).hexdigest()[:24]
    base = storage_key.replace("\\", "/").split("/")[-1]
    return os.path.join(temp_dir, f"edit_{digest}_{base}")


class DocumentManager:
    """Manages document lifecycle with automatic storage sync."""
    
    def __init__(self):
        self.storage = get_storage_adapter()
        self.temp_dir = tempfile.mkdtemp(prefix='doc_edit_')
    
    def get_local_path(self, storage_key: str, create_if_missing: bool = False) -> str:
        """
        Resolve a logical storage document key (e.g. ``workspace/note.docx``) to an
        editable temp file path, downloading when the object exists remotely.
        
        Args:
            storage_key: Normalized POSIX path under DISK_PATH (may include workspaces)
            create_if_missing: If True and key missing in storage, return new temp path
        """
        local_path = _temp_file_for_storage_key(self.temp_dir, storage_key)
        dirname = os.path.dirname(local_path)
        if dirname:
            os.makedirs(dirname, exist_ok=True)

        if self.storage.document_exists(storage_key):
            self.storage.download_document(storage_key, local_path)
            return local_path

        if create_if_missing:
            return local_path

        raise FileNotFoundError(f"Document {storage_key} not found")
    
    def save_document(self, local_path: str, filename: str) -> str:
        """
        Save document to storage and return URL.
        
        Args:
            local_path: Local file path
            filename: Target filename in storage
        
        Returns:
            Document URL
        """
        # Upload to storage
        url = self.storage.upload_document(local_path, filename)
        return url
    
    def get_document_url(self, filename: str) -> str:
        """Get the public URL for a document."""
        return self.storage.get_document_url(filename)
    
    def cleanup_temp_path(self, local_path: Optional[str] = None) -> None:
        """Remove a single temp editing file produced by ``get_local_path``."""
        if local_path and os.path.isfile(local_path) and os.path.abspath(local_path).startswith(
            os.path.abspath(self.temp_dir) + os.sep
        ):
            try:
                os.remove(local_path)
            except OSError:
                pass

    def cleanup_temp(self, filename: Optional[str] = None):
        """Clean up temporary files."""
        if filename:
            temp_path = os.path.join(self.temp_dir, filename)
            if os.path.exists(temp_path):
                os.remove(temp_path)
        else:
            # Clean up entire temp directory
            import shutil
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)


# Global document manager instance
_document_manager: Optional[DocumentManager] = None


def get_document_manager() -> DocumentManager:
    """Get or create the global document manager instance."""
    global _document_manager
    if _document_manager is None:
        _document_manager = DocumentManager()
    return _document_manager


def with_storage_sync(tool_func):
    """
    Decorator that automatically handles storage sync for document operations.
    Downloads document before operation, uploads after.
    """
    async def wrapper(*args, **kwargs):
        # Find filename parameter
        filename = kwargs.get('filename') or kwargs.get('source_filename')
        
        if not filename:
            # No filename, just call the tool normally
            return await tool_func(*args, **kwargs)
        
        manager = get_document_manager()
        local_path: Optional[str] = None
        
        try:
            # Get local path (downloads if exists, creates if new)
            create_if_missing = 'create' in tool_func.__name__.lower() or 'add' in tool_func.__name__.lower()
            local_path = manager.get_local_path(filename, create_if_missing=create_if_missing)
            
            # Update kwargs with local path
            if 'filename' in kwargs:
                kwargs['filename'] = local_path
            if 'source_filename' in kwargs:
                kwargs['source_filename'] = local_path
            
            # Call the tool function
            result = await tool_func(*args, **kwargs)
            
            # Upload back to storage (if document was modified)
            if os.path.exists(local_path):
                doc_url = manager.save_document(local_path, filename)
                # Enhance result with URL
                if isinstance(result, str):
                    result = f"{result}\n\nDocument URL: {doc_url}\nDownload URL: {doc_url}"
            
            return result
        
        except FileNotFoundError as e:
            return f"Error: {str(e)}"
        except Exception as e:
            return f"Error: {str(e)}"
        finally:
            manager.cleanup_temp_path(local_path)

    return wrapper

