import base64
import os
import fitz
import io
import logging
from io import BytesIO
from typing import Dict, Any, List, Optional
from datetime import datetime
from .common import logger, SHP_DOC_LIBRARY, sp_context

# Helper function to safely convert to ISO format
def _to_iso_optional(dt_obj: Optional[datetime]) -> Optional[str]:
    """Converts a datetime object to ISO format string, or returns None if the object is None."""
    if dt_obj is not None:
        return dt_obj.isoformat()
    return None

def _get_sp_path(sub_path: Optional[str] = None) -> str:
    """Create a properly formatted SharePoint path"""
    return f"{SHP_DOC_LIBRARY}/{sub_path or ''}".rstrip('/')

def list_folders(parent_folder: Optional[str] = None) -> List[Dict[str, Any]]:
    """List folders in the specified directory or root if not specified"""
    path = _get_sp_path(parent_folder)
    log_location = parent_folder or "root directory"
    logger.info(f"Listing folders in {log_location}")
    
    # Use the ClientObject.get_items() method which handles loading automatically
    parent = sp_context.web.get_folder_by_server_relative_url(path)
    folders = parent.folders
    sp_context.load(folders, ["ServerRelativeUrl", "Name", "TimeCreated", "TimeLastModified"])
    sp_context.execute_query()
    
    # Convert directly to the required format
    return [{
        "name": f.name,
        "url": f.properties.get("ServerRelativeUrl"),
        "created": _to_iso_optional(f.properties.get("TimeCreated")),
        "modified": _to_iso_optional(f.properties.get("TimeLastModified"))
    } for f in folders]

def list_documents(folder_name: str) -> List[Dict[str, Any]]:
    """List all documents in a specified folder"""
    logger.info(f"Listing documents in folder: {folder_name}")
    path = _get_sp_path(folder_name)
    
    # Load files with specific properties to reduce data transfer
    folder = sp_context.web.get_folder_by_server_relative_url(path)
    files = folder.files
    sp_context.load(files, ["ServerRelativeUrl", "Name", "Length", "TimeCreated", "TimeLastModified"])
    sp_context.execute_query()
    
    # Convert directly to the required format
    return [{
        "name": f.name,
        "url": f.properties.get("ServerRelativeUrl"),
        "size": f.properties.get("Length"),
        "created": _to_iso_optional(f.properties.get("TimeCreated")),
        "modified": _to_iso_optional(f.properties.get("TimeLastModified"))
    } for f in files]

def extract_text_from_pdf(pdf_content):
    """Extract text from PDF using PyMuPDF with fallback methods."""
    logger = logging.getLogger(__name__)
    
    # Process the PDF
    try:
        pdf_document = fitz.open(stream=pdf_content, filetype="pdf")
        page_count = len(pdf_document)
        text_content = ""
        
        for page_num in range(page_count):
            page = pdf_document[page_num]
            text_content += page.get_text() + "\n"
        
        pdf_document.close()
        return text_content.strip(), page_count
    except Exception as e:
        logger.error(f"Error extracting text from PDF: {e}")
        raise

def get_folder_tree(parent_folder: Optional[str] = None) -> Dict[str, Any]:
    """Recursively lists all folders and files from a parent folder."""
    # Determine the current path and log location
    path = _get_sp_path(parent_folder)
    log_location = parent_folder or "root directory"
    logger.info(f"Building tree for {log_location} at path: {path}")

    # Get current folder object to have a starting point
    try:
        current_folder_obj = sp_context.web.get_folder_by_server_relative_url(path)
        sp_context.load(current_folder_obj, ["Name", "ServerRelativeUrl", "TimeCreated", "TimeLastModified"])
        sp_context.execute_query()
    except Exception as e:
        logger.error(f"Could not retrieve folder '{path}': {e}")
        # Return an error structure if the folder is not found
        return {
            "name": os.path.basename(path),
            "path": path,
            "type": "folder",
            "error": f"Could not access folder: {e}",
            "children": []
        }

    # Initialize the tree structure for the current folder
    tree = {
        "name": current_folder_obj.name,
        "path": current_folder_obj.properties.get("ServerRelativeUrl"),
        "type": "folder",
        "created": _to_iso_optional(current_folder_obj.properties.get("TimeCreated")),
        "modified": _to_iso_optional(current_folder_obj.properties.get("TimeLastModified")),
        "children": []
    }

    # Recursively get subfolders
    try:
        subfolders = list_folders(parent_folder)
        for folder_data in subfolders:
            # The path for recursion must be relative to the SHP_DOC_LIBRARY root.
            # The folder_data["url"] is the full server-relative URL, so we need to strip the site prefix and the library path.
            # Ensure web properties are loaded to get the title
            if not sp_context.web.is_property_available('Title'):
                sp_context.load(sp_context.web, ['Title'])
                sp_context.execute_query()

            site_path_segment = sp_context.web.properties['Title']
            full_shp_doc_library_path = f"/sites/{site_path_segment}/{SHP_DOC_LIBRARY}"
            if folder_data["url"].startswith(full_shp_doc_library_path):
                relative_path = folder_data["url"][len(full_shp_doc_library_path):].lstrip('/')
                tree["children"].append(get_folder_tree(relative_path))
            else:
                logger.warning(f"Could not determine relative path for folder: {folder_data['url']}")
    except Exception as e:
        logger.error(f"Could not list subfolders in '{path}': {e}")

    # Get files in the current folder
    try:
        # list_documents expects a relative path from the document library root
        files = list_documents(parent_folder)
        for file_data in files:
            tree["children"].append({
                "name": file_data["name"],
                "path": file_data["url"],
                "type": "file",
                "size": file_data.get("size"),
                "created": file_data.get("created"),
                "modified": file_data.get("modified"),
            })
    except Exception as e:
        logger.error(f"Could not list files in '{path}': {e}")

    return tree

def get_document_content(folder_name: str, file_name: str) -> dict:
    """Get content of a specified document, supports PDFs with text."""

    logger = logging.getLogger(__name__)
    
    file_path = _get_sp_path(f"{folder_name}/{file_name}")
    file = sp_context.web.get_file_by_server_relative_url(file_path)
    sp_context.load(file, ["Exists", "Length", "Name"])
    sp_context.execute_query()
    logger.info(f"File exists: {file.exists}, size: {file.length}")
    
    content_stream = io.BytesIO()
    file.download(content_stream)
    sp_context.execute_query()
    content_stream.seek(0)
    content = content_stream.read()

    is_pdf = file_name.lower().endswith('.pdf')
    is_text_file = file_name.lower().endswith(('.txt', '.csv', '.json', '.xml', '.html', '.md', '.js', '.css', '.py'))

    if is_pdf:
        try:
            # Use our robust PDF text extraction function
            text_content, page_count = extract_text_from_pdf(content)
            
            # Return text content with metadata
            return {
                "name": file_name,
                "content_type": "text",
                "content": text_content,
                "original_type": "pdf",
                "page_count": page_count,
                "size": len(content)
            }

        except ImportError as e:
            logger.warning(f"Missing dependencies for PDF processing: {e}. Returning binary content.")
            return {
                "name": file_name,
                "content_type": "binary",
                "content_base64": base64.b64encode(content).decode('ascii'),
                "original_type": "pdf",
                "error": f"PDF text extraction failed: {str(e)}",
                "size": len(content)
            }
        except Exception as e:
            logger.error(f"Error processing PDF: {e}")
            return {
                "name": file_name,
                "content_type": "binary",
                "content_base64": base64.b64encode(content).decode('ascii'),
                "original_type": "pdf",
                "error": f"PDF processing error: {str(e)}",
                "size": len(content)
            }

    elif is_text_file:
        try:
            return {
                "name": file_name,
                "content_type": "text",
                "content": content.decode('utf-8'),
                "size": len(content)
            }
        except UnicodeDecodeError:
            return {
                "name": file_name,
                "content_type": "binary",
                "content_base64": base64.b64encode(content).decode('ascii'),
                "size": len(content)
            }
    else:
        return {
            "name": file_name,
            "content_type": "binary",
            "content_base64": base64.b64encode(content).decode('ascii'),
            "size": len(content)
        }
