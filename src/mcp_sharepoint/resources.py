import base64, os, fitz, io, logging
from typing import Dict, Any, List, Optional
from .common import logger, SHP_DOC_LIBRARY, sp_context

logger = logging.getLogger(__name__)

# Configuration
FILE_TYPES = {
    'text': ['.txt', '.csv', '.json', '.xml', '.html', '.md', '.js', '.css', '.py'],
    'pdf': ['.pdf']
}

def _get_sp_path(sub_path: Optional[str] = None) -> str:
    """Create a properly formatted SharePoint path"""
    return f"{SHP_DOC_LIBRARY}/{sub_path or ''}".rstrip('/')

def _load_sp_items(path: str, item_type: str) -> List[Dict[str, Any]]:
    """Generic function to load folders or files from SharePoint"""
    folder = sp_context.web.get_folder_by_server_relative_url(path)
    items = getattr(folder, item_type)
    props = ["ServerRelativeUrl", "Name", "TimeCreated", "TimeLastModified"] + (["Length"] if item_type == "files" else [])
    sp_context.load(items, props)
    sp_context.execute_query()
    
    return [{
        "name": item.name,
        "url": item.properties.get("ServerRelativeUrl"),
        **({"size": item.properties.get("Length")} if item_type == "files" else {}),
        "created": item.properties.get("TimeCreated").isoformat() if item.properties.get("TimeCreated") else None,
        "modified": item.properties.get("TimeLastModified").isoformat() if item.properties.get("TimeLastModified") else None
    } for item in items]

def list_folders(parent_folder: Optional[str] = None) -> List[Dict[str, Any]]:
    """List folders in the specified directory or root if not specified"""
    logger.info(f"Listing folders in {parent_folder or 'root directory'}")
    return _load_sp_items(_get_sp_path(parent_folder), "folders")

def list_documents(folder_name: str) -> List[Dict[str, Any]]:
    """List all documents in a specified folder"""
    logger.info(f"Listing documents in folder: {folder_name}")
    return _load_sp_items(_get_sp_path(folder_name), "files")

def extract_text_from_pdf(pdf_content):
    """Extract text from PDF using PyMuPDF"""
    try:
        pdf_document = fitz.open(stream=pdf_content, filetype="pdf")
        text_content = "".join(pdf_document[i].get_text() + "\n" for i in range(len(pdf_document)))
        page_count = len(pdf_document)
        pdf_document.close()
        return text_content.strip(), page_count
    except Exception as e:
        logger.error(f"Error extracting text from PDF: {e}")
        raise

def get_folder_tree(parent_folder: Optional[str] = None) -> Dict[str, Any]:
    """Recursively list all folders and files from a parent folder"""
    path = _get_sp_path(parent_folder)
    logger.info(f"Building tree for {parent_folder or 'root'} at path: {path}")
    
    try:
        folder = sp_context.web.get_folder_by_server_relative_url(path)
        sp_context.load(folder, ["Name", "ServerRelativeUrl", "TimeCreated", "TimeLastModified"])
        sp_context.execute_query()
        
        # Get site title for path processing
        if not sp_context.web.is_property_available('Title'):
            sp_context.load(sp_context.web, ['Title'])
            sp_context.execute_query()
        lib_prefix = f"/sites/{sp_context.web.properties['Title']}/{SHP_DOC_LIBRARY}"
        
        # Build children recursively
        children = []
        for f in list_folders(parent_folder):
            if f["url"].startswith(lib_prefix):
                children.append(get_folder_tree(f["url"][len(lib_prefix):].lstrip('/')))
            else:
                logger.warning(f"Cannot determine relative path for {f['url']}")
        
        children.extend({
            "name": f["name"], "path": f["url"], "type": "file",
            **{k: v for k, v in f.items() if k not in ["name", "url"]}
        } for f in list_documents(parent_folder or ""))
        
        return {
            "name": folder.name,
            "path": folder.properties.get("ServerRelativeUrl"),
            "type": "folder",
            "created": folder.properties.get("TimeCreated").isoformat() if folder.properties.get("TimeCreated") else None,
            "modified": folder.properties.get("TimeLastModified").isoformat() if folder.properties.get("TimeLastModified") else None,
            "children": children
        }
    except Exception as e:
        logger.error(f"Failed to build tree for '{path}': {e}")
        return {"name": os.path.basename(path), "path": path, "type": "folder", "error": "Could not access folder", "children": []}

def get_document_content(folder_name: str, file_name: str) -> dict:
    """Retrieve document content; supports PDF text extraction"""
    file_path = _get_sp_path(f"{folder_name}/{file_name}")
    file = sp_context.web.get_file_by_server_relative_url(file_path)
    sp_context.load(file, ["Exists", "Length", "Name"])
    sp_context.execute_query()
    logger.info(f"File exists: {file.exists}, size: {file.length}")

    content = io.BytesIO()
    file.download(content)
    sp_context.execute_query()
    content_bytes = content.getvalue()
    
    # Determine file type and process accordingly
    lower_name = file_name.lower()
    file_type = next((t for t, exts in FILE_TYPES.items() if any(lower_name.endswith(ext) for ext in exts)), 'binary')
    
    if file_type == 'pdf':
        try:
            text, pages = extract_text_from_pdf(content_bytes)
            return {"name": file_name, "content_type": "text", "content": text, "original_type": "pdf", "page_count": pages, "size": len(content_bytes)}
        except Exception as e:
            logger.warning(f"PDF processing failed: {e}")
            return {"name": file_name, "content_type": "binary", "content_base64": base64.b64encode(content_bytes).decode(), "original_type": "pdf", "size": len(content_bytes)}
    
    if file_type == 'text':
        try:
            return {"name": file_name, "content_type": "text", "content": content_bytes.decode('utf-8'), "size": len(content_bytes)}
        except UnicodeDecodeError:
            pass
    
    return {"name": file_name, "content_type": "binary", "content_base64": base64.b64encode(content_bytes).decode(), "size": len(content_bytes)}
