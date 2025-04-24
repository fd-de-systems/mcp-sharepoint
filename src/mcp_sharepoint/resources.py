import base64
from io import BytesIO
from typing import Dict, Any, List, Optional
from .common import logger, SHP_DOC_LIBRARY, sp_context

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
        "created": f.properties.get("TimeCreated").isoformat() if f.properties.get("TimeCreated") else None,
        "modified": f.properties.get("TimeLastModified").isoformat() if f.properties.get("TimeLastModified") else None
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
        "created": f.properties.get("TimeCreated").isoformat() if f.properties.get("TimeCreated") else None,
        "modified": f.properties.get("TimeLastModified").isoformat() if f.properties.get("TimeLastModified") else None
    } for f in files]

def get_document_content(folder_name: str, file_name: str) -> Dict[str, Any]:
    """Get content of a specified document"""
    logger.info(f"Getting document content for {file_name} in folder {folder_name}")
    file_path = _get_sp_path(f"{folder_name}/{file_name}")
    
    # Use optimized method to get file with needed properties
    file = sp_context.web.get_file_by_server_relative_url(file_path)
    sp_context.load(file, ["Exists", "Length", "Name"])
    sp_context.execute_query()
    logger.info(f"File exists: {file.exists}, size: {file.length}")
    
    # Use binary file handler from library
    content_stream = BytesIO()
    file.download(content_stream)
    sp_context.execute_query()
    content_stream.seek(0)
    content = content_stream.read()
    
    # Process text vs binary files
    is_text_file = file_name.lower().endswith(('.txt', '.csv', '.json', '.xml', '.html', '.md', '.js', '.css', '.py'))
    content_dict = {"content": content.decode('utf-8')} if is_text_file else {"content_base64": base64.b64encode(content).decode('ascii')}
    return {
        "name": file_name,
        "content_type": "text" if is_text_file else "binary",
        **content_dict,
        "size": len(content)
    }