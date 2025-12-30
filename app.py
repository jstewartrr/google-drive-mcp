"""
Google Drive MCP Server v3.2.0
==============================
MCP server for Google Drive file access via service account.
Added: create_folder, upload_file tools
"""

import os
import json
import io
import base64
from flask import Flask, request, jsonify
import logging
from datetime import datetime

# Google APIs
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# =============================================================================
# GOOGLE DRIVE CLIENT
# =============================================================================

def get_drive_service():
    """Get authenticated Drive service using service account."""
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not creds_json:
        raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON environment variable not set")
    
    creds_dict = json.loads(creds_json)
    credentials = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=['https://www.googleapis.com/auth/drive']
    )
    return build('drive', 'v3', credentials=credentials)

# =============================================================================
# TOOL IMPLEMENTATIONS
# =============================================================================

def list_shared_drives():
    """List all Shared Drives accessible to the service account."""
    try:
        service = get_drive_service()
        results = service.drives().list(pageSize=100).execute()
        drives = results.get('drives', [])
        return {
            "success": True,
            "drives": [{"id": d['id'], "name": d['name']} for d in drives]
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def list_folder_contents(folder_id: str, page_size: int = 50):
    """List files in a folder (works with Shared Drives)."""
    try:
        service = get_drive_service()
        query = f"'{folder_id}' in parents and trashed = false"
        results = service.files().list(
            q=query,
            pageSize=page_size,
            fields="files(id, name, mimeType, size, modifiedTime, webViewLink)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        files = results.get('files', [])
        return {
            "success": True,
            "folder_id": folder_id,
            "files": files
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def search_files(query: str, folder_id: str = None, file_type: str = None):
    """Search files by name (includes Shared Drives)."""
    try:
        service = get_drive_service()
        q_parts = [f"name contains '{query}'", "trashed = false"]
        if folder_id:
            q_parts.append(f"'{folder_id}' in parents")
        if file_type:
            mime_map = {
                "pdf": "application/pdf",
                "doc": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "sheet": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "xls": "application/vnd.ms-excel",
                "ppt": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                "folder": "application/vnd.google-apps.folder"
            }
            if file_type.lower() in mime_map:
                q_parts.append(f"mimeType = '{mime_map[file_type.lower()]}'")
        
        results = service.files().list(
            q=" and ".join(q_parts),
            pageSize=50,
            fields="files(id, name, mimeType, size, modifiedTime, webViewLink, parents)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        return {
            "success": True,
            "query": query,
            "files": results.get('files', [])
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def get_file_metadata(file_id: str):
    """Get file metadata."""
    try:
        service = get_drive_service()
        file = service.files().get(
            fileId=file_id,
            fields="id, name, mimeType, size, modifiedTime, createdTime, webViewLink, parents",
            supportsAllDrives=True
        ).execute()
        return {"success": True, "file": file}
    except Exception as e:
        return {"success": False, "error": str(e)}

def read_text_file(file_id: str):
    """Read text files."""
    try:
        service = get_drive_service()
        content = service.files().get_media(fileId=file_id).execute()
        return {
            "success": True,
            "file_id": file_id,
            "content": content.decode('utf-8')
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def read_excel_file(file_id: str, sheet_name: str = None):
    """Read Excel/Sheets files."""
    try:
        import openpyxl
        service = get_drive_service()
        content = service.files().get_media(fileId=file_id).execute()
        
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        sheets_data = {}
        
        target_sheets = [sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.sheetnames
        
        for sheet in target_sheets:
            ws = wb[sheet]
            data = []
            for row in ws.iter_rows(values_only=True):
                data.append([str(cell) if cell is not None else "" for cell in row])
            sheets_data[sheet] = data
        
        return {
            "success": True,
            "file_id": file_id,
            "sheets": sheets_data
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def read_pdf_file(file_id: str, page_numbers: list = None):
    """Extract PDF text."""
    try:
        import fitz  # PyMuPDF
        service = get_drive_service()
        content = service.files().get_media(fileId=file_id).execute()
        
        doc = fitz.open(stream=content, filetype="pdf")
        pages_text = {}
        
        target_pages = page_numbers if page_numbers else range(len(doc))
        for page_num in target_pages:
            if 0 <= page_num < len(doc):
                pages_text[page_num] = doc[page_num].get_text()
        
        return {
            "success": True,
            "file_id": file_id,
            "total_pages": len(doc),
            "pages": pages_text
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def read_word_file(file_id: str):
    """Extract Word text."""
    try:
        from docx import Document
        service = get_drive_service()
        content = service.files().get_media(fileId=file_id).execute()
        
        doc = Document(io.BytesIO(content))
        text = "\n".join([para.text for para in doc.paragraphs])
        
        return {
            "success": True,
            "file_id": file_id,
            "content": text
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def read_powerpoint_file(file_id: str):
    """Extract PowerPoint text."""
    try:
        from pptx import Presentation
        service = get_drive_service()
        content = service.files().get_media(fileId=file_id).execute()
        
        prs = Presentation(io.BytesIO(content))
        slides_text = {}
        
        for i, slide in enumerate(prs.slides):
            text_parts = []
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text_parts.append(shape.text)
            slides_text[i] = "\n".join(text_parts)
        
        return {
            "success": True,
            "file_id": file_id,
            "total_slides": len(prs.slides),
            "slides": slides_text
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def create_folder(name: str, parent_folder_id: str = None, shared_drive_id: str = None):
    """Create a new folder in Google Drive or Shared Drive."""
    try:
        service = get_drive_service()
        
        file_metadata = {
            'name': name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        
        if parent_folder_id:
            file_metadata['parents'] = [parent_folder_id]
        elif shared_drive_id:
            file_metadata['parents'] = [shared_drive_id]
        
        # Create folder with Shared Drive support
        folder = service.files().create(
            body=file_metadata,
            fields='id, name, webViewLink, parents',
            supportsAllDrives=True
        ).execute()
        
        return {
            "success": True,
            "folder": {
                "id": folder.get('id'),
                "name": folder.get('name'),
                "webViewLink": folder.get('webViewLink'),
                "parents": folder.get('parents')
            }
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def upload_file(name: str, content_base64: str, parent_folder_id: str, mime_type: str = None):
    """Upload a file to Google Drive."""
    try:
        service = get_drive_service()
        
        # Decode base64 content
        content = base64.b64decode(content_base64)
        
        # Auto-detect mime type if not provided
        if not mime_type:
            ext = name.lower().split('.')[-1] if '.' in name else ''
            mime_map = {
                'pdf': 'application/pdf',
                'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                'txt': 'text/plain',
                'csv': 'text/csv',
                'json': 'application/json',
                'png': 'image/png',
                'jpg': 'image/jpeg',
                'jpeg': 'image/jpeg'
            }
            mime_type = mime_map.get(ext, 'application/octet-stream')
        
        file_metadata = {
            'name': name,
            'parents': [parent_folder_id]
        }
        
        media = MediaIoBaseUpload(
            io.BytesIO(content),
            mimetype=mime_type,
            resumable=True
        )
        
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, name, webViewLink, size',
            supportsAllDrives=True
        ).execute()
        
        return {
            "success": True,
            "file": {
                "id": file.get('id'),
                "name": file.get('name'),
                "webViewLink": file.get('webViewLink'),
                "size": file.get('size')
            }
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def move_file(file_id: str, new_parent_id: str):
    """Move a file to a different folder."""
    try:
        service = get_drive_service()
        
        # Get current parents
        file = service.files().get(
            fileId=file_id,
            fields='parents',
            supportsAllDrives=True
        ).execute()
        
        previous_parents = ",".join(file.get('parents', []))
        
        # Move file
        file = service.files().update(
            fileId=file_id,
            addParents=new_parent_id,
            removeParents=previous_parents,
            fields='id, name, parents, webViewLink',
            supportsAllDrives=True
        ).execute()
        
        return {
            "success": True,
            "file": file
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

# =============================================================================
# TOOL DEFINITIONS
# =============================================================================

TOOLS = [
    {
        "name": "list_shared_drives",
        "description": "[DRIVE] List all Shared Drives accessible to the service account",
        "inputSchema": {"type": "object", "properties": {}, "required": []}
    },
    {
        "name": "list_folder_contents",
        "description": "[DRIVE] List files in a folder (works with Shared Drives)",
        "inputSchema": {
            "type": "object",
            "properties": {
                "folder_id": {"type": "string", "description": "Folder ID or Shared Drive ID"},
                "page_size": {"type": "integer", "default": 50}
            },
            "required": ["folder_id"]
        }
    },
    {
        "name": "search_files",
        "description": "[DRIVE] Search files by name (includes Shared Drives)",
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {"type": "string"},
                "folder_id": {"type": "string"},
                "file_type": {"type": "string"}
            },
            "required": ["query"]
        }
    },
    {
        "name": "get_file_metadata",
        "description": "[DRIVE] Get file metadata",
        "inputSchema": {
            "type": "object",
            "properties": {"file_id": {"type": "string"}},
            "required": ["file_id"]
        }
    },
    {
        "name": "read_text_file",
        "description": "[DRIVE] Read text files",
        "inputSchema": {
            "type": "object",
            "properties": {"file_id": {"type": "string"}},
            "required": ["file_id"]
        }
    },
    {
        "name": "read_excel_file",
        "description": "[DRIVE] Read Excel/Sheets",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_id": {"type": "string"},
                "sheet_name": {"type": "string"}
            },
            "required": ["file_id"]
        }
    },
    {
        "name": "read_pdf_file",
        "description": "[DRIVE] Extract PDF text",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_id": {"type": "string"},
                "page_numbers": {"type": "array", "items": {"type": "integer"}}
            },
            "required": ["file_id"]
        }
    },
    {
        "name": "read_word_file",
        "description": "[DRIVE] Extract Word text",
        "inputSchema": {
            "type": "object",
            "properties": {"file_id": {"type": "string"}},
            "required": ["file_id"]
        }
    },
    {
        "name": "read_powerpoint_file",
        "description": "[DRIVE] Extract PowerPoint text",
        "inputSchema": {
            "type": "object",
            "properties": {"file_id": {"type": "string"}},
            "required": ["file_id"]
        }
    },
    {
        "name": "create_folder",
        "description": "[DRIVE] Create a new folder in Google Drive or Shared Drive",
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "Name for the new folder"},
                "parent_folder_id": {"type": "string", "description": "Parent folder ID (optional)"},
                "shared_drive_id": {"type": "string", "description": "Shared Drive ID if creating at root of shared drive (optional)"}
            },
            "required": ["name"]
        }
    },
    {
        "name": "upload_file",
        "description": "[DRIVE] Upload a file to Google Drive (base64 encoded content)",
        "inputSchema": {
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "File name with extension"},
                "content_base64": {"type": "string", "description": "Base64 encoded file content"},
                "parent_folder_id": {"type": "string", "description": "Destination folder ID"},
                "mime_type": {"type": "string", "description": "MIME type (optional, auto-detected from extension)"}
            },
            "required": ["name", "content_base64", "parent_folder_id"]
        }
    },
    {
        "name": "move_file",
        "description": "[DRIVE] Move a file to a different folder",
        "inputSchema": {
            "type": "object",
            "properties": {
                "file_id": {"type": "string", "description": "File ID to move"},
                "new_parent_id": {"type": "string", "description": "Destination folder ID"}
            },
            "required": ["file_id", "new_parent_id"]
        }
    }
]

TOOL_HANDLERS = {
    "list_shared_drives": lambda args: list_shared_drives(),
    "list_folder_contents": lambda args: list_folder_contents(args["folder_id"], args.get("page_size", 50)),
    "search_files": lambda args: search_files(args["query"], args.get("folder_id"), args.get("file_type")),
    "get_file_metadata": lambda args: get_file_metadata(args["file_id"]),
    "read_text_file": lambda args: read_text_file(args["file_id"]),
    "read_excel_file": lambda args: read_excel_file(args["file_id"], args.get("sheet_name")),
    "read_pdf_file": lambda args: read_pdf_file(args["file_id"], args.get("page_numbers")),
    "read_word_file": lambda args: read_word_file(args["file_id"]),
    "read_powerpoint_file": lambda args: read_powerpoint_file(args["file_id"]),
    "create_folder": lambda args: create_folder(args["name"], args.get("parent_folder_id"), args.get("shared_drive_id")),
    "upload_file": lambda args: upload_file(args["name"], args["content_base64"], args["parent_folder_id"], args.get("mime_type")),
    "move_file": lambda args: move_file(args["file_id"], args["new_parent_id"])
}

# =============================================================================
# MCP PROTOCOL
# =============================================================================

def handle_initialize(params):
    return {
        "protocolVersion": "2024-11-05",
        "capabilities": {"tools": {"listChanged": True}},
        "serverInfo": {"name": "google-drive-mcp", "version": "3.2.0"}
    }

def handle_tools_list(params):
    return {"tools": TOOLS}

def handle_tools_call(params):
    tool_name = params.get("name", "")
    arguments = params.get("arguments", {})
    
    handler = TOOL_HANDLERS.get(tool_name)
    if not handler:
        return {"content": [{"type": "text", "text": f"Unknown tool: {tool_name}"}], "isError": True}
    
    try:
        result = handler(arguments)
        return {"content": [{"type": "text", "text": json.dumps(result, indent=2, default=str)}]}
    except Exception as e:
        logger.error(f"Tool error: {e}")
        return {"content": [{"type": "text", "text": f"Error: {str(e)}"}], "isError": True}

def process_mcp_message(data):
    method = data.get("method", "")
    params = data.get("params", {})
    request_id = data.get("id", 1)
    
    if method == "initialize":
        result = handle_initialize(params)
    elif method == "tools/list":
        result = handle_tools_list(params)
    elif method == "tools/call":
        result = handle_tools_call(params)
    elif method == "notifications/initialized":
        return {"jsonrpc": "2.0", "id": request_id, "result": {}}
    else:
        return {"jsonrpc": "2.0", "id": request_id, "error": {"code": -32601, "message": f"Method not found: {method}"}}
    
    return {"jsonrpc": "2.0", "id": request_id, "result": result}

# =============================================================================
# FLASK ROUTES
# =============================================================================

@app.route("/", methods=["GET"])
def health_check():
    return jsonify({
        "status": "healthy",
        "service": "google-drive-mcp",
        "version": "3.2.0",
        "tools": [t["name"] for t in TOOLS]
    })

@app.route("/mcp", methods=["POST"])
def mcp_handler():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "Parse error"}}), 400
        response = process_mcp_message(data)
        return jsonify(response)
    except Exception as e:
        logger.error(f"MCP handler error: {e}")
        return jsonify({"jsonrpc": "2.0", "id": 1, "error": {"code": -32603, "message": str(e)}}), 500

if __name__ == "__main__":
    logger.info("Google Drive MCP v3.2.0 starting...")
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
