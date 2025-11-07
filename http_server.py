#!/usr/bin/env python3
"""
HTTP server wrapper for Office Word MCP Server.
Provides OpenAI-compatible JSON-RPC endpoints and document serving.
Force deploy: 2025-11-07
"""

import os
import json
import asyncio
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse
import sys
import inspect

# Add the project root to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import all tool functions directly
from word_document_server.tools import (
    document_tools,
    content_tools,
    format_tools,
    protection_tools,
    footnote_tools,
    extended_document_tools,
    comment_tools
)
from word_document_server.tools import template_tools, document_formatting_tools
from document_manager import get_document_manager
from storage_adapter import get_storage_adapter

# Document storage directory
DOCUMENTS_DIR = os.getenv('DOCUMENTS_DIR', './documents')
BASE_URL = os.getenv('BASE_URL', '')  # Will be set from Render service URL

# Ensure documents directory exists
os.makedirs(DOCUMENTS_DIR, exist_ok=True)

# Build tool registry by inspecting all tool modules
TOOL_REGISTRY = {}

def build_tool_registry():
    """Build a registry of all available tools."""
    global TOOL_REGISTRY
    
    # Map of tool names to their functions
    tools_map = {
        'create_document': document_tools.create_document,
        'copy_document': document_tools.copy_document,
        'get_document_info': document_tools.get_document_info,
        'get_document_text': document_tools.get_document_text,
        'get_document_outline': document_tools.get_document_outline,
        'list_available_documents': document_tools.list_available_documents,
        'get_document_xml': document_tools.get_document_xml_tool,
        'insert_header_near_text': content_tools.insert_header_near_text_tool,
        'insert_line_or_paragraph_near_text': content_tools.insert_line_or_paragraph_near_text_tool,
        'insert_numbered_list_near_text': content_tools.insert_numbered_list_near_text_tool,
        'add_paragraph': content_tools.add_paragraph,
        'add_heading': content_tools.add_heading,
        'add_picture': content_tools.add_picture,
        'add_table': content_tools.add_table,
        'add_page_break': content_tools.add_page_break,
        'delete_paragraph': content_tools.delete_paragraph,
        'search_and_replace': content_tools.search_and_replace,
        'create_custom_style': format_tools.create_custom_style,
        'format_text': format_tools.format_text,
        'format_table': format_tools.format_table,
        'set_table_cell_shading': format_tools.set_table_cell_shading,
        'apply_table_alternating_rows': format_tools.apply_table_alternating_rows,
        'highlight_table_header': format_tools.highlight_table_header,
        'merge_table_cells': format_tools.merge_table_cells,
        'merge_table_cells_horizontal': format_tools.merge_table_cells_horizontal,
        'merge_table_cells_vertical': format_tools.merge_table_cells_vertical,
        'set_table_cell_alignment': format_tools.set_table_cell_alignment,
        'set_table_alignment_all': format_tools.set_table_alignment_all,
        'protect_document': protection_tools.protect_document,
        'unprotect_document': protection_tools.unprotect_document,
        'add_footnote_to_document': footnote_tools.add_footnote_to_document,
        'add_footnote_after_text': footnote_tools.add_footnote_after_text,
        'add_footnote_before_text': footnote_tools.add_footnote_before_text,
        'add_footnote_enhanced': footnote_tools.add_footnote_enhanced,
        'add_endnote_to_document': footnote_tools.add_endnote_to_document,
        'customize_footnote_style': footnote_tools.customize_footnote_style,
        'delete_footnote_from_document': footnote_tools.delete_footnote_from_document,
        'add_footnote_robust': footnote_tools.add_footnote_robust_tool,
        'validate_document_footnotes': footnote_tools.validate_footnotes_tool,
        'delete_footnote_robust': footnote_tools.delete_footnote_robust_tool,
        'get_paragraph_text_from_document': extended_document_tools.get_paragraph_text_from_document,
        'find_text_in_document': extended_document_tools.find_text_in_document,
        'convert_to_pdf': extended_document_tools.convert_to_pdf,
        'get_all_comments': comment_tools.get_all_comments,
        'get_comments_by_author': comment_tools.get_comments_by_author,
        'get_comments_for_paragraph': comment_tools.get_comments_for_paragraph,
        'set_table_column_width': format_tools.set_table_column_width,
        'set_table_column_widths': format_tools.set_table_column_widths,
        'set_table_width': format_tools.set_table_width,
        'auto_fit_table_columns': format_tools.auto_fit_table_columns,
        'format_table_cell_text': format_tools.format_table_cell_text,
        'set_table_cell_padding': format_tools.set_table_cell_padding,
        'replace_paragraph_block_below_header': content_tools.replace_paragraph_block_below_header_tool,
        'replace_block_between_manual_anchors': content_tools.replace_block_between_manual_anchors_tool,
        'set_template_from_file': template_tools.set_template_from_file,
        'get_template_info': template_tools.get_template_info,
        'clear_template': template_tools.clear_template,
        'set_default_font': document_formatting_tools.set_default_font,
        'update_header_title_subtitle': document_formatting_tools.update_header_title_subtitle,
        'get_header_info': document_formatting_tools.get_header_info,
    }
    
    TOOL_REGISTRY = tools_map

# Build registry on import
build_tool_registry()


class MCPHTTPHandler(BaseHTTPRequestHandler):
    """HTTP handler for MCP JSON-RPC requests and document serving."""
    
    def do_OPTIONS(self):
        """Handle CORS preflight requests."""
        self.send_response(200)
        self.send_cors_headers()
        self.end_headers()
    
    def send_cors_headers(self):
        """Send CORS headers."""
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Content-Type', 'application/json')
    
    def do_GET(self):
        """Handle GET requests for tool discovery and document serving."""
        parsed_path = urlparse(self.path)
        path = parsed_path.path
        
        # Tool discovery endpoint
        if path == '/mcp/stream' or path == '/mcp/tools':
            try:
                request = {
                    "jsonrpc": "2.0",
                    "id": 1,
                    "method": "tools/list",
                    "params": {}
                }
                response = asyncio.run(self.handle_mcp_request(request))
                
                self.send_response(200)
                self.send_cors_headers()
                self.end_headers()
                self.wfile.write(json.dumps(response).encode('utf-8'))
            except Exception as e:
                self.send_error(500, f"Error: {str(e)}")
        
        # Document serving endpoint
        elif path.startswith('/documents/'):
            filename = path.replace('/documents/', '')
            self.serve_document(filename)
        
        # Health check
        elif path == '/health':
            self.send_response(200)
            self.send_cors_headers()
            self.end_headers()
            self.wfile.write(json.dumps({"status": "ok"}).encode('utf-8'))
        
        # Template info endpoint
        elif path == '/template/info':
            try:
                info = asyncio.run(template_tools.get_template_info())
                self.send_response(200)
                self.send_cors_headers()
                self.end_headers()
                self.wfile.write(json.dumps({"info": info}).encode('utf-8'))
            except Exception as e:
                self.send_error(500, f"Error: {str(e)}")
        
        else:
            self.send_error(404, "Not found")
    
    def do_POST(self):
        """Handle POST requests for MCP tool calls and template uploads."""
        parsed_path = urlparse(self.path)
        path = parsed_path.path
        
        # Template upload endpoint
        if path == '/upload-template':
            self.handle_template_upload()
            return
        
        if path == '/mcp/stream':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length).decode('utf-8')
            
            try:
                request = json.loads(body)
                response = asyncio.run(self.handle_mcp_request(request))
                
                self.send_response(200)
                self.send_cors_headers()
                self.end_headers()
                self.wfile.write(json.dumps(response).encode('utf-8'))
            except json.JSONDecodeError as e:
                self.send_error(400, f"Invalid JSON: {str(e)}")
            except Exception as e:
                self.send_error(500, f"Error: {str(e)}")
        else:
            self.send_error(404, "Not found")
    
    async def handle_mcp_request(self, request: dict):
        """Handle an MCP JSON-RPC request."""
        method = request.get('method')
        params = request.get('params', {})
        request_id = request.get('id')
        
        try:
            if method == 'initialize':
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "protocolVersion": "2024-11-05",
                        "capabilities": {
                            "tools": {}
                        },
                        "serverInfo": {
                            "name": "office-word-mcp-server",
                            "version": "1.0.0"
                        }
                    }
                }
            
            elif method == 'tools/list':
                tools = []
                for tool_name, tool_func in TOOL_REGISTRY.items():
                    tools.append({
                        "name": tool_name,
                        "description": tool_func.__doc__ or f"Tool: {tool_name}",
                        "inputSchema": self._get_tool_schema(tool_func)
                    })
                
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "tools": tools
                    }
                }
            
            elif method == 'tools/call':
                tool_name = params.get('name')
                arguments = params.get('arguments', {})
                
                if tool_name not in TOOL_REGISTRY:
                    return {
                        "jsonrpc": "2.0",
                        "id": request_id,
                        "error": {
                            "code": -32601,
                            "message": f"Tool not found: {tool_name}"
                        }
                    }
                
                tool_func = TOOL_REGISTRY[tool_name]
                
                # Use storage adapter for document operations
                manager = get_document_manager()
                storage = get_storage_adapter()
                
                # Handle filename parameters - download from storage if exists
                original_filename = None
                local_path = None
                
                if 'filename' in arguments:
                    original_filename = arguments['filename']
                    # Extract just the filename (remove path if present)
                    filename_base = os.path.basename(original_filename)
                    
                    # Check if document exists in storage
                    create_if_missing = 'create' in tool_name or 'add' in tool_name
                    try:
                        local_path = manager.get_local_path(filename_base, create_if_missing=create_if_missing)
                        arguments['filename'] = local_path
                    except FileNotFoundError:
                        if create_if_missing:
                            local_path = manager.get_local_path(filename_base, create_if_missing=True)
                            arguments['filename'] = local_path
                        else:
                            return {
                                "jsonrpc": "2.0",
                                "id": request_id,
                                "error": {
                                    "code": -32602,
                                    "message": f"Document {filename_base} not found"
                                }
                            }
                
                if 'source_filename' in arguments:
                    source_filename_base = os.path.basename(arguments['source_filename'])
                    try:
                        source_local_path = manager.get_local_path(source_filename_base, create_if_missing=False)
                        arguments['source_filename'] = source_local_path
                    except FileNotFoundError:
                        return {
                            "jsonrpc": "2.0",
                            "id": request_id,
                            "error": {
                                "code": -32602,
                                "message": f"Source document {source_filename_base} not found"
                            }
                        }
                
                try:
                    # Call the tool
                    if asyncio.iscoroutinefunction(tool_func):
                        result = await tool_func(**arguments)
                    else:
                        result = tool_func(**arguments)
                    
                    # Upload document back to storage if it was modified
                    if local_path and os.path.exists(local_path):
                        # Get the original filename (before we changed it to local_path)
                        if original_filename:
                            filename_base = os.path.basename(original_filename)
                        else:
                            # Extract from arguments
                            filename_base = os.path.basename(arguments.get('filename', ''))
                        
                        # Ensure .docx extension
                        if filename_base and not filename_base.endswith('.docx'):
                            filename_base = f"{filename_base}.docx"
                        
                        if filename_base:
                            # Save to storage
                            doc_url = manager.save_document(local_path, filename_base)
                            # Enhance result with URL
                            if isinstance(result, str):
                                from urllib.parse import quote
                                # URL encode the filename for the download URL
                                encoded_filename = quote(filename_base)
                                download_url = f"{BASE_URL or 'https://office-word-mcp.onrender.com'}/documents/{encoded_filename}"
                                result = f"{result}\n\nDocument saved: {filename_base}\nDownload URL: {download_url}"
                    
                    enhanced_result = str(result)
                finally:
                    # Cleanup temp files
                    if local_path and os.path.exists(local_path):
                        manager.cleanup_temp(os.path.basename(local_path))
                
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {
                        "content": [
                            {
                                "type": "text",
                                "text": enhanced_result
                            }
                        ]
                    }
                }
            
            else:
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32601,
                        "message": f"Method not found: {method}"
                    }
                }
        
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "error": {
                    "code": -32603,
                    "message": str(e)
                }
            }
    
    def _get_tool_schema(self, tool_func):
        """Extract JSON schema from tool function signature."""
        import typing
        sig = inspect.signature(tool_func)
        
        properties = {}
        required = []
        
        # Get docstring for better descriptions
        docstring = tool_func.__doc__ or ""
        
        for param_name, param in sig.parameters.items():
            if param_name == 'self':
                continue
            
            param_type = param.annotation
            param_default = param.default
            
            # Handle Optional types
            if hasattr(typing, 'get_origin') and typing.get_origin(param_type) is typing.Union:
                args = typing.get_args(param_type)
                # If Union includes None, it's Optional
                if type(None) in args:
                    # Get the actual type (not None)
                    param_type = next((arg for arg in args if arg is not type(None)), str)
            
            # Map Python types to JSON schema types
            prop_schema = {}
            
            if param_type == str or param_type == inspect.Parameter.empty or param_type == type(None):
                prop_schema["type"] = "string"
            elif param_type == int:
                prop_schema["type"] = "integer"
            elif param_type == float:
                prop_schema["type"] = "number"
            elif param_type == bool:
                prop_schema["type"] = "boolean"
            elif param_type == list or (hasattr(typing, '_GenericAlias') and 'list' in str(param_type)):
                prop_schema["type"] = "array"
                prop_schema["items"] = {"type": "string"}  # Default to string array
            elif param_type == dict:
                prop_schema["type"] = "object"
            else:
                prop_schema["type"] = "string"
            
            # Add enum constraints for known parameters
            if param_name == 'position':
                prop_schema["enum"] = ["before", "after"]
                prop_schema["description"] = "Position relative to target: 'before' or 'after'"
            elif param_name == 'bullet_type':
                prop_schema["enum"] = ["bullet", "number"]
                prop_schema["description"] = "List type: 'bullet' for bullets (•) or 'number' for numbered (1,2,3)"
            elif param_name == 'list_items':
                prop_schema["description"] = "Array of strings, each as a list item"
            else:
                # Try to extract description from docstring
                desc = f"Parameter: {param_name}"
                if docstring:
                    # Look for param_name in docstring
                    import re
                    pattern = rf"{param_name}:\s*([^\n]+)"
                    match = re.search(pattern, docstring)
                    if match:
                        desc = match.group(1).strip()
                prop_schema["description"] = desc
            
            properties[param_name] = prop_schema
            
            # Only require if no default value
            if param_default == inspect.Parameter.empty:
                required.append(param_name)
        
        schema = {
            "type": "object",
            "properties": properties
        }
        
        if required:
            schema["required"] = required
        
        return schema
    
    def _enhance_result_with_url(self, result: str, arguments: dict):
        """Enhance tool result with document URL if applicable."""
        filename = arguments.get('filename') or arguments.get('source_filename')
        
        if filename:
            # Extract just the filename if it's a path
            doc_filename = os.path.basename(filename)
            if not doc_filename.endswith('.docx'):
                doc_filename = f"{doc_filename}.docx"
            
            base_url = BASE_URL or f"http://{self.server.server_address[0]}:{self.server.server_address[1]}"
            doc_url = f"{base_url}/documents/{doc_filename}"
            
            if "successfully" in result.lower() or "created" in result.lower() or "saved" in result.lower():
                return f"{result}\n\nDocument URL: {doc_url}\nDownload URL: {doc_url}"
        
        return result
    
    def serve_document(self, filename: str):
        """Serve a document file from storage."""
        from urllib.parse import unquote
        
        # URL decode the filename (handle %20 for spaces, etc.)
        filename = unquote(filename)
        # Security: prevent directory traversal
        filename = os.path.basename(filename)
        
        try:
            storage = get_storage_adapter()
            manager = get_document_manager()
            
            # Check if document exists in storage first
            if not storage.document_exists(filename):
                self.send_error(404, f"Document '{filename}' not found")
                return
            
            # Download from storage to temp location
            local_path = manager.get_local_path(filename, create_if_missing=False)
            
            if not os.path.exists(local_path):
                self.send_error(404, f"Document '{filename}' not found on disk")
                return
            
            with open(local_path, 'rb') as f:
                content = f.read()
            
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(content)))
            self.send_cors_headers()
            self.end_headers()
            self.wfile.write(content)
            
            # Cleanup temp file
            manager.cleanup_temp(filename)
        except FileNotFoundError as e:
            self.send_error(404, f"Document not found: {str(e)}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_error(500, f"Error serving document: {str(e)}")
    
    def handle_template_upload(self):
        """Handle template file upload."""
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            if content_length == 0:
                self.send_error(400, "No file data received")
                return
            
            content_type = self.headers.get('Content-Type', '')
            
            # Handle raw binary upload (application/octet-stream or Word document MIME type)
            if 'multipart' not in content_type.lower():
                # Read the file data directly
                file_data = self.rfile.read(content_length)
                
                # Save as template
                template_path = template_tools.get_template_path()
                os.makedirs(os.path.dirname(template_path), exist_ok=True)
                
                with open(template_path, 'wb') as f:
                    f.write(file_data)
                
                self.send_response(200)
                self.send_cors_headers()
                self.end_headers()
                self.wfile.write(json.dumps({
                    "success": True,
                    "message": "Template uploaded successfully",
                    "template_path": template_path,
                    "size_bytes": len(file_data)
                }).encode('utf-8'))
                return
            
            # Handle multipart/form-data
            import cgi
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={'REQUEST_METHOD': 'POST', 'CONTENT_TYPE': content_type}
            )
            
            if 'file' not in form:
                self.send_error(400, "No file field in form data. Use field name 'file'")
                return
            
            file_item = form['file']
            if not file_item.filename:
                self.send_error(400, "No filename provided")
                return
            
            # Save the uploaded file as template
            template_path = template_tools.get_template_path()
            os.makedirs(os.path.dirname(template_path), exist_ok=True)
            
            if hasattr(file_item, 'file'):
                # File-like object
                with open(template_path, 'wb') as f:
                    f.write(file_item.file.read())
            else:
                # String data
                with open(template_path, 'wb') as f:
                    if isinstance(file_item.value, bytes):
                        f.write(file_item.value)
                    else:
                        f.write(file_item.value.encode('utf-8'))
            
            self.send_response(200)
            self.send_cors_headers()
            self.end_headers()
            self.wfile.write(json.dumps({
                "success": True,
                "message": f"Template '{file_item.filename}' uploaded successfully",
                "template_path": template_path
            }).encode('utf-8'))
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_error(500, f"Error uploading template: {str(e)}")
    
    def log_message(self, format, *args):
        """Override to use print instead of stderr."""
        print(f"{self.address_string()} - {format % args}")


def run_http_server():
    """Run the HTTP server."""
    port = int(os.getenv('PORT', 8000))
    host = os.getenv('HOST', '0.0.0.0')
    
    # Set BASE_URL if not already set
    global BASE_URL
    if not BASE_URL:
        # Try to get from Render environment
        render_service_url = os.getenv('RENDER_SERVICE_URL')
        if render_service_url:
            BASE_URL = render_service_url
        else:
            BASE_URL = f"http://{host}:{port}"
    
    server = HTTPServer((host, port), MCPHTTPHandler)
    print(f"Office Word MCP Server running on http://{host}:{port}")
    print(f"Documents directory: {DOCUMENTS_DIR}")
    print(f"Base URL: {BASE_URL}")
    print(f"MCP endpoint: http://{host}:{port}/mcp/stream")
    print(f"Documents endpoint: http://{host}:{port}/documents/")
    
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server...")
        server.shutdown()


if __name__ == "__main__":
    run_http_server()
