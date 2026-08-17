from mcp_app import (
    call_ollama,
    convert_document_file,
    convert_document_tool,
    read_document_file,
    read_document_tool,
    enhance_document_tool,
    create_server,
    extract_blocks,
    to_markdown,
    to_yaml,
)

__all__ = [
    "call_ollama",
    "convert_document_file",
    "convert_document_tool",
    "read_document_file",
    "read_document_tool",
    "enhance_document_tool",
    "create_server",
    "extract_blocks",
    "to_markdown",
    "to_yaml",
]


if __name__ == "__main__":
    server = create_server()
    server.run(transport="stdio")
