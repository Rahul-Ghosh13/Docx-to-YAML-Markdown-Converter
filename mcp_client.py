import argparse
import asyncio
import sys
from pathlib import Path
from mcp.client.stdio import StdioServerParameters, stdio_client
from mcp import ClientSession


async def call_document_converter(
    document_path: str,
    output_format: str = "markdown",
    *,
    enhance: bool = False,
    ollama_model: str = "mistral",
    ollama_base_url: str = "http://localhost:11434",
) -> str:
    """Call a DocFlux conversion tool through its MCP stdio server."""
    server_script = Path(__file__).with_name("mcp_app.py")
    server_params = StdioServerParameters(command=sys.executable, args=[str(server_script)])
    tool_error: str | None = None

    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tool_name = "enhance_document" if enhance else "convert_document"
            tool_arguments = {
                "document_path": document_path,
                "output_format": output_format,
                "ollama_model": ollama_model,
                "base_url": ollama_base_url,
            }
            if enhance:
                tool_arguments.update({
                    "ollama_model": ollama_model,
                    "base_url": ollama_base_url,
                })
            result = await session.call_tool(tool_name, tool_arguments)
            if result.is_error:
                tool_error = result.content[0].text if result.content else "MCP conversion failed."
            else:
                return result.content[0].text

    raise RuntimeError(tool_error or "MCP conversion failed.")


async def main() -> None:
    parser = argparse.ArgumentParser(description="Call the DocFlux MCP server")
    parser.add_argument("document_path", help="Path to the supported document file")
    parser.add_argument("--format", choices=["markdown", "yaml"], default="markdown", help="Output format")
    parser.add_argument("--enhance", action="store_true", help="Use the MCP server's Ollama enhancement tool")
    parser.add_argument("--ollama-model", default="mistral", help="Ollama model name")
    parser.add_argument("--ollama-base-url", default="http://localhost:11434", help="Ollama base URL")
    args = parser.parse_args()

    result = await call_document_converter(
        args.document_path,
        args.format,
        enhance=args.enhance,
        ollama_model=args.ollama_model,
        ollama_base_url=args.ollama_base_url,
    )
    print(result)


if __name__ == "__main__":
    asyncio.run(main())
