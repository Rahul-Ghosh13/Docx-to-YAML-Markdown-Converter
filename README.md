# DocFlux MCP

DocFlux MCP is a Python document-conversion service that turns Word documents and common text files into structured Markdown or YAML. It can run as an MCP server, a command-line tool, or an interactive local-Ollama workflow.

## Features

- Converts `.docx` files to Markdown or YAML.
- Supports text-based input: `.txt`, `.md`, `.markdown`, `.rst`, `.csv`, `.json`, `.yaml`, `.yml`, and `.log`.
- Preserves document order, headings, lists, tables, inline formatting, code-like blocks, and embedded images where possible.
- Saves converted output next to the source document.
- Exposes `read_document`, `convert_document`, and `enhance_document` MCP tools over stdio.
- Optionally uses a local [Ollama](https://ollama.com/) model to polish a conversion while checking document fidelity.

## Requirements

- Python 3.10 or later
- `python-docx`
- An MCP Python package that provides `mcp.server.mcpserver.server.MCPServer`

Ollama is only required for the interactive assistant and the optional enhancement workflow.

## Installation

Clone the repository and create a virtual environment:

```bash
git clone <your-repository-url>
cd Task3
python -m venv .venv
```

Activate it:

```bash
# Windows PowerShell
.\.venv\Scripts\Activate.ps1

# macOS / Linux
source .venv/bin/activate
```

Install the dependencies:

```bash
pip install python-docx mcp
```

For optional Ollama-powered features, install and start Ollama, then pull a model such as Mistral:

```bash
ollama pull mistral
ollama serve
```

## Command-line usage

Convert a document to Markdown:

```bash
python app.py sample.docx --format markdown
```

Convert a document to YAML:

```bash
python app.py sample.docx --format yaml
```

The converted content is printed to the terminal and saved alongside the original file. For example, `sample.docx` produces `sample.md` or `sample.yaml`.

When converting a text file into the same extension (for example, Markdown to Markdown), DocFlux avoids overwriting the source and writes a sibling file such as `notes.converted.md`.

## Run as an MCP server

Start the stdio MCP server:

```bash
python mcp_server.py
```

Or run the server implementation directly:

```bash
python mcp_app.py
```

The server provides these tools:

| Tool | Description |
| --- | --- |
| `read_document` | Reads a supported document and returns normalized Markdown or YAML without saving a file. |
| `convert_document` | Converts a document, saves the converted result next to it, and returns the result. |
| `enhance_document` | Converts then improves readability with Ollama while preserving the source content. |

Each tool accepts `document_path` and `output_format` (`markdown` or `yaml`). Ollama-enabled tools also accept `ollama_model` (default: `mistral`) and `base_url` (default: `http://localhost:11434`).

## MCP client example

Use the included client to call the server through stdio:

```bash
python mcp_client.py sample.docx --format markdown
```

Enable the Ollama enhancement flow:

```bash
python mcp_client.py sample.docx --format markdown --enhance --ollama-model mistral
```

## Interactive Ollama workflow

The interactive interface asks for an output format and document path, converts through MCP, then lets you ask questions grounded in the converted result:

```bash
python interactive_ollama.py
```

To ask MCP to enhance the conversion with Ollama:

```bash
python interactive_ollama.py --enhance --ollama-model mistral
```

## Tests

Run the test suite from the repository root:

```bash
python -m unittest discover -s tests -v
```

The tests cover CLI conversion, MCP tool conversion, source-file protection, and the Ollama integration using a local test server.

## Project structure

```text
app.py                 Command-line converter
document_reader.py     Validated DOCX and text-file reader
mcp_app.py             Conversion logic and MCP tool definitions
mcp_server.py          MCP server entry point
mcp_client.py          Example MCP stdio client
interactive_ollama.py  Interactive Ollama-assisted workflow
tests/                 Automated tests
```

## License

No license has been specified yet. Add a `LICENSE` file before publishing if you want to define how others may use this project.
