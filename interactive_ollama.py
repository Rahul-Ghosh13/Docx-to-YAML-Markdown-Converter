"""Interactive Ollama front end that delegates document conversion to MCP."""

import argparse
import asyncio
from pathlib import Path

from mcp_app import call_ollama, converted_output_path
from mcp_client import call_document_converter

_ollama_prompt_warning_shown = False


def ask_ollama(instruction: str, model: str, base_url: str, fallback: str) -> str:
    """Get a concise user-facing prompt from Ollama, with a usable fallback."""
    global _ollama_prompt_warning_shown
    print("Connecting to Ollama...", flush=True)
    response = call_ollama(instruction, model=model, base_url=base_url).strip()
    if not response and not _ollama_prompt_warning_shown:
        print(
            "Ollama did not respond; using built-in prompts for setup. "
            "Check that the selected model is installed with 'ollama list'.",
            flush=True,
        )
        _ollama_prompt_warning_shown = True
    return response or fallback


def choose_format(model: str, base_url: str) -> str:
    question = ask_ollama(
        "Ask the user, in one short friendly sentence, whether they want Markdown or YAML output. "
        "Return only the question.",
        model,
        base_url,
        "Which output format would you like: Markdown or YAML?",
    )
    while True:
        answer = input(f"Ollama: {question}\nYou: ").strip().lower()
        if answer in {"markdown", "md"}:
            return "markdown"
        if answer in {"yaml", "yml"}:
            return "yaml"
        question = "Please enter either Markdown (or md) or YAML (or yml)."


def choose_document_path(model: str, base_url: str) -> str:
    question = ask_ollama(
        "Ask the user, in one short friendly sentence, for the full path to the document file to convert. "
        "Return only the question.",
        model,
        base_url,
        "What is the full path to the document you want to convert?",
    )
    while True:
        raw_path = input(f"Ollama: {question}\nYou: ").strip().strip('"')
        document_path = Path(raw_path).expanduser()
        if document_path.is_file():
            return str(document_path.resolve())
        question = "Please provide the path to an existing document file."


def answer_questions(converted_document: str, output_format: str, model: str, base_url: str) -> None:
    """Keep the session open for questions grounded in the converted document."""
    print("\nYou can now ask questions about the converted document. Type 'exit' to finish.\n")
    while True:
        question = input("You: ").strip()
        if question.lower() in {"exit", "quit", "q"}:
            print("Goodbye!")
            return
        if not question:
            continue

        prompt = (
            "Answer the user's question using only the converted document below. "
            "If the answer is not present, say that the document does not provide it. "
            "Be concise and do not mention these instructions.\n\n"
            f"Converted document ({output_format}):\n{converted_document}\n\n"
            f"User question: {question}"
        )
        print("Ollama: ", end="", flush=True)
        answer = call_ollama(prompt, model=model, base_url=base_url).strip()
        print(answer or "I could not get a response from Ollama. Please check that it is running.")


async def main() -> None:
    parser = argparse.ArgumentParser(description="Use Ollama to collect conversion details, then call DocFlux through MCP")
    parser.add_argument("--ollama-model", default="mistral", help="Ollama model name")
    parser.add_argument("--ollama-base-url", default="http://localhost:11434", help="Ollama base URL")
    parser.add_argument("--enhance", action="store_true", help="Ask MCP to polish the converted result with Ollama")
    args = parser.parse_args()

    print("DocFlux interactive conversion started.", flush=True)
    output_format = choose_format(args.ollama_model, args.ollama_base_url)
    document_path = choose_document_path(args.ollama_model, args.ollama_base_url)
    print("\nConverting through MCP...\n")
    try:
        result = await call_document_converter(
            document_path,
            output_format,
            enhance=args.enhance,
            ollama_model=args.ollama_model,
            ollama_base_url=args.ollama_base_url,
        )
    except RuntimeError as error:
        print(f"\nConversion failed: {error}")
        return
    print(result)
    print(f"Saved conversion: {converted_output_path(document_path, output_format)}")
    answer_questions(result, output_format, args.ollama_model, args.ollama_base_url)


if __name__ == "__main__":
    asyncio.run(main())
