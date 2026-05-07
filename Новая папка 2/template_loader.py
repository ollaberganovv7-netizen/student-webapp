"""
template_loader.py

Loads document structure templates (doc_templates/) and quality examples (doc_examples/)
to inject into AI prompts for consistent, high-quality document generation.

Folder layout:
    doc_templates/
        referat_structure.txt
        tezis_structure.txt
        coursework_structure.txt
        presentation_structure.txt

    doc_examples/
        referat_example.txt
        tezis_example.txt
        coursework_example.txt
        presentation_example.txt

Supported file formats: .txt, .md, .docx
"""

import os
import logging

logger = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(BASE_DIR, "doc_templates")
EXAMPLES_DIR = os.path.join(BASE_DIR, "doc_examples")


# ── File readers ─────────────────────────────────────────────────────────────

def _read_txt(filepath: str) -> str:
    """Read a plain .txt or .md file."""
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        logger.warning(f"template_loader: cannot read txt '{filepath}': {e}")
        return ""


def _read_docx(filepath: str) -> str:
    """Read a .docx file using python-docx."""
    try:
        import docx
        doc = docx.Document(filepath)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except ImportError:
        logger.error("template_loader: python-docx is not installed.")
        return ""
    except Exception as e:
        logger.warning(f"template_loader: cannot read docx '{filepath}': {e}")
        return ""


_READERS = {
    ".txt": _read_txt,
    ".md":  _read_txt,
    ".docx": _read_docx,
}


def _read_file(filepath: str) -> str:
    """Dispatch to the correct reader based on file extension."""
    ext = os.path.splitext(filepath)[1].lower()
    reader = _READERS.get(ext)
    if reader:
        return reader(filepath).strip()
    return ""


# ── Core loaders ──────────────────────────────────────────────────────────────

def _find_file(directory: str, doc_type: str) -> str | None:
    """
    Find a file inside *directory* whose base name starts with *doc_type*
    and has a supported extension.

    Examples:
        doc_type="tezis" matches "tezis_structure.txt", "tezis.docx", etc.
    """
    if not os.path.isdir(directory):
        return None

    for filename in sorted(os.listdir(directory)):
        name_no_ext, ext = os.path.splitext(filename)
        if ext.lower() in _READERS and name_no_ext.lower().startswith(doc_type.lower()):
            return os.path.join(directory, filename)

    return None


def load_template(doc_type: str) -> str:
    """
    Load the structure template for *doc_type* (e.g. 'tezis', 'referat', 'coursework').

    Returns a formatted string ready to embed in an AI prompt, or "" if not found.
    """
    filepath = _find_file(TEMPLATES_DIR, doc_type)

    if filepath is None:
        logger.info(f"template_loader: no template found for '{doc_type}' in {TEMPLATES_DIR}")
        return ""

    content = _read_file(filepath)
    if not content:
        logger.warning(f"template_loader: template file '{filepath}' is empty.")
        return ""

    logger.info(f"template_loader: loaded template '{filepath}' ({len(content)} chars)")
    return (
        "=== DOCUMENT STRUCTURE TEMPLATE ===\n"
        "Strictly follow this structure and formatting rules when creating the document:\n\n"
        f"{content}\n"
        "====================================\n"
    )


def load_example(doc_type: str) -> str:
    """
    Load the quality example for *doc_type*.

    Returns a formatted string ready to embed in an AI prompt, or "" if not found.

    IMPORTANT: The example is a quality benchmark — use it to understand the expected
    level of detail, style, and formatting. Do NOT copy it verbatim.

    NOTE: Examples are truncated to MAX_EXAMPLE_CHARS to avoid exceeding
    the OpenAI context window / rate limits.
    """
    # Max characters from example to inject into prompt (~3000 tokens)
    MAX_EXAMPLE_CHARS = 4000

    filepath = _find_file(EXAMPLES_DIR, doc_type)

    if filepath is None:
        logger.info(f"template_loader: no example found for '{doc_type}' in {EXAMPLES_DIR}")
        return ""

    content = _read_file(filepath)
    if not content:
        logger.warning(f"template_loader: example file '{filepath}' is empty.")
        return ""

    original_len = len(content)

    # Truncate to safe limit and cut at last paragraph boundary
    if len(content) > MAX_EXAMPLE_CHARS:
        content = content[:MAX_EXAMPLE_CHARS]
        # Cut at last newline to avoid mid-sentence truncation
        last_newline = content.rfind("\n")
        if last_newline > MAX_EXAMPLE_CHARS // 2:
            content = content[:last_newline]
        content += "\n[... example truncated for context efficiency ...]"

    logger.info(
        f"template_loader: loaded example '{filepath}' "
        f"({original_len} chars total, {len(content)} chars used)"
    )
    return (
        "=== QUALITY EXAMPLE (benchmark only) ===\n"
        "Study the structure, style, depth, and academic tone of this example carefully.\n"
        "DO NOT copy it word for word — use it solely as a quality and format benchmark.\n\n"
        f"{content}\n"
        "=========================================\n"
    )



def load_template_and_example(doc_type: str) -> str:
    """
    Convenience function: loads both template + example and merges them
    into a single context string for the AI prompt.

    If neither exists, returns "".
    """
    template = load_template(doc_type)
    example  = load_example(doc_type)

    parts = [p for p in (template, example) if p]
    if not parts:
        return ""

    return "\n".join(parts)
