import os
import docx

# Determine the absolute path to the knowledge base directory
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KNOWLEDGE_BASE_DIR = os.path.join(BASE_DIR, "knowledge_base")

def read_txt(filepath: str) -> str:
    """Reads .txt and .md files."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"Error reading text file {filepath}: {e}")
        return ""

def read_docx(filepath: str) -> str:
    """Reads .docx files using python-docx."""
    try:
        doc = docx.Document(filepath)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        print(f"Error reading docx file {filepath}: {e}")
        return ""

def load_knowledge(category: str) -> str:
    """
    Loads all knowledge base files for a given category (e.g., 'thesis', 'article', 'presentation').
    Returns a combined string of instructions formatted for the AI prompt.
    """
    category_dir = os.path.join(KNOWLEDGE_BASE_DIR, category)
    
    if not os.path.exists(category_dir):
        # Category folder doesn't exist yet, return empty context
        return ""
        
    combined_knowledge = []
    
    # Handlers for supported extensions
    handlers = {
        ".txt": read_txt,
        ".md": read_txt,
        ".docx": read_docx
    }
    
    # Sort files to ensure predictable order (e.g., 1_rules.md, 2_guide.docx)
    try:
        files = sorted(os.listdir(category_dir))
    except Exception as e:
        print(f"Error reading directory {category_dir}: {e}")
        return ""
    
    for filename in files:
        filepath = os.path.join(category_dir, filename)
        if not os.path.isfile(filepath):
            continue
            
        ext = os.path.splitext(filename)[1].lower()
        if ext in handlers:
            content = handlers[ext](filepath).strip()
            if content:
                # Format block title based on filename (e.g., "format_rules" -> "FORMAT RULES")
                title = os.path.splitext(filename)[0].replace('_', ' ').replace('-', ' ').upper()
                
                block = f"[{title}]\n{content}\n"
                combined_knowledge.append(block)
                
    if not combined_knowledge:
        return ""
        
    final_context = (
        "=== INTERNAL KNOWLEDGE BASE RULES ===\n"
        "Follow these institutional rules and internal standards strictly:\n\n"
        + "\n".join(combined_knowledge) +
        "=====================================\n"
    )
    
    return final_context
