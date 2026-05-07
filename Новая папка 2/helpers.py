from database.models import User
from config import ADMIN_IDS

def has_access(user: User, required_amount: int = 0) -> bool:
    """Returns True if user can generate content (free trial OR enough balance OR admin)."""
    if user.id in ADMIN_IDS:
        return True
    if not user.free_used:
        return True
    return (user.balance or 0) >= required_amount


def is_free_trial(user: User) -> bool:
    """Returns True if user is about to use their free trial."""
    return not user.free_used


def format_price(amount) -> str:
    try:
        val = int(amount)
        return f"{val:,} so'm".replace(",", " ")
    except:
        return f"{amount} so'm"


import json
import os

def safe_topic(topic: str, max_len: int = 200) -> str:
    return topic.strip()[:max_len]

def get_template_path(template_id: str) -> str:
    """Returns the file path for a given template ID from templates.json."""
    try:
        # The scanner updates the one in webapp/data/
        path = "webapp/data/templates.json"
            
        with open(path, "r", encoding="utf-8") as f:
            templates = json.load(f)
            for t in templates:
                if t["id"] == template_id:
                    return t["file_path"]
    except Exception as e:
        print(f"Error loading template path: {e}")
    
    return f"templates/{template_id}.pptx" # Fallback
