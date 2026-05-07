import os
import json
import shutil

def update_templates_json(templates_dir="templates", output_path="webapp/data/templates.json"):
    """
    Plug-and-play scanner. 
    Supports:
    - templates/file.pptx + templates/file.png
    - templates/Category/file.pptx + templates/Category/file.png
    """
    templates = []
    
    if not os.path.exists(templates_dir):
        os.makedirs(templates_dir, exist_ok=True)
        return

    # Helper to process a directory
    def process_dir(directory, category_name):
        files = os.listdir(directory)
        pptx_files = [f for f in files if f.lower().endswith(".pptx")]
        
        for pptx in pptx_files:
            base_name = os.path.splitext(pptx)[0]
            # Find matching preview
            preview_file = None
            for ext in [".png", ".jpg", ".jpeg"]:
                img_name = base_name + ext
                # Case insensitive check
                match = next((f for f in files if f.lower() == img_name.lower()), None)
                if match:
                    # Found preview! Copy to webapp static
                    src_img = os.path.join(directory, match)
                    dest_dir = f"webapp/static/previews/{category_name}"
                    os.makedirs(dest_dir, exist_ok=True)
                    dest_img = os.path.join(dest_dir, match)
                    shutil.copy2(src_img, dest_img)
                    preview_file = f"static/previews/{category_name}/{match}"
                    break
            
            template_id = f"{category_name.lower()}_{base_name.lower()}".replace(" ", "_")
            templates.append({
                "id": template_id,
                "name": base_name.replace("_", " ").title(),
                "category": category_name,
                "preview": preview_file or "https://via.placeholder.com/280x180?text=No+Preview",
                "file_path": os.path.join(directory, pptx),
                "file_size": f"{os.path.getsize(os.path.join(directory, pptx)) // 1024} KB"
            })
        
        return len(pptx_files)

    # 1. Process root templates folder (Category: General) if any files are there
    process_dir(templates_dir, "General")

    # 2. Process subfolders
    for item in sorted(os.listdir(templates_dir)):
        path = os.path.join(templates_dir, item)
        if os.path.isdir(path) and item.lower() not in ["__pycache__"]:
            # If the folder is literally named 'General', we process it as 'General'
            # But avoid adding dummy placeholders if it's empty
            count = process_dir(path, item)
            # Add placeholder for empty categories so they show in catalog
            if count == 0:
                templates.append({
                    "id": f"dummy_{item.lower()}",
                    "name": "Tez orada...",
                    "category": item,
                    "preview": f"https://via.placeholder.com/200x120/1c1c1e/fff?text={item}",
                    "file_path": "",
                    "file_size": "0 KB",
                    "is_empty": True
                })

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(templates, f, indent=4, ensure_ascii=False)
        
    print(f"Successfully updated {output_path} with {len(templates)} templates.")
    return templates

if __name__ == "__main__":
    update_templates_json()
