import re

def normalize_title(t: str) -> str:
    t = t.strip()
    t = re.sub(r"[^A-Za-z0-9._-]+", "_", t)
    return t or "Course"

def find_used_assets(blocks):
    assets = []
    for b in blocks:
        if b["type"] == "slide" and b.get("image"):
            # Just the basename is fine for placeholders
            name = b["image"].split("/")[-1].split("\\")[-1]
            if name:
                assets.append(name)
    return sorted(set(assets))
