import openpyxl
import json
import re
import os
import sys

EXCEL_FILE = "menu.xlsx"
OUTPUT_JS = "menu_data.js"
OUTPUT_JSON = "menu.json"

SHOP = {
    "name_ar": "أطياب السعيد",
    "subtitle_ar": "للعطور والبخور",
    "logo": "logo.avif",
    "currency_ar": "ر.س"
}


def normalize_size(size_str):
    """Normalize size string to a standard display format."""
    s = str(size_str or "").strip()
    if not s:
        return ""
    sl = s.lower()

    # Handle set patterns like "2 * 50 ml", "3*50ml", "2 × 25 ml"
    m = re.search(r'(\d+)\s*[*×x]\s*(\d+(?:\.\d+)?)\s*(?:ml|مل)', sl, re.IGNORECASE)
    if m:
        return f"{m.group(1)}×{m.group(2)} ml"

    # Single ml
    m = re.search(r'(\d+(?:\.\d+)?)\s*(?:ml|مل)', sl, re.IGNORECASE)
    if m:
        return f"{m.group(1)} ml"

    # Tola with optional number/fraction; bare "تولة" means 1 tola
    m = re.search(r'(ربع|نصف|نص|½|¼|\d+)?\s*(?:ال)?(?:تولة|توله)', sl, re.IGNORECASE)
    if m:
        val = m.group(1)
        if not val:
            return "1 تولة"
        mp = {"ربع": "ربع", "نصف": "نصف", "نص": "نصف", "½": "نصف", "¼": "ربع"}
        if val in mp:
            return f"{mp[val]} تولة"
        return f"{val} تولة"

    # Size column contains invalid data; let caller fall back to name extraction
    return ""


def extract_size_from_name(name):
    """Extract size from product name as fallback."""
    s = str(name or "").strip()
    if not s:
        return ""
    sl = s.lower()

    m = re.search(r'(\d+)\s*[*×x]\s*(\d+(?:\.\d+)?)\s*(?:ml|مل)', sl, re.IGNORECASE)
    if m:
        return f"{m.group(1)}×{m.group(2)} ml"

    m = re.search(r'(\d+(?:\.\d+)?)\s*(?:ml|مل)', sl, re.IGNORECASE)
    if m:
        return f"{m.group(1)} ml"

    # Tola with optional number/fraction; bare "تولة" means 1 tola
    m = re.search(r'(ربع|نصف|نص|½|¼|\d+)?\s*(?:ال)?(?:تولة|توله)', sl, re.IGNORECASE)
    if m:
        val = m.group(1)
        if not val:
            return "1 تولة"
        mp = {"ربع": "ربع", "نصف": "نصف", "نص": "نصف", "½": "نصف", "¼": "ربع"}
        if val in mp:
            return f"{mp[val]} تولة"
        return f"{val} تولة"

    return ""


def remove_size_patterns(name):
    """Remove size patterns from product name."""
    name = str(name or "").strip()

    # Remove set size patterns: 2×50 ml, 3 * 50 ml, etc.
    name = re.sub(r'\s*[-|]?\s*\(?\s*\d+\s*[*×x]\s*\d+(?:\.\d+)?\s*(?:ml|مل)\s*\)?', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\d+\s*[*×x]\s*\d+(?:\.\d+)?\s*(?:ml|مل)', '', name, flags=re.IGNORECASE)

    # Remove single ml patterns
    name = re.sub(r'\s*[-|]\s*\(?\s*\d+(?:\.\d+)?\s*(?:ml|مل)\s*\)?', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\(?\s*\d+(?:\.\d+)?\s*(?:ml|مل)\s*\)?', '', name, flags=re.IGNORECASE)

    # Remove tola patterns (with optional number/fraction and optional "ال")
    name = re.sub(r'\s*[-|]\s*\(?\s*(?:ربع|نصف|نص|½|¼|\d+)?\s*(?:ال)?(?:تولة|توله)\s*\)?', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\(?\s*(?:ربع|نصف|نص|½|¼|\d+)?\s*(?:ال)?(?:تولة|توله)\s*\)?', '', name, flags=re.IGNORECASE)

    return name


def base_name(name):
    """Remove size patterns and clean product name for grouping."""
    name = remove_size_patterns(name)

    # Remove redundant brand suffixes so variants like:
    #   "عطر نيش من أطياب السعيد" and "عطر نيش"
    #   "Atyab Alsaeed - Niche Perfume" and "Niche Perfume"
    # group together.
    name = re.sub(r'\s*من\s+أطياب\s+السعيد\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\bAtyab\s+Al-?Saeed\s*[-]\s*', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\bby\s+Atyab\s+Al-?Saeed\b', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\bAtyab\s+Al-?Saeed\b', '', name, flags=re.IGNORECASE)

    # For perfumes: truncate descriptive text after first separator (| or - or :)
    # so that "عطر بريف | رحلة..." and "عطر بريف- توليفة..." group together.
    if name.lower().startswith("عطر "):
        parts = re.split(r'\s*[|\-:]\s*', name, maxsplit=1)
        if len(parts) > 1 and len(parts[0].split()) >= 2:
            name = parts[0]

    # Clean up extra spaces and trailing punctuation
    name = re.sub(r'\s+', ' ', name).strip()
    name = re.sub(r'[\-:|]+$', '', name).strip()
    return name


def canonical_key(name):
    """Create a grouping key that ignores minor descriptive differences.

    For perfumes we keep 'عطر' + the core brand word, plus hair/body markers
    so that 'عطر ألماس ...' variants merge while 'عطر نيش للشعر' stays separate.
    """
    name = base_name(name)
    words = name.split()
    if not words:
        return name

    # For Arabic perfumes, use first 2 words + hair/body markers
    if words[0] == "عطر" and len(words) >= 2:
        core = [words[0], words[1]]
        hair_body = {"للشعر", "والشعر", "للجسم", "والجسم", "body", "hair"}
        for w in words[2:]:
            if w.lower() in hair_body:
                core.append(w)
        return " ".join(core)

    # For English perfumes, similar: first 2 words + hair/body markers
    if len(words) >= 2 and words[0].lower() in ("perfume", "niche", "mukhallat", "oud"):
        core = [words[0], words[1]]
        hair_body = {"body", "hair"}
        for w in words[2:]:
            if w.lower() in hair_body:
                core.append(w)
        return " ".join(core)

    return name


def is_available(val):
    if val is None:
        return False
    return str(val).strip().lower() in ("نعم", "yes", "true", "1", "y")


def size_sort_key(size_label):
    """Return numeric sort key for a size label."""
    s = str(size_label).lower()
    # Sets last
    m = re.search(r'(\d+)\s*[*×x]\s*(\d+(?:\.\d+)?)', s)
    if m:
        return 10000 + float(m.group(2)) * int(m.group(1))
    # Single ml
    m = re.search(r'(\d+(?:\.\d+)?)', s)
    if m:
        return float(m.group(1))
    # Tola fractions
    if "ربع" in s or "¼" in s:
        return 0.25
    if "نصف" in s or "½" in s:
        return 0.5
    if "تولة" in s or "توله" in s:
        return 1.0
    return 9999


def build():
    if not os.path.exists(EXCEL_FILE):
        print(f"Error: {EXCEL_FILE} not found.", file=sys.stderr)
        sys.exit(1)

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        print("Error: Excel file is empty.", file=sys.stderr)
        sys.exit(1)

    # Correct column indices based on actual Excel structure:
    # name_ar | name_en | quantity | categories_ar | categories_en | Size | price | sale_price | available
    groups = {}

    for r in rows[1:]:
        name_ar = str(r[0] or "").strip()
        name_en = str(r[1] or "").strip()
        cat_ar = str(r[3] or "").strip() or "أخرى"
        cat_en = str(r[4] or "").strip() or "Others"
        size_raw = str(r[5] or "").strip()
        price = r[7]  # sale_price
        available = is_available(r[8])

        if not name_ar:
            continue
        # Skip section/header rows that are just numbers
        if re.match(r"^\d+$", name_ar):
            continue

        # Prefer size from product name (more reliable) and fall back to Size column
        size = extract_size_from_name(name_ar) or extract_size_from_name(name_en) or normalize_size(size_raw)
        base_ar = base_name(name_ar)
        key_ar = canonical_key(name_ar)

        try:
            price_num = float(price) if price is not None else None
        except (ValueError, TypeError):
            price_num = None

        key = (cat_ar, key_ar.lower())
        if key not in groups:
            groups[key] = {
                "category": cat_ar,
                "name_ar": base_ar,
                "available": available,
                "sizes": []
            }
        else:
            # Keep the shortest/cleanest name as display name
            if len(base_ar) < len(groups[key]["name_ar"]):
                groups[key]["name_ar"] = base_ar

        if size and price_num is not None:
            groups[key]["sizes"].append({
                "size": size,
                "price": price_num
            })

    # Sort sizes within each group and remove duplicates
    items = []
    for g in groups.values():
        seen = set()
        unique = []
        for sz in g["sizes"]:
            if sz["size"] not in seen:
                seen.add(sz["size"])
                unique.append(sz)
        unique.sort(key=lambda x: size_sort_key(x["size"]))
        g["sizes"] = unique
        if g["sizes"]:
            items.append(g)

    # Sort items by category then name
    cat_order = {"العطور": 1, "أدهان العود والمخلطات": 2, "بخور طبيعي": 3, "الرشوش والمعاميل": 4}
    items.sort(key=lambda x: (cat_order.get(x["category"], 99), x["name_ar"]))

    data = {"shop": SHOP, "items": items}
    json_text = json.dumps(data, ensure_ascii=False, indent=2)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        f.write(json_text)

    js = "/* Auto-generated from menu.xlsx by build_menu.py */\n"
    js += "window.MENU_DATA = " + json_text + ";\n"
    with open(OUTPUT_JS, "w", encoding="utf-8") as f:
        f.write(js)

    cats = sorted(set(i["category"] for i in items), key=lambda c: cat_order.get(c, 99))
    print(f"Generated {OUTPUT_JSON} and {OUTPUT_JS} with {len(items)} items in {len(cats)} categories.")


if __name__ == "__main__":
    build()
