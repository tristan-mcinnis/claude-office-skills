# PowerPoint Editing Skill

Edit existing PowerPoint presentations. For creating new presentations, use the `pptx` skill instead.

## When to Use This Skill

Use this skill when:
- Updating text content in an existing deck
- Changing fonts, colors, or sizes throughout a presentation
- Standardizing formatting across slides
- Batch find-and-replace operations
- Fixing inconsistent styling
- Updating brand elements (colors, fonts)

Do NOT use this skill when:
- Creating a new presentation from scratch → use `pptx` skill
- Building from a template with new content → use `pptx` skill
- Converting HTML to PowerPoint → use `pptx` skill

---

## Quick Reference: Common Tasks

| Task | Command |
|------|---------|
| See all text | `venv/bin/python -m markitdown file.pptx` |
| Extract text inventory | `venv/bin/python scripts/inventory.py file.pptx inventory.json` |
| Replace text | `venv/bin/python scripts/replace.py file.pptx replacements.json output.pptx` |
| Visual preview | `venv/bin/python scripts/thumbnail.py file.pptx thumbnails` |

---

## Workflow 1: Text Replacement (Most Common)

**Use for:** Updating text content while preserving all formatting.

### Step 1: Extract Text Inventory

```bash
venv/bin/python public/pptx/scripts/inventory.py presentation.pptx inventory.json
```

This creates a JSON map of all text shapes with:
- Slide and shape IDs
- Current text content
- Position (left, top, width, height)
- Formatting (font, size, color, bold, italic, bullets)

### Step 2: Read the Entire Inventory

**MANDATORY:** Read `inventory.json` completely before creating replacements. Never set range limits.

The inventory structure:
```json
{
  "slide-0": {
    "shape-0": {
      "left": 0.5,
      "top": 0.3,
      "width": 12.3,
      "height": 1.0,
      "placeholder_type": "TITLE",
      "paragraphs": [
        {"text": "Current Title", "bold": true, "font_size": 44}
      ]
    }
  }
}
```

### Step 3: Create Replacement JSON

Create a JSON file specifying what to replace:

```json
{
  "slide-0": {
    "shape-0": {
      "paragraphs": [
        {"text": "New Title", "bold": true, "alignment": "CENTER"}
      ]
    }
  },
  "slide-1": {
    "shape-2": {
      "paragraphs": [
        {"text": "First bullet", "bullet": true, "level": 0},
        {"text": "Second bullet", "bullet": true, "level": 0},
        {"text": "Sub-bullet", "bullet": true, "level": 1}
      ]
    }
  }
}
```

### Step 4: Apply Replacements

```bash
venv/bin/python public/pptx/scripts/replace.py presentation.pptx replacements.json output.pptx
```

### CRITICAL Rules for Text Replacement

1. **Shapes not in replacement JSON are CLEARED** - Only include shapes you want to keep content in
2. **Bullets need `bullet: true` and `level: 0`** - Don't include bullet symbols in text
3. **Bullets force LEFT alignment** - Don't override with `alignment` when using bullets
4. **Run inventory FIRST** - Never guess at shape IDs

---

## Workflow 2: Batch Style Changes (Fonts, Sizes, Colors)

**Use for:** "Change all headlines to 22pt" or "Update brand color throughout"

### Step 1: Inventory Current Styles

First, understand what styles exist in the presentation:

```bash
venv/bin/python public/pptx/scripts/inventory.py presentation.pptx inventory.json
```

Analyze the inventory to identify:
- What font sizes are used for headlines vs body
- What colors appear throughout
- Which shapes need to change

### Step 2: Plan Your Changes

**MANDATORY:** Before any XML editing, document your plan:

```python
# Example change plan
changes = {
    "font_size": {
        "target": "headlines",  # shapes with placeholder_type TITLE or CENTER_TITLE
        "from": "any",
        "to": 22
    },
    "color": {
        "target": "all_text",
        "from": "1234AB",
        "to": "5678CD"
    }
}
```

### Step 3: Choose Your Approach

**Option A: Simple cases → Use replace.py**

If you're only changing a few specific shapes, generate a replacement JSON and use the text replacement workflow.

**Option B: Batch changes → Use OOXML workflow**

For changes across many shapes/slides, use direct XML editing:

1. Unpack the presentation:
```bash
venv/bin/python public/pptx/ooxml/scripts/unpack.py presentation.pptx unpacked/
```

2. Write a Python script to modify the XML (see "XML Editing Patterns" below)

3. Validate:
```bash
venv/bin/python public/pptx/ooxml/scripts/validate.py unpacked/ --original presentation.pptx
```

4. Pack only if validation passes:
```bash
venv/bin/python public/pptx/ooxml/scripts/pack.py unpacked/ output.pptx
```

---

## Workflow 3: Find and Replace Text

**Use for:** Replacing specific text patterns throughout (e.g., "Q3" → "Q4", "Old Corp" → "New Corp")

### Simple Approach (Small Decks)

1. Extract inventory
2. Search inventory JSON for target text
3. Create replacement JSON with updated text
4. Apply replacements

### Pattern-Based Approach (Large Decks)

Write a script to automate the find/replace:

```python
import json
import re

def find_replace_inventory(inventory_path, find_pattern, replace_with, output_path):
    with open(inventory_path) as f:
        inventory = json.load(f)

    replacements = {}

    for slide_id, shapes in inventory.items():
        if not slide_id.startswith("slide-"):
            continue
        for shape_id, shape_data in shapes.items():
            paragraphs = shape_data.get("paragraphs", [])
            new_paragraphs = []
            changed = False

            for para in paragraphs:
                text = para.get("text", "")
                new_text = re.sub(find_pattern, replace_with, text)
                if new_text != text:
                    changed = True
                new_para = para.copy()
                new_para["text"] = new_text
                new_paragraphs.append(new_para)

            if changed:
                replacements.setdefault(slide_id, {})[shape_id] = {
                    "paragraphs": new_paragraphs
                }

    with open(output_path, "w") as f:
        json.dump(replacements, f, indent=2)

    return len(replacements)

# Usage
count = find_replace_inventory("inventory.json", r"Q3 2024", "Q4 2024", "replacements.json")
print(f"Found {count} shapes to update")
```

---

## XML Editing Patterns

When the replacement scripts don't cover your needs, edit XML directly.

### MANDATORY: Inventory → Plan → Execute → Validate

1. **Inventory:** Always run analysis BEFORE editing
2. **Plan:** Document exactly what you'll change (as JSON/dict)
3. **Execute:** Use ElementTree, NEVER string replacement
4. **Validate:** Run validation BEFORE packing

### Key Files in PPTX Structure

```
unpacked/
├── [Content_Types].xml      # File type declarations
├── ppt/
│   ├── presentation.xml     # Main metadata, slide order
│   ├── slides/
│   │   ├── slide1.xml       # Slide content
│   │   ├── slide2.xml
│   │   └── ...
│   ├── slideLayouts/        # Layout templates
│   ├── slideMasters/        # Master slides
│   └── theme/               # Colors and fonts
└── _rels/                   # Relationships
```

### Common XML Edits

**Change font size:**
```xml
<!-- Find: -->
<a:rPr sz="4400"/>
<!-- sz is in hundredths of a point, so 4400 = 44pt -->

<!-- Change to 22pt: -->
<a:rPr sz="2200"/>
```

**Change color:**
```xml
<!-- Find: -->
<a:solidFill>
  <a:srgbClr val="1234AB"/>
</a:solidFill>

<!-- Replace with: -->
<a:solidFill>
  <a:srgbClr val="5678CD"/>
</a:solidFill>
```

**Change font family:**
```xml
<!-- Find: -->
<a:latin typeface="Arial"/>

<!-- Replace with: -->
<a:latin typeface="Calibri"/>
```

### XML Editing Script Template

```python
import xml.etree.ElementTree as ET
from pathlib import Path

# Namespaces used in PPTX
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

def change_font_size(slide_path, old_size, new_size):
    """Change font sizes in a slide XML file.

    Args:
        slide_path: Path to slide XML file
        old_size: Size to find (in points, e.g., 44)
        new_size: Size to replace with (in points, e.g., 22)
    """
    tree = ET.parse(slide_path)
    root = tree.getroot()

    # Register namespaces to preserve them on write
    for prefix, uri in NS.items():
        ET.register_namespace(prefix, uri)

    old_sz = str(old_size * 100)  # Convert to hundredths
    new_sz = str(new_size * 100)

    # Find all text run properties
    for rPr in root.iter('{%s}rPr' % NS['a']):
        if rPr.get('sz') == old_sz:
            rPr.set('sz', new_sz)

    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')

# Usage
for slide in Path("unpacked/ppt/slides").glob("slide*.xml"):
    change_font_size(slide, old_size=44, new_size=22)
```

### FORBIDDEN Practices

- ❌ **NEVER** use sed/awk/string replacement on XML files
- ❌ **NEVER** edit without running inventory first
- ❌ **NEVER** pack without validation
- ❌ **NEVER** guess at element IDs or structure
- ❌ **NEVER** remove elements without understanding relationships

---

## Validation

**ALWAYS validate after any XML edit:**

```bash
venv/bin/python public/pptx/ooxml/scripts/validate.py unpacked/ --original presentation.pptx
```

Validation checks:
- XML schema compliance
- Required elements present
- Relationships intact
- No orphaned references

**Only pack after validation passes:**

```bash
venv/bin/python public/pptx/ooxml/scripts/pack.py unpacked/ output.pptx
```

---

## Troubleshooting

### "File is corrupt" after editing

1. Did you validate before packing? Run validation to see errors.
2. Did you use string replacement instead of XML parsing?
3. Did you remove a required element or break a relationship?

### Text replacement didn't work

1. Did you read the full inventory first?
2. Are your shape IDs correct? (They're 0-indexed)
3. Did you include ALL shapes you want to keep? (Missing shapes are cleared)

### Colors didn't change

1. Check if color is defined as theme color vs RGB
2. Theme colors are in `ppt/theme/theme1.xml`
3. RGB colors are inline in shape definitions

### Fonts didn't change everywhere

1. Some text may inherit from master slide
2. Check both `slide*.xml` and `slideMaster*.xml`
3. Font may be defined at paragraph level vs run level

---

## Output Directory Convention

All edited files go to:
```
outputs/<document-name>/
```

Example workflow:
```bash
mkdir -p outputs/q4-update/
venv/bin/python scripts/inventory.py deck.pptx outputs/q4-update/inventory.json
# ... create replacements.json ...
venv/bin/python scripts/replace.py deck.pptx outputs/q4-update/replacements.json outputs/q4-update/deck-updated.pptx
```
