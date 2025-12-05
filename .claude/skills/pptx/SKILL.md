---
name: pptx
description: "Presentation creation, editing, and analysis. When Claude needs to work with presentations (.pptx files) for: (1) Creating new presentations, (2) Modifying or editing content, (3) Extracting or analyzing text, (4) Working with layouts, (5) Batch style changes (fonts, sizes, colors), (6) Find and replace operations, (7) Adding comments or speaker notes, or any other presentation tasks"
---

# PPTX Skill

Create, edit, and analyze PowerPoint presentations.

## Quick Reference: Common Tasks

| Task | Command |
|------|---------|
| Extract all text | `venv/bin/python -m markitdown file.pptx` |
| Extract text inventory | `venv/bin/python .claude/skills/pptx/scripts/inventory.py file.pptx inventory.json` |
| Replace text | `venv/bin/python .claude/skills/pptx/scripts/replace.py file.pptx replacements.json output.pptx` |
| Visual preview | `venv/bin/python .claude/skills/pptx/scripts/thumbnail.py file.pptx thumbnails` |
| Rearrange slides | `venv/bin/python .claude/skills/pptx/scripts/rearrange.py file.pptx output.pptx 0,2,2,5` |

---

## Workflow Decision Tree

**What do you need to do?**

| Situation | Workflow |
|-----------|----------|
| Read/analyze text content | Text Extraction (below) |
| Extract structured data | Text Inventory (Workflow 1) |
| Update text in existing deck | Text Replacement (Workflow 1) |
| Change fonts, colors, sizes throughout | Batch Style Changes (Workflow 2) |
| Apply new outline to existing deck/template | Apply Outline (Workflow 3) |
| Find and replace specific text | Find & Replace (Workflow 4) |
| Create new presentation from scratch | HTML Creation (Workflow 5) |

---

## Text Extraction & Analysis

### Quick Text Extraction

Convert presentation to markdown for reading:

```bash
venv/bin/python -m markitdown file.pptx
```

### Raw XML Access

For comments, speaker notes, slide layouts, animations, and complex formatting, unpack the presentation:

```bash
venv/bin/python .claude/skills/pptx/ooxml/scripts/unpack.py file.pptx unpacked/
```

**Key file structures:**
* `ppt/presentation.xml` - Main presentation metadata and slide references
* `ppt/slides/slide{N}.xml` - Individual slide contents
* `ppt/notesSlides/notesSlide{N}.xml` - Speaker notes
* `ppt/comments/modernComment_*.xml` - Comments
* `ppt/slideLayouts/` - Layout templates
* `ppt/slideMasters/` - Master slides
* `ppt/theme/` - Theme and styling (colors, fonts)
* `ppt/media/` - Images and media files

### Typography and Color Extraction

When analyzing an existing design:
1. Read theme: `ppt/theme/theme1.xml` for colors (`<a:clrScheme>`) and fonts (`<a:fontScheme>`)
2. Sample slides: Check `ppt/slides/slide1.xml` for actual usage
3. Search patterns: grep for `<a:solidFill>`, `<a:srgbClr>`, font references

---

## Workflow 1: Text Replacement

**Use for:** Updating text content while preserving formatting.

### Step 1: Extract Text Inventory

```bash
venv/bin/python .claude/skills/pptx/scripts/inventory.py presentation.pptx inventory.json
```

Creates JSON with all text shapes, positions, and formatting.

### Step 2: Read the Entire Inventory

**MANDATORY:** Read `inventory.json` completely before creating replacements.

### Step 3: Create Replacement JSON

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
# Full mode (default): clears ALL shapes, fills only those with "paragraphs" in JSON
venv/bin/python .claude/skills/pptx/scripts/replace.py presentation.pptx replacements.json output.pptx

# Selective mode: only modifies shapes listed in JSON, preserves unlisted shapes
venv/bin/python .claude/skills/pptx/scripts/replace.py presentation.pptx replacements.json output.pptx --selective
```

### CRITICAL Rules

1. **Full mode (default):** Shapes not in replacement JSON are CLEARED - Include all shapes you want to keep
2. **Selective mode (`--selective`):** Only shapes in JSON are modified - Use for targeted edits like changing just a title
3. **Bullets need `bullet: true` and `level: 0`** - Don't include bullet symbols in text
4. **Bullets force LEFT alignment** - Don't override with `alignment`
5. **Run inventory FIRST** - Never guess at shape IDs

---

## Workflow 2: Batch Style Changes

**Use for:** "Change all headlines to 22pt" or "Update brand color throughout"

### Step 1: Inventory Current Styles

```bash
venv/bin/python .claude/skills/pptx/scripts/inventory.py presentation.pptx inventory.json
```

Analyze to identify font sizes, colors, and which shapes need changes.

### Step 2: Choose Approach

**Option A: Few shapes → Use replace.py**

Generate replacement JSON with new formatting.

**Option B: Many shapes → Use OOXML workflow**

```bash
# Unpack
venv/bin/python .claude/skills/pptx/ooxml/scripts/unpack.py presentation.pptx unpacked/

# Edit XML (see XML Editing Patterns below)

# Validate
venv/bin/python .claude/skills/pptx/ooxml/scripts/validate.py unpacked/ --original presentation.pptx

# Pack (only if validation passes)
venv/bin/python .claude/skills/pptx/ooxml/scripts/pack.py unpacked/ output.pptx
```

### Common XML Edits

**Change font size** (sz is hundredths of a point):
```xml
<a:rPr sz="4400"/>  <!-- 44pt -->
<a:rPr sz="2200"/>  <!-- 22pt -->
```

**Change color:**
```xml
<a:solidFill>
  <a:srgbClr val="1234AB"/>
</a:solidFill>
```

**Change font:**
```xml
<a:latin typeface="Calibri"/>
```

---

## Workflow 3: Apply New Outline to Existing Deck

**Use for:** Populating an existing deck/template with new content.

### Step 1: Analyze the Existing Deck

```bash
# Extract text
venv/bin/python -m markitdown existing.pptx > outputs/project/deck-text.md

# Create thumbnails
venv/bin/python .claude/skills/pptx/scripts/thumbnail.py existing.pptx outputs/project/thumbnails

# Get inventory
venv/bin/python .claude/skills/pptx/scripts/inventory.py existing.pptx outputs/project/inventory.json
```

### Step 2: Create Slide Layout Inventory

**MANDATORY:** Document ALL available slide layouts before mapping content.

```markdown
| Index | Layout Type | Placeholders | Best For |
|-------|-------------|--------------|----------|
| 0 | Title slide | title, subtitle | Opening |
| 1 | Section header | title only | Section breaks |
| 2 | Title + bullets | title, body | Lists |
| 3 | Two-column | title, left, right | 2 items |
| 4 | Three-column | title, 3 cols | 3 items |
```

### Step 3: Map Content to Layouts

**CRITICAL LAYOUT MATCHING:**

| Content | Correct Layout | WRONG |
|---------|----------------|-------|
| 2 items | Two-column | Three-column (empty!) |
| 3 items | Three-column | Two-column (overflow!) |
| 5+ bullets | Single-column bullets | Force into columns |
| No image | Text-only layout | Image layout (empty!) |

### Step 4: Rearrange Slides

```bash
venv/bin/python .claude/skills/pptx/scripts/rearrange.py existing.pptx outputs/project/working.pptx 0,2,4,3,2
```

Numbers are slide indices (0-based), can repeat.

### Step 5: Get Fresh Inventory

```bash
venv/bin/python .claude/skills/pptx/scripts/inventory.py outputs/project/working.pptx outputs/project/working-inventory.json
```

**MANDATORY:** Re-read inventory after rearrangement - shape IDs may change.

### Step 6: Create and Apply Replacements

Create replacement JSON, then:

```bash
venv/bin/python .claude/skills/pptx/scripts/replace.py outputs/project/working.pptx outputs/project/replacements.json outputs/project/final.pptx
```

### Step 7: Visual Verification

```bash
venv/bin/python .claude/skills/pptx/scripts/thumbnail.py outputs/project/final.pptx outputs/project/final-thumbnails
```

---

## Workflow 4: Find and Replace Text

**Use for:** "Q3" → "Q4", "Old Corp" → "New Corp"

### Simple (Small Decks)

1. Extract inventory
2. Search for target text
3. Create replacement JSON
4. Apply replacements

### Pattern-Based (Large Decks)

```python
import json, re

def find_replace(inventory_path, find_pattern, replace_with, output_path):
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
                replacements.setdefault(slide_id, {})[shape_id] = {"paragraphs": new_paragraphs}

    with open(output_path, "w") as f:
        json.dump(replacements, f, indent=2)
    return len(replacements)

count = find_replace("inventory.json", r"Q3 2024", "Q4 2024", "replacements.json")
```

---

## Workflow 5: Create New Presentation (HTML)

**Use for:** Creating presentations from scratch with custom design.

### Design Principles

**CRITICAL**: Before creating any presentation:
1. Consider the subject matter - what tone/mood?
2. Check for branding requirements
3. Select colors that match content
4. State your design approach BEFORE writing code

**Requirements:**
- Use web-safe fonts only: Arial, Helvetica, Times New Roman, Georgia, Courier New, Verdana, Tahoma, Trebuchet MS, Impact
- Create clear visual hierarchy
- Ensure readability with strong contrast
- Be consistent across slides

### Color Palette Examples

1. **Classic Blue**: Navy #1C2833, slate #2E4053, silver #AAB7B8
2. **Teal & Coral**: Teal #5EA8A7, coral #FE4447
3. **Burgundy Luxury**: Burgundy #5D1D2E, gold #997929
4. **Black & Gold**: Gold #BF9A4A, black #000000
5. **Forest Green**: Green #4E9F3D, black #191A19

### Workflow

1. **MANDATORY**: Read [`html2pptx.md`](html2pptx.md) completely
2. Create HTML file for each slide (720pt × 405pt for 16:9)
3. Use `<p>`, `<h1>`-`<h6>`, `<ul>`, `<ol>` for text
4. Use `class="placeholder"` for charts/tables
5. **Rasterize gradients/icons as PNG first** using Sharp
6. Run JavaScript with html2pptx.js library
7. **Visual validation**: Generate thumbnails and check for issues
8. Iterate until visually correct

---

## Creating Thumbnail Grids

```bash
venv/bin/python .claude/skills/pptx/scripts/thumbnail.py template.pptx output_prefix
```

- Creates `thumbnails.jpg` (or numbered for large decks)
- Default: 5 columns, max 30 slides per grid
- Custom columns: `--cols 4`
- Slides are 0-indexed

---

## XML Editing Patterns

### MANDATORY Process

1. **Inventory** - Always analyze BEFORE editing
2. **Plan** - Document exactly what you'll change
3. **Execute** - Use ElementTree, NEVER string replacement
4. **Validate** - Run validation BEFORE packing

### Script Template

```python
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

def change_font_size(slide_path, old_size, new_size):
    tree = ET.parse(slide_path)
    root = tree.getroot()
    for prefix, uri in NS.items():
        ET.register_namespace(prefix, uri)

    old_sz = str(old_size * 100)
    new_sz = str(new_size * 100)

    for rPr in root.iter('{%s}rPr' % NS['a']):
        if rPr.get('sz') == old_sz:
            rPr.set('sz', new_sz)

    tree.write(slide_path, xml_declaration=True, encoding='UTF-8')

for slide in Path("unpacked/ppt/slides").glob("slide*.xml"):
    change_font_size(slide, old_size=44, new_size=22)
```

### FORBIDDEN Practices

- ❌ NEVER use sed/awk/string replacement on XML
- ❌ NEVER edit without inventory first
- ❌ NEVER pack without validation
- ❌ NEVER guess at element IDs

---

## Validation

**ALWAYS validate after XML edits:**

```bash
venv/bin/python .claude/skills/pptx/ooxml/scripts/validate.py unpacked/ --original presentation.pptx
```

**Only pack after validation passes:**

```bash
venv/bin/python .claude/skills/pptx/ooxml/scripts/pack.py unpacked/ output.pptx
```

---

## Working with Messy Decks

Signs of a messy deck:
- `placeholder_type: null` for most shapes
- 5+ random text boxes at odd positions
- Inconsistent sizing and formatting

### Strategy

1. **Thumbnail first** - See what you're working with
2. **Inventory** - Understand actual positions
3. **Document visually** - Create your own layout inventory
4. **Pick best-fit slides** - Choose by visual structure
5. **Accept imperfection** - Focus on content correctness
6. **Verify output** - Thumbnail the result

### When to Give Up

Use bullets when:
- Slides have 6+ overlapping boxes
- Positions make no sense
- Analysis takes longer than creating fresh

---

## Output Directory Convention

All files go to:
```
outputs/<document-name>/
```

Example:
```bash
mkdir -p outputs/q4-update/
venv/bin/python .claude/skills/pptx/scripts/inventory.py deck.pptx outputs/q4-update/inventory.json
venv/bin/python .claude/skills/pptx/scripts/replace.py deck.pptx outputs/q4-update/replacements.json outputs/q4-update/output.pptx
```

---

## Dependencies

- **markitdown**: Text extraction
- **pptxgenjs**: HTML to PPTX conversion
- **playwright**: HTML rendering
- **sharp**: Image processing
- **LibreOffice**: PDF conversion
- **Poppler**: PDF to images
- **defusedxml**: Secure XML parsing
