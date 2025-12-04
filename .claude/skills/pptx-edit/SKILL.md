---
name: pptx-edit
description: "Edit existing PowerPoint presentations. Use when modifying an existing .pptx file: (1) Updating text content, (2) Applying a new outline to an existing deck/template, (3) Batch style changes (fonts, sizes, colors), (4) Find and replace operations, (5) Standardizing formatting. Do NOT use for creating new presentations from scratch - use the pptx skill instead."
---

# PowerPoint Editing Skill

Edit existing PowerPoint presentations. For creating new presentations from scratch, use the `pptx` skill instead.

## When to Use This Skill

Use this skill when:
- Updating text content in an existing deck
- Applying a new content outline to an existing deck/template
- Changing fonts, colors, or sizes throughout a presentation
- Standardizing formatting across slides
- Batch find-and-replace operations
- Fixing inconsistent styling
- Updating brand elements (colors, fonts)
- Repurposing a previous project's deck with new content

Do NOT use this skill when:
- Creating a new presentation from scratch (no existing file) → use `pptx` skill
- Converting HTML to PowerPoint → use `pptx` skill
- Need complex visual designs not in your existing deck → use `pptx` skill

---

## Quick Reference: Common Tasks

| Task | Command |
|------|---------|
| See all text | `venv/bin/python -m markitdown file.pptx` |
| Extract text inventory | `venv/bin/python scripts/inventory.py file.pptx inventory.json` |
| Replace text | `venv/bin/python scripts/replace.py file.pptx replacements.json output.pptx` |
| Visual preview | `venv/bin/python scripts/thumbnail.py file.pptx thumbnails` |

**Note**: Scripts are located at `.claude/skills/pptx/scripts/` relative to the project root.

---

## Workflow 1: Text Replacement (Most Common)

**Use for:** Updating text content while preserving all formatting.

### Step 1: Extract Text Inventory

```bash
venv/bin/python .claude/skills/pptx/scripts/inventory.py presentation.pptx inventory.json
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
venv/bin/python .claude/skills/pptx/scripts/replace.py presentation.pptx replacements.json output.pptx
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
venv/bin/python .claude/skills/pptx/scripts/inventory.py presentation.pptx inventory.json
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
venv/bin/python .claude/skills/pptx/ooxml/scripts/unpack.py presentation.pptx unpacked/
```

2. Write a Python script to modify the XML (see "XML Editing Patterns" below)

3. Validate:
```bash
venv/bin/python .claude/skills/pptx/ooxml/scripts/validate.py unpacked/ --original presentation.pptx
```

4. Pack only if validation passes:
```bash
venv/bin/python .claude/skills/pptx/ooxml/scripts/pack.py unpacked/ output.pptx
```

---

## Workflow 3: Apply New Content Outline to Existing Deck

**Use for:** You have an existing deck (template or previous project) + a new content outline (markdown/text) and want to populate the deck with the new content.

This is the most complex editing workflow. Follow each step carefully.

### Step 1: Analyze the Existing Deck

Create a visual inventory of what slide layouts exist:

```bash
# Extract text to understand content structure
venv/bin/python -m markitdown existing.pptx > outputs/project/deck-text.md

# Create thumbnail grid for visual reference
venv/bin/python .claude/skills/pptx/scripts/thumbnail.py existing.pptx outputs/project/thumbnails

# Extract detailed inventory
venv/bin/python .claude/skills/pptx/scripts/inventory.py existing.pptx outputs/project/inventory.json
```

### Step 2: Create a Slide Layout Inventory

**MANDATORY:** Before mapping content, document ALL available slide layouts.

Create `outputs/project/layout-inventory.md`:

```markdown
# Slide Layout Inventory

| Index | Layout Type | Placeholders | Best For |
|-------|-------------|--------------|----------|
| 0 | Title slide | title, subtitle | Opening |
| 1 | Section header | title only | Section breaks |
| 2 | Title + bullets | title, body (5 bullets) | Lists, key points |
| 3 | Two-column | title, left, right | Comparisons (2 items) |
| 4 | Three-column | title, col1, col2, col3 | Comparisons (3 items) |
| 5 | Image + text | title, text, image placeholder | Visual + explanation |
| ...
```

### Step 3: Map Your Outline to Layouts

**CRITICAL LAYOUT MATCHING RULES:**

| Your Content | Correct Layout | WRONG Choice |
|--------------|----------------|--------------|
| 2 comparison items | Two-column (index 3) | Three-column (empty column!) |
| 3 pillars/items | Three-column (index 4) | Two-column (overflow!) |
| 5+ bullet points | Single-column bullets | Force into columns |
| No image available | Text-only layout | Image layout (empty!) |
| Quote with speaker | Quote layout | Content (loses impact) |

Create your mapping plan:

```markdown
# Content-to-Slide Mapping

## My Outline:
1. Title: "Q4 Strategy Review"
2. Agenda (4 items)
3. Market Overview (3 key trends)
4. Our Response (2 strategic pillars)
5. Next Steps (5 action items)
6. Summary

## Mapping:
| Outline Section | Slide Index | Layout Used | Why |
|-----------------|-------------|-------------|-----|
| Title | 0 | Title slide | Standard opening |
| Agenda | 2 | Bullets | 4 items = bullet list |
| Market Overview | 4 | Three-column | Exactly 3 trends |
| Our Response | 3 | Two-column | Exactly 2 pillars |
| Next Steps | 2 | Bullets | 5 items = bullet list |
| Summary | 2 | Bullets | Key takeaways |
```

### Step 4: Rearrange Slides to Match Outline

Use `rearrange.py` to build your slide sequence:

```bash
# Format: indices of slides to include (0-based), can repeat
venv/bin/python .claude/skills/pptx/scripts/rearrange.py \
  existing.pptx \
  outputs/project/working.pptx \
  0,2,4,3,2,2
```

The numbers correspond to:
- `0` = Title slide
- `2` = Bullets layout (for Agenda)
- `4` = Three-column (for Market Overview)
- `3` = Two-column (for Our Response)
- `2` = Bullets layout (for Next Steps)
- `2` = Bullets layout (for Summary)

### Step 5: Get Fresh Inventory of Working Deck

```bash
venv/bin/python .claude/skills/pptx/scripts/inventory.py \
  outputs/project/working.pptx \
  outputs/project/working-inventory.json
```

**MANDATORY:** Read the entire `working-inventory.json`. The shape IDs may have changed after rearrangement.

### Step 6: Create Replacement JSON from Your Outline

Map your content to the inventory structure:

```json
{
  "slide-0": {
    "shape-0": {
      "paragraphs": [{"text": "Q4 Strategy Review", "bold": true, "alignment": "CENTER"}]
    },
    "shape-1": {
      "paragraphs": [{"text": "Board Presentation | January 2025", "alignment": "CENTER"}]
    }
  },
  "slide-1": {
    "shape-0": {
      "paragraphs": [{"text": "Agenda", "bold": true}]
    },
    "shape-1": {
      "paragraphs": [
        {"text": "Market Overview", "bullet": true, "level": 0},
        {"text": "Our Strategic Response", "bullet": true, "level": 0},
        {"text": "Action Items", "bullet": true, "level": 0},
        {"text": "Summary & Discussion", "bullet": true, "level": 0}
      ]
    }
  },
  "slide-2": {
    "shape-0": {
      "paragraphs": [{"text": "Market Overview", "bold": true}]
    },
    "shape-1": {
      "paragraphs": [{"text": "Trend 1: Digital Transformation", "bold": true}]
    },
    "shape-2": {
      "paragraphs": [{"text": "Trend 2: Sustainability Focus", "bold": true}]
    },
    "shape-3": {
      "paragraphs": [{"text": "Trend 3: AI Integration", "bold": true}]
    }
  }
}
```

### Step 7: Apply Replacements

```bash
venv/bin/python .claude/skills/pptx/scripts/replace.py \
  outputs/project/working.pptx \
  outputs/project/replacements.json \
  outputs/project/final.pptx
```

### Step 8: Visual Verification

```bash
venv/bin/python .claude/skills/pptx/scripts/thumbnail.py \
  outputs/project/final.pptx \
  outputs/project/final-thumbnails
```

Review thumbnails for:
- ✓ All slides have content (no empty placeholders)
- ✓ Text fits without overflow
- ✓ Layout matches content quantity
- ✓ Visual balance looks correct

### Common Mistakes to Avoid

| Mistake | Result | Fix |
|---------|--------|-----|
| Using 3-column for 2 items | Empty column visible | Use 2-column layout |
| Forgetting to re-inventory after rearrange | Wrong shape IDs | Always re-run inventory.py |
| Not including all shapes in JSON | Shapes get cleared | Include every shape you want to keep |
| Adding bullet symbol in text | Double bullets | Use `"bullet": true`, no symbol in text |

---

## Workflow 4: Find and Replace Text

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
venv/bin/python .claude/skills/pptx/ooxml/scripts/validate.py unpacked/ --original presentation.pptx
```

Validation checks:
- XML schema compliance
- Required elements present
- Relationships intact
- No orphaned references

**Only pack after validation passes:**

```bash
venv/bin/python .claude/skills/pptx/ooxml/scripts/pack.py unpacked/ output.pptx
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

## Working with Messy Decks

Most real-world PowerPoint files aren't clean templates. They're messy — slides copied from multiple sources, manual text boxes instead of placeholders, inconsistent formatting. This section covers how to work with them.

### Recognizing a Messy Deck

Signs you have a messy deck:
- `placeholder_type: null` for most shapes in inventory
- Slides with 5+ random text boxes at odd positions
- Inconsistent sizing (similar-looking boxes have different widths)
- No clear master/layout structure
- Mix of fonts, sizes, colors with no pattern

### The Inventory Tells the Truth

Even in messy decks, `inventory.py` reveals the actual structure:

```json
{
  "slide-3": {
    "shape-0": {
      "left": 0.5,
      "top": 0.3,
      "width": 12.3,
      "placeholder_type": null,  // Not a real placeholder - manual text box
      "paragraphs": [{"text": "Some Heading"}]
    },
    "shape-1": {
      "left": 0.5,
      "top": 1.5,
      "width": 5.8,             // Left half of slide
      "paragraphs": [{"text": "Left content"}]
    },
    "shape-2": {
      "left": 6.8,
      "top": 1.5,
      "width": 5.8,             // Right half of slide
      "paragraphs": [{"text": "Right content"}]
    }
  }
}
```

**Key insight:** `placeholder_type: null` means manually placed shape. You can still use it — just know it wasn't designed as a template element.

### Visual Analysis is Critical

When there's no structure to rely on, thumbnails are your guide:

```bash
venv/bin/python .claude/skills/pptx/scripts/thumbnail.py messy.pptx outputs/analysis/thumbs
```

Then manually document what you see:

```markdown
# Visual Slide Analysis (messy deck)

| Index | What I See | Usable For |
|-------|------------|------------|
| 0 | Big text center, small text below | Title slide |
| 1 | Just a big heading | Section break |
| 2 | Heading + one text box below | Single content |
| 3 | Heading + two boxes side by side | Two-column-ish |
| 4 | Heading + THREE boxes (uneven sizes) | Three-column (imperfect) |
| 5 | Heading + giant empty space + small text | Was image+text, image deleted |
| 6 | Total chaos - 8 random text boxes | AVOID - too messy |
```

### Identify Usable Slides by Position

Use the inventory positions to identify layout structure:

```python
# Pseudo-logic for identifying layout from inventory
def identify_layout(shapes):
    # Filter to content shapes (exclude tiny shapes, likely decorative)
    content_shapes = [s for s in shapes if s['width'] > 2.0 and s['height'] > 0.5]

    # Check for side-by-side arrangement
    lefts = [s['left'] for s in content_shapes]
    if len(content_shapes) == 3:
        # One at top (title), two side by side below
        return "two-column-ish"
    elif len(content_shapes) == 4:
        # Title + three columns
        return "three-column-ish"
    else:
        return "single-column or chaos"
```

### Fallback Strategies

When the deck doesn't have what you need:

| You Need | Deck Has | Solution |
|----------|----------|----------|
| 3-column | Only 2-column | Use 2-col for 2 items, bullets for 3+ |
| Clean bullet layout | Only messy text boxes | Find the cleanest single-box slide |
| Image+text | No image slides | Use text-only, mention image separately |
| Quote layout | Nothing suitable | Use any centered text slide |
| Consistent styling | Total chaos | Pick ONE slide style, use it repeatedly |

### The "Good Enough" Principle

With messy decks, aim for:
- ✓ Content is correct and readable
- ✓ No obviously broken layouts
- ✓ Consistent within YOUR new content
- ✗ Won't perfectly match original deck's chaos
- ✗ May need manual cleanup afterward

### When to Give Up

Sometimes a deck is too messy. Consider giving up and using bullets when:
- Slides have 6+ overlapping text boxes
- Positions make no logical sense
- You'd spend more time analyzing than creating fresh
- The "cleanest" slide is still chaotic

**Fallback:** Use the simplest slide with fewest shapes, put everything in bullets.

### Messy Deck Workflow Summary

1. **Thumbnail first** — See what you're working with
2. **Inventory** — Understand actual shape positions
3. **Document visually** — Create your own layout inventory based on what you SEE
4. **Pick best-fit slides** — Choose by visual structure, not by name
5. **Accept imperfection** — Focus on content correctness over visual perfection
6. **Verify output** — Thumbnail the result to catch issues

---

## Output Directory Convention

All edited files go to:
```
outputs/<document-name>/
```

Example workflow:
```bash
mkdir -p outputs/q4-update/
venv/bin/python .claude/skills/pptx/scripts/inventory.py deck.pptx outputs/q4-update/inventory.json
# ... create replacements.json ...
venv/bin/python .claude/skills/pptx/scripts/replace.py deck.pptx outputs/q4-update/replacements.json outputs/q4-update/deck-updated.pptx
```
