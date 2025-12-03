# PowerPoint Skill Capability Tests

This test suite evaluates the capabilities and limitations of the Claude Code PPTX skill for real-world use cases.

## Purpose

When using Claude Code to create or modify PowerPoint presentations, users need to understand:
1. **How instruction variations affect quality** - Which SKILL.md instructions lead to better outputs
2. **What works reliably** - Operations that can be trusted
3. **What has limitations** - Operations that work with caveats
4. **What doesn't work** - Operations that will fail or produce poor results
5. **Complexity tolerance** - How messy/complex a template can be before the skill breaks down

## PRIMARY: Skill Instruction Variations (`test_skill_variations.py`)

**This is the most important test file.** It evaluates how different SKILL.md instruction configurations affect Claude Code's behavior when editing PowerPoint templates.

### Skill Variations Tested

| Variation | Description | Token Cost | Expected Quality |
|-----------|-------------|------------|------------------|
| `baseline` | Minimal instructions | Low | May miss edge cases |
| `enhanced` | Added layout matching rules | Medium | Better layout choices |
| `strict` | Forbidden patterns + decision tree | Higher | Highest constraint satisfaction |
| `examples` | Correct/incorrect examples | Medium | Easy pattern matching |
| `checklist` | Pre-flight verification | Medium | Catches errors before applying |

### Key Scenarios Tested

| ID | Scenario | Correct Choice | Common Error |
|----|----------|----------------|--------------|
| S01 | 2 items to compare | 2-column | 3-column (empty column) |
| S02 | 3 strategic pillars | 3-column | 2-column (overflow) |
| S03 | 5 bullet points | bullet list | cramming into columns |
| S04 | No image available | text-only | image layout (empty area) |
| S05 | Emphasis text | content layout | quote layout (no attribution) |
| S06 | Actual quote with speaker | quote layout | content (loses impact) |
| S07 | 4 awkward items | list or 2 slides | force into 3-column |

### Recommended Instruction Additions

Based on testing, these instruction additions to SKILL.md improve output quality:

```markdown
### CRITICAL Layout Matching Rules

FORBIDDEN PATTERNS (NEVER DO):
- NEVER use 3-column layout for 2 items (leaves empty column)
- NEVER use image placeholder without actual image
- NEVER use quote layout for non-quotes (requires attribution)
- NEVER leave any placeholder empty

DECISION TREE:
1. Count content items
2. Find layout with EXACT placeholder match
3. If no exact match: use bullet list format
4. If 5+ items: split across multiple slides
```

---

## Test Categories

### 1. Text Operations (`test_text_ops.py`)
Tests text extraction, replacement, and formatting preservation.

| Test ID | Description | Complexity | Expected |
|---------|-------------|------------|----------|
| T01 | Simple text replacement | Low | Pass |
| T02 | Multi-paragraph replacement | Low | Pass |
| T03 | Bullet list formatting | Medium | Pass |
| T04 | Numbered list formatting | Medium | Pass |
| T05 | Mixed formatting (bold, italic, color) | Medium | Pass |
| T06 | Nested bullet levels (3+ levels) | High | Pass with caveats |
| T07 | Text overflow detection | Medium | Pass |
| T08 | Unicode/special characters | Medium | Pass |
| T09 | Right-to-left text | High | May fail |
| T10 | Text in grouped shapes | High | Pass with caveats |

### 2. Template Complexity (`test_template_complexity.py`)
Tests how well the skill handles templates of varying complexity.

| Test ID | Description | Complexity | Expected |
|---------|-------------|------------|----------|
| TC01 | Clean corporate template (5 slides) | Low | Pass |
| TC02 | Medium template (20 slides, standard layouts) | Medium | Pass |
| TC03 | Complex template (50+ slides, mixed layouts) | High | Pass with caveats |
| TC04 | Template with hidden slides | Medium | Pass |
| TC05 | Template with master slide variations | High | Pass |
| TC06 | Template with grouped shapes | High | Pass with caveats |
| TC07 | Template with nested groups (3+ levels) | Very High | May fail |
| TC08 | Template with SmartArt | Very High | Likely fail |
| TC09 | Template with embedded charts | High | Partial support |
| TC10 | Template with media (audio/video) | High | Not supported |

### 3. Slide Operations (`test_slide_ops.py`)
Tests slide manipulation capabilities.

| Test ID | Description | Complexity | Expected |
|---------|-------------|------------|----------|
| SO01 | Duplicate single slide | Low | Pass |
| SO02 | Duplicate multiple slides | Low | Pass |
| SO03 | Reorder slides | Low | Pass |
| SO04 | Delete slides | Low | Pass |
| SO05 | Complex rearrangement (mix of all) | Medium | Pass |
| SO06 | Duplicate slide with images | Medium | Pass |
| SO07 | Duplicate slide with charts | High | Pass with caveats |
| SO08 | Preserve notes on duplication | Medium | May have issues |
| SO09 | Handle 100+ slide deck | High | Performance test |
| SO10 | Handle corrupt slide indices | Low | Should error gracefully |

### 4. HTML-to-PPTX Creation (`test_html_creation.py`)
Tests creating presentations from scratch using HTML.

| Test ID | Description | Complexity | Expected |
|---------|-------------|------------|----------|
| HC01 | Simple text slide | Low | Pass |
| HC02 | Two-column layout | Medium | Pass |
| HC03 | Image placement | Medium | Pass |
| HC04 | Shape with background color | Low | Pass |
| HC05 | Shape with border | Medium | Pass |
| HC06 | Rounded rectangle | Medium | Pass |
| HC07 | Drop shadow | Medium | Pass |
| HC08 | Gradient background | High | Requires pre-rasterization |
| HC09 | Complex flexbox layout | High | Pass |
| HC10 | Chart placeholder | Medium | Pass |
| HC11 | Table placeholder | Medium | Pass |
| HC12 | Mixed content slide | High | Pass |
| HC13 | Icon integration | High | Requires pre-rasterization |
| HC14 | Vertical text rotation | Medium | Pass |
| HC15 | Custom fonts | High | Limited to web-safe |

### 5. Visual Quality (`test_visual_quality.py`)
Tests output quality and visual accuracy.

| Test ID | Description | Complexity | Expected |
|---------|-------------|------------|----------|
| VQ01 | Text alignment accuracy | Medium | Pass |
| VQ02 | Font size preservation | Medium | Pass |
| VQ03 | Color accuracy (RGB) | Low | Pass |
| VQ04 | Theme color preservation | Medium | Pass |
| VQ05 | Shape positioning accuracy | Medium | Pass |
| VQ06 | Image aspect ratio | Medium | Pass |
| VQ07 | Line spacing accuracy | High | Pass with variance |
| VQ08 | Bullet indentation accuracy | Medium | Pass |
| VQ09 | Border thickness accuracy | Medium | Pass |
| VQ10 | Shadow rendering | High | Approximate |

### 6. Edge Cases and Limitations (`test_edge_cases.py`)
Tests known limitations and edge cases.

| Test ID | Description | Complexity | Expected |
|---------|-------------|------------|----------|
| EC01 | Empty shape handling | Low | Pass (cleared) |
| EC02 | Shape with only whitespace | Low | Pass |
| EC03 | Very long text (1000+ chars) | Medium | Pass with overflow |
| EC04 | Zero-width shape | Low | Should error |
| EC05 | Overlapping shapes | Medium | Detected |
| EC06 | Shape outside slide bounds | Medium | Detected |
| EC07 | Missing font fallback | Medium | Falls back to default |
| EC08 | Corrupt XML handling | High | Should error gracefully |
| EC09 | Password-protected file | High | Not supported |
| EC10 | Macro-enabled (.pptm) | High | Not tested |

### 7. Real-World Scenarios (`test_real_world.py`)
Tests based on common real-world use cases.

| Test ID | Description | Complexity | Expected |
|---------|-------------|------------|----------|
| RW01 | Update quarterly report template | Medium | Pass |
| RW02 | Create pitch deck from scratch | Medium | Pass |
| RW03 | Localize presentation (text swap) | Medium | Pass |
| RW04 | Create multiple versions from template | Medium | Pass |
| RW05 | Extract and report all text content | Low | Pass |
| RW06 | Batch update company branding | High | Partial |
| RW07 | Convert outline to presentation | Medium | Pass |
| RW08 | Create presentation with web images | High | Not directly supported |
| RW09 | Add charts from data | Medium | Via placeholder |
| RW10 | Professional design from scratch | High | Pass with skill |

### 8. Performance Tests (`test_performance.py`)
Tests for performance with large files.

| Test ID | Description | Metric | Threshold |
|---------|-------------|--------|-----------|
| P01 | Inventory extraction (10 slides) | Time | < 5s |
| P02 | Inventory extraction (50 slides) | Time | < 15s |
| P03 | Inventory extraction (100 slides) | Time | < 30s |
| P04 | Replace operation (10 slides) | Time | < 10s |
| P05 | Thumbnail grid (50 slides) | Time | < 60s |
| P06 | HTML conversion (10 slides) | Time | < 30s |
| P07 | OOXML validation | Time | < 20s |
| P08 | Memory usage (large deck) | Memory | < 500MB |

## Known Limitations Summary

### Completely Unsupported
- CSS gradients (must pre-rasterize to PNG)
- Audio/video embedding
- Animations and transitions
- SmartArt editing
- 3D effects
- Custom fonts (web-safe only)
- Password-protected files
- Direct web image search/insertion

### Partially Supported
- Charts (via placeholder only, limited types)
- Tables (basic formatting only)
- Nested grouped shapes (may have position issues)
- Complex bullet formatting (may lose some properties)
- Theme colors (preserved but not editable)

### Works with Caveats
- Text overflow detection (approximate)
- Font measurement (may vary from PowerPoint)
- Shadow rendering (outer only, approximate)
- Complex layouts (requires careful HTML construction)

## Running Tests

```bash
# Run all tests
venv/bin/python -m pytest tests/pptx/ -v

# Run specific category
venv/bin/python -m pytest tests/pptx/test_text_ops.py -v

# Run with coverage
venv/bin/python -m pytest tests/pptx/ --cov=public/pptx/scripts

# Generate HTML report
venv/bin/python -m pytest tests/pptx/ --html=tests/pptx/results/report.html
```

## Test Fixtures

Test fixtures are located in `tests/pptx/fixtures/`:
- `simple_template.pptx` - 5 slide basic template
- `medium_template.pptx` - 20 slide corporate template
- `complex_template.pptx` - 50+ slide template with various elements
- `edge_cases.pptx` - File with intentional edge cases

## Contributing New Tests

When adding new tests:
1. Follow the naming convention: `test_<category>_<specific>.py`
2. Document the test in this file
3. Create necessary fixtures
4. Note expected results and any known issues
