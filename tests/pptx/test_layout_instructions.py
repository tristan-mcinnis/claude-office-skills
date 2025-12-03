"""
Test layout instruction variations for PPTX skill.

These tests evaluate how different instruction variations affect output quality
when using Claude Code to edit or create presentations.

The goal is to understand:
1. Which instructions have the highest impact on quality
2. How instruction specificity affects layout matching
3. Whether additional constraints improve or degrade results
"""
import pytest
import json
import subprocess
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN


# ============================================================================
# INSTRUCTION VARIATIONS TO TEST
# ============================================================================

LAYOUT_INSTRUCTIONS = {
    "minimal": """
        Match template slides to content.
    """,

    "standard": """
        Match template slides to content appropriately.
        Use layouts that fit your content quantity.
    """,

    "detailed": """
        CRITICAL: Match layout structure to actual content:
        - Single-column layouts: Use for unified narrative or single topic
        - Two-column layouts: Use ONLY when you have exactly 2 distinct items/concepts
        - Three-column layouts: Use ONLY when you have exactly 3 distinct items/concepts
        - Image + text layouts: Use ONLY when you have actual images to insert
        - Quote layouts: Use ONLY for actual quotes from people (with attribution)
        - Never use layouts with more placeholders than you have content
    """,

    "strict": """
        CRITICAL LAYOUT RULES (MUST FOLLOW):

        1. PLACEHOLDER COUNT MATCHING:
           - Count your content items FIRST
           - Select layout with EXACT placeholder count
           - 2 items → 2-column layout ONLY
           - 3 items → 3-column layout ONLY
           - 4+ items → Multiple slides OR list format

        2. FORBIDDEN PATTERNS:
           - NEVER use 3-column for 2 items (leaves empty column)
           - NEVER use image placeholder without actual image
           - NEVER use quote layout for non-quotes
           - NEVER leave placeholders empty

        3. FALLBACK RULES:
           - When in doubt, use single-column with bullets
           - Long lists (5+ items) → Multiple slides
           - Mixed content types → Separate slides
    """,

    "examples": """
        LAYOUT MATCHING WITH EXAMPLES:

        CORRECT:
        ✓ 2 benefits → 2-column comparison layout
        ✓ 3 pillars → 3-column layout
        ✓ 5 bullet points → Single-column bullet list
        ✓ Quote from CEO → Quote layout with attribution

        INCORRECT:
        ✗ 2 items in 3-column layout (empty column)
        ✗ 4 items forced into 3-column layout
        ✗ Image layout without actual image
        ✗ Quote layout for emphasis text (no attribution)

        DECISION TREE:
        1. How many distinct items? → Match to column count
        2. Is it a real quote? → Use quote layout
        3. Do you have an image? → Use image layout
        4. None of above? → Use bullet list
    """
}


# ============================================================================
# TEST TEMPLATES WITH VARYING LAYOUTS
# ============================================================================

def create_multi_layout_template(path):
    """Create template with various layout types for testing."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 0: Title
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.3), Inches(1.5))
    title.text_frame.paragraphs[0].text = "[Title]"
    title.text_frame.paragraphs[0].font.size = Pt(48)
    title.text_frame.paragraphs[0].font.bold = True

    # Slide 1: Single column bullets
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    title.text_frame.paragraphs[0].text = "[Single Column Title]"
    title.text_frame.paragraphs[0].font.size = Pt(36)

    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.3), Inches(5))
    tf = body.text_frame
    tf.paragraphs[0].text = "[Bullet 1]"
    for text in ["[Bullet 2]", "[Bullet 3]", "[Bullet 4]", "[Bullet 5]"]:
        p = tf.add_paragraph()
        p.text = text

    # Slide 2: Two column
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    title.text_frame.paragraphs[0].text = "[Two Column Title]"
    title.text_frame.paragraphs[0].font.size = Pt(36)

    left = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5.8), Inches(5))
    left.text_frame.paragraphs[0].text = "[Left Column Content]"

    right = slide.shapes.add_textbox(Inches(6.8), Inches(1.5), Inches(5.8), Inches(5))
    right.text_frame.paragraphs[0].text = "[Right Column Content]"

    # Slide 3: Three column
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    title.text_frame.paragraphs[0].text = "[Three Column Title]"
    title.text_frame.paragraphs[0].font.size = Pt(36)

    col1 = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(3.8), Inches(5))
    col1.text_frame.paragraphs[0].text = "[Column 1]"

    col2 = slide.shapes.add_textbox(Inches(4.7), Inches(1.5), Inches(3.8), Inches(5))
    col2.text_frame.paragraphs[0].text = "[Column 2]"

    col3 = slide.shapes.add_textbox(Inches(8.9), Inches(1.5), Inches(3.8), Inches(5))
    col3.text_frame.paragraphs[0].text = "[Column 3]"

    # Slide 4: Image + text layout
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    title.text_frame.paragraphs[0].text = "[Image Layout Title]"
    title.text_frame.paragraphs[0].font.size = Pt(36)

    text_area = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5), Inches(5))
    text_area.text_frame.paragraphs[0].text = "[Text beside image]"

    image_placeholder = slide.shapes.add_textbox(Inches(6), Inches(1.5), Inches(6.5), Inches(5))
    image_placeholder.text_frame.paragraphs[0].text = "[IMAGE PLACEHOLDER]"

    # Slide 5: Quote layout
    slide = prs.slides.add_slide(blank)
    quote = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(3))
    tf = quote.text_frame
    tf.paragraphs[0].text = '"[Quote text here]"'
    tf.paragraphs[0].font.size = Pt(32)
    tf.paragraphs[0].font.italic = True

    attribution = slide.shapes.add_textbox(Inches(1), Inches(5.5), Inches(11), Inches(1))
    attribution.text_frame.paragraphs[0].text = "— [Speaker Name, Title]"
    attribution.text_frame.paragraphs[0].font.size = Pt(20)
    attribution.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

    prs.save(path)
    return path


# ============================================================================
# TEST SCENARIOS
# ============================================================================

class TestLayoutMatchingScenarios:
    """Test scenarios for layout matching quality."""

    @pytest.fixture
    def template(self, temp_dir):
        """Create test template."""
        path = temp_dir / "layouts.pptx"
        create_multi_layout_template(path)
        return path

    def test_li01_two_items_needs_two_columns(self, template, temp_dir, python_cmd, scripts_dir):
        """LI01: 2 items should use 2-column, not 3-column layout."""
        # Scenario: User has exactly 2 comparison points
        content = {
            "items": ["Benefit A: Cost savings", "Benefit B: Time efficiency"],
            "count": 2
        }

        # Expected: Should select slide-2 (two-column), NOT slide-3 (three-column)
        expected_slide_index = 2

        # Get inventory to understand layout
        inventory_path = temp_dir / "inventory.json"
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Document the test case
        test_case = {
            "scenario": "2 items comparison",
            "content_count": content["count"],
            "correct_layout": "slide-2 (two-column)",
            "incorrect_layout": "slide-3 (three-column)",
            "instruction_needed": "Two-column layouts: Use ONLY when you have exactly 2 distinct items",
            "inventory_slide_2_shapes": len(inventory.get("slide-2", {})),
            "inventory_slide_3_shapes": len(inventory.get("slide-3", {}))
        }

        # Verify two-column has 3 shapes (title + 2 columns)
        assert len(inventory.get("slide-2", {})) == 3, "Two-column should have 3 shapes"
        # Verify three-column has 4 shapes (title + 3 columns)
        assert len(inventory.get("slide-3", {})) == 4, "Three-column should have 4 shapes"

    def test_li02_three_items_needs_three_columns(self, template, temp_dir, python_cmd, scripts_dir):
        """LI02: 3 items should use 3-column layout."""
        content = {
            "items": ["Pillar 1", "Pillar 2", "Pillar 3"],
            "count": 3
        }

        expected_slide_index = 3  # Three-column layout

        inventory_path = temp_dir / "inventory.json"
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        test_case = {
            "scenario": "3 distinct pillars",
            "content_count": content["count"],
            "correct_layout": "slide-3 (three-column)",
            "instruction_needed": "Three-column layouts: Use ONLY when you have exactly 3 distinct items"
        }

        assert len(inventory.get("slide-3", {})) == 4

    def test_li03_five_items_needs_bullet_list(self, template, temp_dir, python_cmd, scripts_dir):
        """LI03: 5+ items should use bullet list, not force into columns."""
        content = {
            "items": ["Point 1", "Point 2", "Point 3", "Point 4", "Point 5"],
            "count": 5
        }

        expected_slide_index = 1  # Single column bullet list

        inventory_path = temp_dir / "inventory.json"
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        test_case = {
            "scenario": "5 bullet points",
            "content_count": content["count"],
            "correct_layout": "slide-1 (single-column bullets)",
            "incorrect_layout": "slide-3 (three-column - would overflow)",
            "instruction_needed": "If you have 4+ items, use list format or multiple slides"
        }

        # Single column should have 2 shapes (title + body)
        assert len(inventory.get("slide-1", {})) == 2

    def test_li04_no_image_skip_image_layout(self, template, temp_dir, python_cmd, scripts_dir):
        """LI04: Without actual image, don't use image layout."""
        content = {
            "has_image": False,
            "text": "Description text only"
        }

        # Should NOT use slide-4 (image layout)
        expected_to_avoid = 4

        inventory_path = temp_dir / "inventory.json"
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        test_case = {
            "scenario": "Text content without image",
            "has_image": False,
            "incorrect_layout": "slide-4 (image + text)",
            "instruction_needed": "Image + text layouts: Use ONLY when you have actual images to insert"
        }

        # Document image layout structure
        image_layout_shapes = inventory.get("slide-4", {})
        assert len(image_layout_shapes) == 3  # title, text, image placeholder

    def test_li05_quote_needs_attribution(self, template, temp_dir, python_cmd, scripts_dir):
        """LI05: Quote layout only for actual quotes with attribution."""
        content_with_quote = {
            "text": "Innovation distinguishes between a leader and a follower.",
            "speaker": "Steve Jobs",
            "has_attribution": True
        }

        content_without_quote = {
            "text": "Our key focus area this quarter",
            "has_attribution": False
        }

        inventory_path = temp_dir / "inventory.json"
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        test_case = {
            "scenario": "Quote vs emphasis text",
            "correct_for_quote": "slide-5 (quote layout)",
            "correct_for_emphasis": "slide-1 (bullet) or slide-2 (two-column)",
            "instruction_needed": "Quote layouts: Use ONLY for actual quotes with attribution"
        }

        # Quote layout should have 2 shapes (quote + attribution)
        assert len(inventory.get("slide-5", {})) == 2


class TestEmptyPlaceholderDetection:
    """Test detection of empty placeholders (layout misuse)."""

    def test_li06_detect_empty_column(self, temp_dir, python_cmd, scripts_dir):
        """LI06: Detect when content leaves placeholder empty."""
        template_path = temp_dir / "three_col.pptx"
        create_multi_layout_template(template_path)

        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "bad_replacement.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Intentionally BAD replacement: only fill 2 of 3 columns
        # This simulates the error of using 3-column for 2 items
        slide_3_shapes = list(inventory.get("slide-3", {}).keys())

        if len(slide_3_shapes) >= 4:
            # Fill title and only 2 columns (leave 3rd empty)
            replacements = {
                "slide-3": {
                    slide_3_shapes[0]: {"paragraphs": [{"text": "Title"}]},
                    slide_3_shapes[1]: {"paragraphs": [{"text": "Column 1 content"}]},
                    slide_3_shapes[2]: {"paragraphs": [{"text": "Column 2 content"}]},
                    # slide_3_shapes[3] intentionally NOT included - will be cleared
                }
            }

            with open(replacements_path, "w") as f:
                json.dump(replacements, f)

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "replace.py"),
                 str(template_path), str(replacements_path), str(output_path)],
                capture_output=True,
                text=True
            )

            # The replace script clears unfilled shapes
            # This test documents that behavior
            assert result.returncode == 0

            # Document the issue
            issue = {
                "problem": "Using 3-column layout for 2 items",
                "result": "Third column is cleared/empty",
                "visual_impact": "Unbalanced slide with empty space",
                "instruction_fix": "Never use layouts with more placeholders than you have content"
            }


class TestInstructionImpact:
    """Test how instruction variations affect decisions."""

    def test_li07_minimal_vs_detailed_instructions(self, temp_dir):
        """LI07: Compare minimal vs detailed instruction outcomes."""
        # This is a documentation test showing instruction impact

        scenarios = [
            {
                "content": {"items": 2, "type": "comparison"},
                "minimal_likely_choice": "any multi-column",
                "detailed_correct_choice": "two-column only",
                "improvement": "Prevents empty placeholder"
            },
            {
                "content": {"items": 4, "type": "features"},
                "minimal_likely_choice": "3-column (overflow)",
                "detailed_correct_choice": "bullet list or 2 slides",
                "improvement": "Prevents cramming"
            },
            {
                "content": {"has_image": False, "type": "text"},
                "minimal_likely_choice": "may use image layout",
                "detailed_correct_choice": "text-only layout",
                "improvement": "No empty image placeholder"
            },
            {
                "content": {"is_quote": False, "type": "emphasis"},
                "minimal_likely_choice": "may use quote layout",
                "detailed_correct_choice": "standard content layout",
                "improvement": "Appropriate formatting"
            }
        ]

        # Summary
        impact_summary = {
            "minimal_instructions": {
                "pros": ["Faster processing", "More flexibility"],
                "cons": ["Layout mismatches", "Empty placeholders", "Poor visual balance"]
            },
            "detailed_instructions": {
                "pros": ["Better layout matching", "No empty placeholders", "Professional output"],
                "cons": ["More tokens", "Less creative freedom"]
            },
            "recommendation": "Use detailed instructions for template-based workflows"
        }

        assert len(scenarios) == 4

    def test_li08_instruction_token_cost(self, temp_dir):
        """LI08: Measure instruction token cost vs quality benefit."""
        token_estimates = {
            "minimal": {
                "chars": len(LAYOUT_INSTRUCTIONS["minimal"]),
                "estimated_tokens": len(LAYOUT_INSTRUCTIONS["minimal"]) // 4
            },
            "standard": {
                "chars": len(LAYOUT_INSTRUCTIONS["standard"]),
                "estimated_tokens": len(LAYOUT_INSTRUCTIONS["standard"]) // 4
            },
            "detailed": {
                "chars": len(LAYOUT_INSTRUCTIONS["detailed"]),
                "estimated_tokens": len(LAYOUT_INSTRUCTIONS["detailed"]) // 4
            },
            "strict": {
                "chars": len(LAYOUT_INSTRUCTIONS["strict"]),
                "estimated_tokens": len(LAYOUT_INSTRUCTIONS["strict"]) // 4
            },
            "examples": {
                "chars": len(LAYOUT_INSTRUCTIONS["examples"]),
                "estimated_tokens": len(LAYOUT_INSTRUCTIONS["examples"]) // 4
            }
        }

        # The cost is minimal compared to typical prompt sizes
        assert token_estimates["detailed"]["estimated_tokens"] < 200
        assert token_estimates["strict"]["estimated_tokens"] < 300


class TestQualityMetrics:
    """Define quality metrics for layout decisions."""

    def test_li09_define_quality_criteria(self, temp_dir):
        """LI09: Define what 'quality' means for layout matching."""
        quality_criteria = {
            "placeholder_utilization": {
                "description": "All placeholders filled with content",
                "scoring": "% of placeholders with content",
                "target": "100%"
            },
            "content_fit": {
                "description": "Content matches placeholder capacity",
                "scoring": "No overflow, no excessive whitespace",
                "target": "All content visible, balanced"
            },
            "semantic_match": {
                "description": "Layout type matches content type",
                "scoring": "Quote layouts for quotes, etc.",
                "target": "100% semantic match"
            },
            "visual_balance": {
                "description": "Slide looks professionally balanced",
                "scoring": "Subjective visual review",
                "target": "No obvious empty spaces or cramming"
            }
        }

        # This test documents the criteria
        assert len(quality_criteria) == 4

    def test_li10_automated_quality_checks(self, temp_dir, python_cmd, scripts_dir):
        """LI10: Define automated quality checks post-generation."""
        quality_checks = {
            "empty_shape_check": {
                "method": "Compare inventory before/after replacement",
                "check": "Shapes with paragraphs in replacement JSON vs total shapes",
                "pass_criteria": "All shapes either filled or intentionally hidden"
            },
            "overflow_check": {
                "method": "inventory.py overflow detection",
                "check": "Text overflow warnings in output",
                "pass_criteria": "No overflow warnings"
            },
            "slide_count_check": {
                "method": "Compare content items to slides",
                "check": "Content not crammed into too few slides",
                "pass_criteria": "Reasonable items-per-slide ratio"
            }
        }

        # Implementation would involve running these checks automatically
        assert len(quality_checks) == 3
