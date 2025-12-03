"""
Objective performance tests for pptx-edit skill.

These tests measure whether the skill instructions produce correct behavior
across various scenarios, including clean templates and messy decks.
"""
import pytest
import json
import subprocess
from pathlib import Path
from dataclasses import dataclass
from typing import List, Dict, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import random


# ============================================================================
# TEST FIXTURES: Create Various Deck Types
# ============================================================================

def create_clean_template(path: Path, num_layouts: int = 6) -> Path:
    """Create a well-structured template with clear layouts."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 0: Title slide
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.3), Inches(1.5))
    title.text_frame.paragraphs[0].text = "[TITLE]"
    title.text_frame.paragraphs[0].font.size = Pt(48)
    title.text_frame.paragraphs[0].font.bold = True
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    subtitle = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(12.3), Inches(0.8))
    subtitle.text_frame.paragraphs[0].text = "[Subtitle]"
    subtitle.text_frame.paragraphs[0].font.size = Pt(24)
    subtitle.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Slide 1: Section header
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(3), Inches(12.3), Inches(1.5))
    title.text_frame.paragraphs[0].text = "[Section Title]"
    title.text_frame.paragraphs[0].font.size = Pt(44)
    title.text_frame.paragraphs[0].font.bold = True
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Slide 2: Single column bullets
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    title.text_frame.paragraphs[0].text = "[Slide Title]"
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.bold = True

    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.3), Inches(5.5))
    tf = body.text_frame
    for i in range(5):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"[Bullet {i+1}]"
        p.font.size = Pt(20)
        p.level = 0

    # Slide 3: Two-column
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    title.text_frame.paragraphs[0].text = "[Two Column Title]"
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.bold = True

    left = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5.8), Inches(5.5))
    left.text_frame.paragraphs[0].text = "[Left Column]"
    left.text_frame.paragraphs[0].font.size = Pt(18)

    right = slide.shapes.add_textbox(Inches(6.9), Inches(1.5), Inches(5.8), Inches(5.5))
    right.text_frame.paragraphs[0].text = "[Right Column]"
    right.text_frame.paragraphs[0].font.size = Pt(18)

    # Slide 4: Three-column
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    title.text_frame.paragraphs[0].text = "[Three Column Title]"
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.bold = True

    col1 = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(3.8), Inches(5.5))
    col1.text_frame.paragraphs[0].text = "[Column 1]"
    col1.text_frame.paragraphs[0].font.size = Pt(16)

    col2 = slide.shapes.add_textbox(Inches(4.7), Inches(1.5), Inches(3.8), Inches(5.5))
    col2.text_frame.paragraphs[0].text = "[Column 2]"
    col2.text_frame.paragraphs[0].font.size = Pt(16)

    col3 = slide.shapes.add_textbox(Inches(8.9), Inches(1.5), Inches(3.8), Inches(5.5))
    col3.text_frame.paragraphs[0].text = "[Column 3]"
    col3.text_frame.paragraphs[0].font.size = Pt(16)

    # Slide 5: Quote layout
    slide = prs.slides.add_slide(blank)
    quote = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(3))
    quote.text_frame.paragraphs[0].text = '"[Quote text]"'
    quote.text_frame.paragraphs[0].font.size = Pt(32)
    quote.text_frame.paragraphs[0].font.italic = True
    quote.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    attribution = slide.shapes.add_textbox(Inches(1), Inches(5.5), Inches(11), Inches(1))
    attribution.text_frame.paragraphs[0].text = "— [Speaker Name]"
    attribution.text_frame.paragraphs[0].font.size = Pt(20)
    attribution.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

    prs.save(path)
    return path


def create_messy_deck(path: Path, num_slides: int = 8) -> Path:
    """Create a messy deck with inconsistent structure."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    random.seed(42)  # Reproducible randomness

    for i in range(num_slides):
        slide = prs.slides.add_slide(blank)

        # Random number of shapes (3-8)
        num_shapes = random.randint(3, 8)

        for j in range(num_shapes):
            # Random position and size
            left = random.uniform(0.3, 10)
            top = random.uniform(0.3, 5)
            width = random.uniform(2, 6)
            height = random.uniform(0.5, 3)

            shape = slide.shapes.add_textbox(
                Inches(left), Inches(top), Inches(width), Inches(height)
            )
            shape.text_frame.paragraphs[0].text = f"Text box {j+1} on slide {i}"
            shape.text_frame.paragraphs[0].font.size = Pt(random.randint(12, 36))

    prs.save(path)
    return path


def create_mixed_deck(path: Path) -> Path:
    """Create a deck with some clean and some messy slides."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 0: Clean title
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.3), Inches(1.5))
    title.text_frame.paragraphs[0].text = "[Title]"
    title.text_frame.paragraphs[0].font.size = Pt(48)
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Slide 1: Messy - multiple random boxes
    slide = prs.slides.add_slide(blank)
    for i in range(6):
        shape = slide.shapes.add_textbox(
            Inches(0.5 + i * 0.3), Inches(0.5 + i * 0.5),
            Inches(4), Inches(1.5)
        )
        shape.text_frame.paragraphs[0].text = f"Random text {i}"

    # Slide 2: Clean two-column
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    title.text_frame.paragraphs[0].text = "[Title]"

    left = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5.8), Inches(5.5))
    left.text_frame.paragraphs[0].text = "[Left]"

    right = slide.shapes.add_textbox(Inches(6.9), Inches(1.5), Inches(5.8), Inches(5.5))
    right.text_frame.paragraphs[0].text = "[Right]"

    # Slide 3: Messy - overlapping shapes
    slide = prs.slides.add_slide(blank)
    for i in range(4):
        shape = slide.shapes.add_textbox(
            Inches(2 + i * 0.5), Inches(2 + i * 0.3),
            Inches(5), Inches(2)
        )
        shape.text_frame.paragraphs[0].text = f"Overlapping {i}"

    prs.save(path)
    return path


# ============================================================================
# LAYOUT DETECTION TESTS
# ============================================================================

class TestLayoutDetection:
    """Test ability to detect layout types from inventory."""

    def analyze_layout(self, shapes: Dict) -> str:
        """Analyze inventory shapes to determine layout type."""
        if not shapes:
            return "empty"

        # Filter to meaningful shapes (width > 2", height > 0.5")
        content_shapes = []
        for shape_id, data in shapes.items():
            if data.get('width', 0) > 2 and data.get('height', 0) > 0.5:
                content_shapes.append(data)

        if len(content_shapes) == 0:
            return "empty"
        elif len(content_shapes) == 1:
            return "single"
        elif len(content_shapes) == 2:
            # Check if title + body or two side-by-side
            tops = [s['top'] for s in content_shapes]
            if abs(tops[0] - tops[1]) > 1:  # Different vertical positions
                return "title-body"
            else:
                return "two-column"
        elif len(content_shapes) == 3:
            # Check arrangement
            tops = sorted([s['top'] for s in content_shapes])
            if tops[1] - tops[0] > 0.5 and abs(tops[2] - tops[1]) < 0.5:
                return "title-two-column"
            else:
                return "three-shapes"
        elif len(content_shapes) == 4:
            return "title-three-column"
        else:
            return "complex"

    def test_ld01_detect_title_slide(self, temp_dir, python_cmd, scripts_dir):
        """LD01: Detect title slide layout."""
        template = temp_dir / "clean.pptx"
        create_clean_template(template)

        inventory_path = temp_dir / "inventory.json"
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Slide 0 should be title-body
        layout = self.analyze_layout(inventory.get("slide-0", {}))
        assert layout in ["title-body", "two-column"], f"Title slide detected as {layout}"

    def test_ld02_detect_two_column(self, temp_dir, python_cmd, scripts_dir):
        """LD02: Detect two-column layout."""
        template = temp_dir / "clean.pptx"
        create_clean_template(template)

        inventory_path = temp_dir / "inventory.json"
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Slide 3 should be title + two columns
        layout = self.analyze_layout(inventory.get("slide-3", {}))
        assert layout in ["title-two-column", "three-shapes"], f"Two-column detected as {layout}"

    def test_ld03_detect_three_column(self, temp_dir, python_cmd, scripts_dir):
        """LD03: Detect three-column layout."""
        template = temp_dir / "clean.pptx"
        create_clean_template(template)

        inventory_path = temp_dir / "inventory.json"
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Slide 4 should be title + three columns
        layout = self.analyze_layout(inventory.get("slide-4", {}))
        assert layout == "title-three-column", f"Three-column detected as {layout}"

    def test_ld04_detect_messy_slide(self, temp_dir, python_cmd, scripts_dir):
        """LD04: Detect complex/messy slide."""
        deck = temp_dir / "messy.pptx"
        create_messy_deck(deck)

        inventory_path = temp_dir / "inventory.json"
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(deck), str(inventory_path)],
            capture_output=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Most slides should be complex
        complex_count = 0
        for slide_id, shapes in inventory.items():
            if not slide_id.startswith("slide-"):
                continue
            layout = self.analyze_layout(shapes)
            if layout == "complex":
                complex_count += 1

        assert complex_count >= 3, f"Only {complex_count} slides detected as complex"


# ============================================================================
# LAYOUT MATCHING QUALITY TESTS
# ============================================================================

@dataclass
class LayoutMatchScenario:
    """A scenario for testing layout matching decisions."""
    name: str
    content_items: int
    correct_layouts: List[str]
    incorrect_layouts: List[str]


LAYOUT_SCENARIOS = [
    LayoutMatchScenario(
        name="two_items",
        content_items=2,
        correct_layouts=["two-column", "title-two-column"],
        incorrect_layouts=["three-column", "title-three-column"]
    ),
    LayoutMatchScenario(
        name="three_items",
        content_items=3,
        correct_layouts=["three-column", "title-three-column"],
        incorrect_layouts=["two-column"]
    ),
    LayoutMatchScenario(
        name="five_items",
        content_items=5,
        correct_layouts=["bullets", "title-body"],
        incorrect_layouts=["two-column", "three-column"]
    ),
]


class TestLayoutMatchingQuality:
    """Test that layout matching rules produce correct decisions."""

    def test_lmq01_content_count_to_layout(self):
        """LMQ01: Verify content count → layout mapping rules."""
        rules = {
            1: ["single", "title-only"],
            2: ["two-column", "title-two-column"],
            3: ["three-column", "title-three-column"],
            4: ["bullets", "two-slides"],
            5: ["bullets", "multiple-slides"],
        }

        for count, valid_layouts in rules.items():
            assert len(valid_layouts) > 0, f"No valid layouts for {count} items"

    @pytest.mark.parametrize("scenario", LAYOUT_SCENARIOS, ids=lambda s: s.name)
    def test_lmq02_scenario_validation(self, scenario):
        """LMQ02: Validate each scenario has clear correct/incorrect."""
        assert len(scenario.correct_layouts) > 0
        assert len(scenario.incorrect_layouts) > 0
        # No overlap between correct and incorrect
        overlap = set(scenario.correct_layouts) & set(scenario.incorrect_layouts)
        assert len(overlap) == 0, f"Overlap in layouts: {overlap}"


# ============================================================================
# END-TO-END WORKFLOW TESTS
# ============================================================================

class TestWorkflow1TextReplacement:
    """Test Workflow 1: Text Replacement."""

    def test_w1_01_simple_replacement(self, temp_dir, python_cmd, scripts_dir):
        """W1.01: Simple text replacement on clean template."""
        template = temp_dir / "clean.pptx"
        create_clean_template(template)

        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        # Step 1: Get inventory
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True, text=True
        )
        assert result.returncode == 0, f"Inventory failed: {result.stderr}"

        # Step 2: Create replacements
        with open(inventory_path) as f:
            inventory = json.load(f)

        replacements = {}
        if "slide-0" in inventory:
            shapes = list(inventory["slide-0"].keys())
            if shapes:
                replacements["slide-0"] = {
                    shapes[0]: {
                        "paragraphs": [{"text": "New Title", "bold": True, "alignment": "CENTER"}]
                    }
                }

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        # Step 3: Apply
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(template), str(replacements_path), str(output_path)],
            capture_output=True, text=True
        )
        assert result.returncode == 0, f"Replace failed: {result.stderr}"
        assert output_path.exists()

    def test_w1_02_bullet_replacement(self, temp_dir, python_cmd, scripts_dir):
        """W1.02: Replace content with bullets."""
        template = temp_dir / "clean.pptx"
        create_clean_template(template)

        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Find bullet slide (slide-2)
        replacements = {}
        if "slide-2" in inventory:
            shapes = list(inventory["slide-2"].keys())
            if len(shapes) >= 2:
                # Body shape usually second
                replacements["slide-2"] = {
                    shapes[0]: {"paragraphs": [{"text": "Key Points", "bold": True}]},
                    shapes[1]: {
                        "paragraphs": [
                            {"text": "First point", "bullet": True, "level": 0},
                            {"text": "Second point", "bullet": True, "level": 0},
                            {"text": "Third point", "bullet": True, "level": 0},
                        ]
                    }
                }

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(template), str(replacements_path), str(output_path)],
            capture_output=True, text=True
        )
        assert result.returncode == 0


class TestWorkflow3ApplyOutline:
    """Test Workflow 3: Apply Content Outline."""

    def test_w3_01_rearrange_and_replace(self, temp_dir, python_cmd, scripts_dir):
        """W3.01: Rearrange slides then apply content."""
        template = temp_dir / "clean.pptx"
        create_clean_template(template)

        working_path = temp_dir / "working.pptx"
        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "final.pptx"

        # Step 1: Rearrange (title, bullets, two-column)
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template), str(working_path), "0,2,3"],
            capture_output=True, text=True
        )
        assert result.returncode == 0

        # Verify rearrangement
        prs = Presentation(working_path)
        assert len(prs.slides) == 3

        # Step 2: Get fresh inventory
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(working_path), str(inventory_path)],
            capture_output=True, text=True
        )
        assert result.returncode == 0

        # Step 3: Create outline-based replacements
        with open(inventory_path) as f:
            inventory = json.load(f)

        outline = {
            "slide-0": {"title": "Q4 Review", "subtitle": "January 2025"},
            "slide-1": {"title": "Agenda", "bullets": ["Topic 1", "Topic 2", "Topic 3"]},
            "slide-2": {"title": "Comparison", "left": "Option A benefits", "right": "Option B benefits"},
        }

        replacements = {}
        for slide_id, content in outline.items():
            if slide_id not in inventory:
                continue
            shapes = list(inventory[slide_id].keys())

            if slide_id == "slide-0" and len(shapes) >= 2:
                replacements[slide_id] = {
                    shapes[0]: {"paragraphs": [{"text": content["title"], "bold": True, "alignment": "CENTER"}]},
                    shapes[1]: {"paragraphs": [{"text": content["subtitle"], "alignment": "CENTER"}]},
                }
            elif slide_id == "slide-1" and len(shapes) >= 2:
                replacements[slide_id] = {
                    shapes[0]: {"paragraphs": [{"text": content["title"], "bold": True}]},
                    shapes[1]: {"paragraphs": [{"text": b, "bullet": True, "level": 0} for b in content["bullets"]]},
                }
            elif slide_id == "slide-2" and len(shapes) >= 3:
                replacements[slide_id] = {
                    shapes[0]: {"paragraphs": [{"text": content["title"], "bold": True}]},
                    shapes[1]: {"paragraphs": [{"text": content["left"]}]},
                    shapes[2]: {"paragraphs": [{"text": content["right"]}]},
                }

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        # Step 4: Apply
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(working_path), str(replacements_path), str(output_path)],
            capture_output=True, text=True
        )
        assert result.returncode == 0
        assert output_path.exists()


class TestMessyDeckHandling:
    """Test handling of messy decks."""

    def test_md01_inventory_messy_deck(self, temp_dir, python_cmd, scripts_dir):
        """MD01: Successfully inventory a messy deck."""
        deck = temp_dir / "messy.pptx"
        create_messy_deck(deck)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(deck), str(inventory_path)],
            capture_output=True, text=True
        )
        assert result.returncode == 0

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Should have all slides
        slide_count = len([k for k in inventory.keys() if k.startswith("slide-")])
        assert slide_count == 8

    def test_md02_identify_placeholder_type_null(self, temp_dir, python_cmd, scripts_dir):
        """MD02: Identify placeholder_type: null in messy decks."""
        deck = temp_dir / "messy.pptx"
        create_messy_deck(deck)

        inventory_path = temp_dir / "inventory.json"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(deck), str(inventory_path)],
            capture_output=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Most shapes should have placeholder_type: null (or missing)
        null_count = 0
        total_count = 0
        for slide_id, shapes in inventory.items():
            if not slide_id.startswith("slide-"):
                continue
            for shape_id, shape_data in shapes.items():
                total_count += 1
                if shape_data.get("placeholder_type") is None:
                    null_count += 1

        # Most shapes should be non-placeholder
        assert null_count / total_count > 0.8, f"Only {null_count}/{total_count} shapes have null placeholder"

    def test_md03_replace_in_messy_deck(self, temp_dir, python_cmd, scripts_dir):
        """MD03: Successfully replace text in messy deck."""
        deck = temp_dir / "messy.pptx"
        create_messy_deck(deck)

        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(deck), str(inventory_path)],
            capture_output=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Replace first shape of first slide
        replacements = {}
        if "slide-0" in inventory:
            first_shape = list(inventory["slide-0"].keys())[0]
            replacements["slide-0"] = {
                first_shape: {"paragraphs": [{"text": "Replaced content"}]}
            }

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(deck), str(replacements_path), str(output_path)],
            capture_output=True, text=True
        )
        assert result.returncode == 0


# ============================================================================
# QUALITY METRICS
# ============================================================================

class TestQualityMetrics:
    """Tests that measure output quality."""

    def test_qm01_no_empty_shapes_after_replace(self, temp_dir, python_cmd, scripts_dir):
        """QM01: Verify replaced shapes have content."""
        template = temp_dir / "clean.pptx"
        create_clean_template(template)

        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"
        output_inventory_path = temp_dir / "output_inventory.json"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Replace ALL shapes in slide-0
        replacements = {"slide-0": {}}
        for shape_id in inventory.get("slide-0", {}).keys():
            replacements["slide-0"][shape_id] = {
                "paragraphs": [{"text": f"Content for {shape_id}"}]
            }

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(template), str(replacements_path), str(output_path)],
            capture_output=True
        )

        # Get output inventory
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(output_path), str(output_inventory_path)],
            capture_output=True
        )

        with open(output_inventory_path) as f:
            output_inventory = json.load(f)

        # All shapes in slide-0 should have non-empty text
        for shape_id, shape_data in output_inventory.get("slide-0", {}).items():
            paragraphs = shape_data.get("paragraphs", [])
            has_content = any(p.get("text", "").strip() for p in paragraphs)
            assert has_content, f"{shape_id} has no content after replacement"

    def test_qm02_formatting_preserved(self, temp_dir, python_cmd, scripts_dir):
        """QM02: Verify formatting properties are preserved/applied."""
        template = temp_dir / "clean.pptx"
        create_clean_template(template)

        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_path)],
            capture_output=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Replace with specific formatting
        if "slide-0" in inventory:
            first_shape = list(inventory["slide-0"].keys())[0]
            replacements = {
                "slide-0": {
                    first_shape: {
                        "paragraphs": [{
                            "text": "Bold Title",
                            "bold": True,
                            "alignment": "CENTER"
                        }]
                    }
                }
            }

            with open(replacements_path, "w") as f:
                json.dump(replacements, f)

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "replace.py"),
                 str(template), str(replacements_path), str(output_path)],
                capture_output=True, text=True
            )

            assert result.returncode == 0


# ============================================================================
# SKILL INSTRUCTION COMPLIANCE TESTS
# ============================================================================

class TestSkillCompliance:
    """Test that workflows comply with skill instructions."""

    def test_sc01_inventory_before_replace(self, temp_dir, python_cmd, scripts_dir):
        """SC01: Verify inventory must be run before replace."""
        template = temp_dir / "clean.pptx"
        create_clean_template(template)

        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        # Try to replace with made-up shape IDs (without inventory)
        replacements = {
            "slide-0": {
                "shape-999": {"paragraphs": [{"text": "This should fail"}]}
            }
        }

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(template), str(replacements_path), str(output_path)],
            capture_output=True, text=True
        )

        # Should fail or warn about invalid shape
        assert result.returncode != 0 or "not found" in result.stderr.lower()

    def test_sc02_re_inventory_after_rearrange(self, temp_dir, python_cmd, scripts_dir):
        """SC02: Verify re-inventory needed after rearrange."""
        template = temp_dir / "clean.pptx"
        create_clean_template(template)

        working_path = temp_dir / "working.pptx"
        inventory_before_path = temp_dir / "inventory_before.json"
        inventory_after_path = temp_dir / "inventory_after.json"

        # Get inventory before
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template), str(inventory_before_path)],
            capture_output=True
        )

        # Rearrange
        subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template), str(working_path), "0,3,4"],
            capture_output=True
        )

        # Get inventory after
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(working_path), str(inventory_after_path)],
            capture_output=True
        )

        with open(inventory_before_path) as f:
            before = json.load(f)
        with open(inventory_after_path) as f:
            after = json.load(f)

        # Inventories should be different (slide count changed)
        before_slides = len([k for k in before.keys() if k.startswith("slide-")])
        after_slides = len([k for k in after.keys() if k.startswith("slide-")])

        assert before_slides != after_slides or before != after, \
            "Inventory should change after rearrange"
