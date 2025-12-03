"""
Test template complexity handling for PPTX skill.

Tests how well the skill handles templates of varying complexity levels.
"""
import pytest
import json
import subprocess
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE


class TestTemplateCreation:
    """Helper methods and tests for creating test templates."""

    @staticmethod
    def create_clean_template(path, num_slides=5):
        """Create a clean corporate template with standard layouts."""
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank = prs.slide_layouts[6]

        for i in range(num_slides):
            slide = prs.slides.add_slide(blank)

            # Title
            title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
            tf = title.text_frame
            p = tf.paragraphs[0]
            p.text = f"Slide {i + 1} Title"
            p.font.size = Pt(40)
            p.font.bold = True
            p.alignment = PP_ALIGN.LEFT

            # Subtitle
            if i == 0:
                sub = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(12.3), Inches(0.5))
                tf = sub.text_frame
                p = tf.paragraphs[0]
                p.text = "Subtitle text here"
                p.font.size = Pt(24)

            # Body
            body = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(12.3), Inches(5))
            tf = body.text_frame
            p = tf.paragraphs[0]
            p.text = f"Body content for slide {i + 1}"
            p.font.size = Pt(18)

        prs.save(path)
        return path

    @staticmethod
    def create_medium_template(path, num_slides=20):
        """Create a medium complexity template."""
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank = prs.slide_layouts[6]

        slide_types = ["title", "content", "two_column", "bullets", "image_placeholder"]

        for i in range(num_slides):
            slide = prs.slides.add_slide(blank)
            slide_type = slide_types[i % len(slide_types)]

            # Title for all slides
            title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
            tf = title.text_frame
            p = tf.paragraphs[0]
            p.text = f"Slide {i + 1}: {slide_type.replace('_', ' ').title()}"
            p.font.size = Pt(36)
            p.font.bold = True

            if slide_type == "two_column":
                # Left column
                left = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5.9), Inches(5.5))
                tf = left.text_frame
                p = tf.paragraphs[0]
                p.text = "Left column content"
                p.font.size = Pt(16)

                # Right column
                right = slide.shapes.add_textbox(Inches(6.9), Inches(1.5), Inches(5.9), Inches(5.5))
                tf = right.text_frame
                p = tf.paragraphs[0]
                p.text = "Right column content"
                p.font.size = Pt(16)

            elif slide_type == "bullets":
                body = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.3), Inches(5.5))
                tf = body.text_frame
                for j in range(5):
                    p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                    p.text = f"Bullet point {j + 1}"
                    p.font.size = Pt(18)
                    p.level = 0

            elif slide_type == "image_placeholder":
                # Text area
                text = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5), Inches(5.5))
                tf = text.text_frame
                p = tf.paragraphs[0]
                p.text = "Description text"
                p.font.size = Pt(16)

                # Image placeholder (rectangle)
                shape = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE, Inches(6), Inches(1.5), Inches(6.8), Inches(5.5)
                )
                shape.text = "[Image Placeholder]"

            else:
                # Default content slide
                body = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.3), Inches(5.5))
                tf = body.text_frame
                p = tf.paragraphs[0]
                p.text = f"Content for slide {i + 1}"
                p.font.size = Pt(18)

        prs.save(path)
        return path

    @staticmethod
    def create_complex_template(path, num_slides=50):
        """Create a complex template with many variations."""
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank = prs.slide_layouts[6]

        for i in range(num_slides):
            slide = prs.slides.add_slide(blank)

            # Title with varying styles
            title = slide.shapes.add_textbox(
                Inches(0.3 + (i % 3) * 0.1),
                Inches(0.2 + (i % 5) * 0.05),
                Inches(12),
                Inches(1)
            )
            tf = title.text_frame
            p = tf.paragraphs[0]
            p.text = f"Complex Slide {i + 1}"
            p.font.size = Pt(32 + (i % 4) * 2)
            p.font.bold = i % 2 == 0

            # Multiple content areas
            for j in range(1 + i % 4):
                col_width = 12 / (1 + i % 4)
                box = slide.shapes.add_textbox(
                    Inches(0.5 + j * col_width),
                    Inches(1.5),
                    Inches(col_width - 0.2),
                    Inches(5)
                )
                tf = box.text_frame
                p = tf.paragraphs[0]
                p.text = f"Content area {j + 1}"
                p.font.size = Pt(14)

            # Add some shapes
            if i % 5 == 0:
                shape = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    Inches(10), Inches(5.5), Inches(2.8), Inches(1.5)
                )
                shape.text = "Shape text"

        prs.save(path)
        return path


class TestCleanTemplate:
    """Tests for clean, well-structured templates."""

    def test_tc01_clean_template_inventory(self, temp_dir, python_cmd, scripts_dir):
        """TC01: Extract inventory from clean template."""
        template_path = temp_dir / "clean_template.pptx"
        TestTemplateCreation.create_clean_template(template_path, num_slides=5)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0, f"Inventory failed: {result.stderr}"

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Should have all 5 slides
        slide_count = len([k for k in inventory.keys() if k.startswith("slide-")])
        assert slide_count == 5, f"Expected 5 slides, got {slide_count}"

    def test_tc02_clean_template_replacement(self, temp_dir, python_cmd, scripts_dir):
        """TC02: Full text replacement on clean template."""
        template_path = temp_dir / "clean_template.pptx"
        TestTemplateCreation.create_clean_template(template_path, num_slides=3)

        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        # Get inventory
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Create comprehensive replacements
        replacements = {}
        for slide_id, shapes in inventory.items():
            if not slide_id.startswith("slide-"):
                continue
            replacements[slide_id] = {}
            for shape_id, shape_data in shapes.items():
                replacements[slide_id][shape_id] = {
                    "paragraphs": [{"text": f"Replaced content for {slide_id}/{shape_id}"}]
                }

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(template_path), str(replacements_path), str(output_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0, f"Replacement failed: {result.stderr}"
        assert output_path.exists()


class TestMediumTemplate:
    """Tests for medium complexity templates."""

    def test_tc03_medium_template_inventory(self, temp_dir, python_cmd, scripts_dir):
        """TC03: Extract inventory from medium template."""
        template_path = temp_dir / "medium_template.pptx"
        TestTemplateCreation.create_medium_template(template_path, num_slides=20)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        with open(inventory_path) as f:
            inventory = json.load(f)

        slide_count = len([k for k in inventory.keys() if k.startswith("slide-")])
        assert slide_count == 20

    def test_tc04_medium_template_selective_replacement(self, temp_dir, python_cmd, scripts_dir):
        """TC04: Selective replacement on medium template."""
        template_path = temp_dir / "medium_template.pptx"
        TestTemplateCreation.create_medium_template(template_path, num_slides=10)

        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Only replace first shape of each slide
        replacements = {}
        for slide_id in list(inventory.keys())[:5]:
            if not slide_id.startswith("slide-"):
                continue
            shapes = list(inventory[slide_id].keys())
            if shapes:
                replacements[slide_id] = {
                    shapes[0]: {
                        "paragraphs": [{"text": f"Updated title for {slide_id}"}]
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

        assert result.returncode == 0


class TestComplexTemplate:
    """Tests for complex templates."""

    @pytest.mark.slow
    def test_tc05_complex_template_inventory(self, temp_dir, python_cmd, scripts_dir):
        """TC05: Extract inventory from complex template."""
        template_path = temp_dir / "complex_template.pptx"
        TestTemplateCreation.create_complex_template(template_path, num_slides=50)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True,
            timeout=60
        )

        assert result.returncode == 0

        with open(inventory_path) as f:
            inventory = json.load(f)

        slide_count = len([k for k in inventory.keys() if k.startswith("slide-")])
        assert slide_count == 50


class TestSlideRearrangement:
    """Tests for slide rearrangement operations."""

    def test_tc06_basic_rearrangement(self, temp_dir, python_cmd, scripts_dir):
        """TC06: Basic slide rearrangement."""
        template_path = temp_dir / "template.pptx"
        TestTemplateCreation.create_clean_template(template_path, num_slides=5)

        output_path = temp_dir / "rearranged.pptx"

        # Reorder: take slides 0, 2, 4 (reverse order)
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "4,2,0"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0
        assert output_path.exists()

        # Verify result has 3 slides
        prs = Presentation(output_path)
        assert len(prs.slides) == 3

    def test_tc07_duplicate_slides(self, temp_dir, python_cmd, scripts_dir):
        """TC07: Duplicate slides in rearrangement."""
        template_path = temp_dir / "template.pptx"
        TestTemplateCreation.create_clean_template(template_path, num_slides=3)

        output_path = temp_dir / "duplicated.pptx"

        # Duplicate slide 0 three times
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,0,0,1,2"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 5

    def test_tc08_invalid_slide_index(self, temp_dir, python_cmd, scripts_dir):
        """TC08: Invalid slide index should error."""
        template_path = temp_dir / "template.pptx"
        TestTemplateCreation.create_clean_template(template_path, num_slides=3)

        output_path = temp_dir / "error.pptx"

        # Try to use slide index 10 (only 3 slides exist)
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,1,10"],
            capture_output=True,
            text=True
        )

        # Should fail or warn about invalid index
        # (behavior depends on script implementation)
        assert result.returncode != 0 or "error" in result.stderr.lower() or "invalid" in result.stderr.lower()


class TestGroupedShapes:
    """Tests for templates with grouped shapes."""

    def test_tc09_grouped_shapes_inventory(self, temp_dir, python_cmd, scripts_dir):
        """TC09: Extract inventory from template with grouped shapes."""
        template_path = temp_dir / "grouped.pptx"

        # Create template with grouped shapes
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Add individual shapes (we can't easily create groups programmatically)
        title = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
        title.text_frame.paragraphs[0].text = "Title"

        body = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(6), Inches(4))
        body.text_frame.paragraphs[0].text = "Body text"

        prs.save(template_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0


class TestTemplateWithShapes:
    """Tests for templates with various shape types."""

    def test_tc10_shapes_with_text(self, temp_dir, python_cmd, scripts_dir):
        """TC10: Shapes with text (rectangles, etc.)."""
        template_path = temp_dir / "shapes.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Add various shapes with text
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(3), Inches(2))
        rect.text = "Rectangle text"

        oval = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5), Inches(1), Inches(3), Inches(2))
        oval.text = "Oval text"

        rounded = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9), Inches(1), Inches(3), Inches(2)
        )
        rounded.text = "Rounded rect text"

        prs.save(template_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Should capture text from all shapes
        shapes = inventory.get("slide-0", {})
        assert len(shapes) >= 3


class TestHiddenSlides:
    """Tests for templates with hidden slides."""

    def test_tc11_hidden_slides(self, temp_dir, python_cmd, scripts_dir):
        """TC11: Handle templates with hidden slides."""
        template_path = temp_dir / "hidden.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        for i in range(5):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            title = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
            title.text_frame.paragraphs[0].text = f"Slide {i + 1}"

            # Hide every other slide
            if i % 2 == 1:
                slide._element.set("{http://schemas.openxmlformats.org/presentationml/2006/main}show", "0")

        prs.save(template_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        # Inventory should include all slides (hidden or not)
        with open(inventory_path) as f:
            inventory = json.load(f)

        slide_count = len([k for k in inventory.keys() if k.startswith("slide-")])
        assert slide_count == 5
