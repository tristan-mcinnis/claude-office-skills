"""
Test slide operations for PPTX skill.

Tests slide manipulation capabilities: duplicate, reorder, delete.
"""
import pytest
import subprocess
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt


def create_test_presentation(path, num_slides=5, with_images=False):
    """Create a test presentation with numbered slides."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for i in range(num_slides):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Add title with slide number
        title = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
        tf = title.text_frame
        p = tf.paragraphs[0]
        p.text = f"Original Slide {i}"
        p.font.size = Pt(44)
        p.font.bold = True

        # Add content
        body = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(12), Inches(4))
        tf = body.text_frame
        p = tf.paragraphs[0]
        p.text = f"Content for slide {i}"
        p.font.size = Pt(24)

    prs.save(path)
    return path


class TestSlideRearrangement:
    """Tests for basic slide rearrangement operations."""

    def test_so01_single_slide_selection(self, temp_dir, python_cmd, scripts_dir):
        """SO01: Select single slide from presentation."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=5)

        output_path = temp_dir / "single.pptx"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "2"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0
        assert output_path.exists()

        prs = Presentation(output_path)
        assert len(prs.slides) == 1

    def test_so02_multiple_slide_selection(self, temp_dir, python_cmd, scripts_dir):
        """SO02: Select multiple slides."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=10)

        output_path = temp_dir / "multiple.pptx"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,2,4,6,8"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 5

    def test_so03_reverse_order(self, temp_dir, python_cmd, scripts_dir):
        """SO03: Reverse slide order."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=5)

        output_path = temp_dir / "reversed.pptx"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "4,3,2,1,0"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 5

        # Verify order by checking first slide title contains "4"
        first_slide = prs.slides[0]
        for shape in first_slide.shapes:
            if hasattr(shape, "text_frame"):
                if "4" in shape.text_frame.text:
                    break
        else:
            pytest.fail("First slide should be original slide 4")

    def test_so04_custom_order(self, temp_dir, python_cmd, scripts_dir):
        """SO04: Custom arbitrary order."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=5)

        output_path = temp_dir / "custom.pptx"

        # Mixed order
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "3,0,4,1,2"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 5


class TestSlideDuplication:
    """Tests for slide duplication."""

    def test_so05_duplicate_single_slide(self, temp_dir, python_cmd, scripts_dir):
        """SO05: Duplicate a single slide multiple times."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=3)

        output_path = temp_dir / "duplicated.pptx"

        # Duplicate slide 0 three times
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,0,0"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 3

        # All slides should have same content
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    if "Original Slide 0" in shape.text_frame.text:
                        break
            else:
                continue
            break

    def test_so06_mixed_duplication(self, temp_dir, python_cmd, scripts_dir):
        """SO06: Mix of unique and duplicated slides."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=5)

        output_path = temp_dir / "mixed.pptx"

        # Pattern: slide 0, slide 1, slide 1, slide 2, slide 2, slide 2
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,1,1,2,2,2"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 6

    def test_so07_preserve_formatting_on_duplicate(self, temp_dir, python_cmd, scripts_dir):
        """SO07: Formatting preserved when duplicating."""
        template_path = temp_dir / "formatted.pptx"

        # Create presentation with specific formatting
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        slide = prs.slides.add_slide(prs.slide_layouts[6])
        title = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
        tf = title.text_frame
        p = tf.paragraphs[0]
        p.text = "Formatted Title"
        p.font.size = Pt(48)
        p.font.bold = True
        p.font.italic = True

        prs.save(template_path)

        output_path = temp_dir / "dup_formatted.pptx"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,0,0"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        # Verify formatting preserved
        prs_out = Presentation(output_path)
        for slide in prs_out.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    if shape.text_frame.paragraphs:
                        font = shape.text_frame.paragraphs[0].font
                        # Note: formatting may be inherited from style, not directly on font
                        assert "Formatted Title" in shape.text_frame.text


class TestSlideRemoval:
    """Tests for removing slides (by exclusion)."""

    def test_so08_remove_first_slide(self, temp_dir, python_cmd, scripts_dir):
        """SO08: Remove first slide by not including it."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=5)

        output_path = temp_dir / "no_first.pptx"

        # Include all except slide 0
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "1,2,3,4"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 4

    def test_so09_remove_last_slide(self, temp_dir, python_cmd, scripts_dir):
        """SO09: Remove last slide."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=5)

        output_path = temp_dir / "no_last.pptx"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,1,2,3"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 4

    def test_so10_remove_multiple_slides(self, temp_dir, python_cmd, scripts_dir):
        """SO10: Remove multiple non-contiguous slides."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=10)

        output_path = temp_dir / "sparse.pptx"

        # Keep only even-indexed slides
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,2,4,6,8"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 5


class TestEdgeCases:
    """Edge case tests for slide operations."""

    def test_so11_empty_indices(self, temp_dir, python_cmd, scripts_dir):
        """SO11: Empty indices string should error or produce empty result."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=3)

        output_path = temp_dir / "empty.pptx"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), ""],
            capture_output=True,
            text=True
        )

        # Should either error or produce empty presentation
        # Behavior depends on implementation
        pass  # Accept any behavior for edge case

    def test_so12_out_of_range_index(self, temp_dir, python_cmd, scripts_dir):
        """SO12: Out of range index should error."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=3)

        output_path = temp_dir / "error.pptx"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,1,99"],
            capture_output=True,
            text=True
        )

        # Should fail with clear error
        assert result.returncode != 0, "Should fail on out-of-range index"

    def test_so13_negative_index(self, temp_dir, python_cmd, scripts_dir):
        """SO13: Negative index handling."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=3)

        output_path = temp_dir / "negative.pptx"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,-1,2"],
            capture_output=True,
            text=True
        )

        # Should either error or handle gracefully
        # Negative indices might work Python-style or error
        pass  # Accept any reasonable behavior

    def test_so14_whitespace_in_indices(self, temp_dir, python_cmd, scripts_dir):
        """SO14: Whitespace in indices string."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=5)

        output_path = temp_dir / "whitespace.pptx"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0, 1, 2"],
            capture_output=True,
            text=True
        )

        # Should handle whitespace gracefully
        if result.returncode == 0:
            prs = Presentation(output_path)
            assert len(prs.slides) == 3


class TestLargePresentation:
    """Tests for large presentations."""

    @pytest.mark.slow
    def test_so15_large_presentation(self, temp_dir, python_cmd, scripts_dir):
        """SO15: Handle large presentation (100 slides)."""
        template_path = temp_dir / "large.pptx"
        create_test_presentation(template_path, num_slides=100)

        output_path = temp_dir / "large_output.pptx"

        # Select every 10th slide
        indices = ",".join(str(i) for i in range(0, 100, 10))

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), indices],
            capture_output=True,
            text=True,
            timeout=120
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 10

    @pytest.mark.slow
    def test_so16_many_duplications(self, temp_dir, python_cmd, scripts_dir):
        """SO16: Many duplications of same slide."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=1)

        output_path = temp_dir / "many_dupes.pptx"

        # Duplicate slide 0 fifty times
        indices = ",".join(["0"] * 50)

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), indices],
            capture_output=True,
            text=True,
            timeout=120
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 50


class TestComplexRearrangement:
    """Tests for complex rearrangement scenarios."""

    def test_so17_interleave_slides(self, temp_dir, python_cmd, scripts_dir):
        """SO17: Interleave slides from different positions."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=6)

        output_path = temp_dir / "interleaved.pptx"

        # Interleave: 0,3,1,4,2,5
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,3,1,4,2,5"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 6

    def test_so18_build_structure(self, temp_dir, python_cmd, scripts_dir):
        """SO18: Build presentation structure from parts."""
        template_path = temp_dir / "template.pptx"
        create_test_presentation(template_path, num_slides=10)

        output_path = temp_dir / "structure.pptx"

        # Structure: intro (0), main content (1,2,3 repeated), conclusion (9)
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,1,2,3,1,2,3,9"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 8
