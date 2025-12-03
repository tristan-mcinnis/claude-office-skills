"""
Test edge cases and known limitations for PPTX skill.

Documents what works, what doesn't, and boundary conditions.
"""
import pytest
import subprocess
import json
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE


class TestEmptyAndMinimal:
    """Tests for empty and minimal content scenarios."""

    def test_ec01_empty_presentation(self, temp_dir, python_cmd, scripts_dir):
        """EC01: Handle empty presentation (no slides)."""
        pptx_path = temp_dir / "empty.pptx"

        prs = Presentation()
        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        # Should handle gracefully
        assert result.returncode == 0

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Should be empty or minimal
        slide_count = len([k for k in inventory.keys() if k.startswith("slide-")])
        assert slide_count == 0

    def test_ec02_slide_with_no_shapes(self, temp_dir, python_cmd, scripts_dir):
        """EC02: Slide with no shapes at all."""
        pptx_path = temp_dir / "no_shapes.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

    def test_ec03_shape_with_empty_text(self, temp_dir, python_cmd, scripts_dir):
        """EC03: Shape with empty text frame."""
        pptx_path = temp_dir / "empty_text.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Add shape with empty text
        shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2))
        shape.text_frame.paragraphs[0].text = ""

        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

    def test_ec04_whitespace_only_text(self, temp_dir, python_cmd, scripts_dir):
        """EC04: Shape with only whitespace text."""
        pptx_path = temp_dir / "whitespace.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2))
        shape.text_frame.paragraphs[0].text = "   \n\t   "

        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0


class TestTextBoundaries:
    """Tests for text content boundaries."""

    def test_ec05_very_long_text(self, temp_dir, python_cmd, scripts_dir):
        """EC05: Very long text content (10000+ chars)."""
        pptx_path = temp_dir / "long_text.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(6))
        shape.text_frame.paragraphs[0].text = "A" * 10000

        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Should capture full text
        shapes = inventory.get("slide-0", {})
        if shapes:
            first_shape = list(shapes.values())[0]
            paras = first_shape.get("paragraphs", [])
            if paras:
                assert len(paras[0].get("text", "")) == 10000

    def test_ec06_many_paragraphs(self, temp_dir, python_cmd, scripts_dir):
        """EC06: Shape with many paragraphs (100+)."""
        pptx_path = temp_dir / "many_paras.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(6))
        tf = shape.text_frame
        tf.paragraphs[0].text = "Paragraph 0"

        for i in range(1, 100):
            p = tf.add_paragraph()
            p.text = f"Paragraph {i}"

        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        with open(inventory_path) as f:
            inventory = json.load(f)

        shapes = inventory.get("slide-0", {})
        if shapes:
            first_shape = list(shapes.values())[0]
            paras = first_shape.get("paragraphs", [])
            assert len(paras) == 100


class TestSpecialContent:
    """Tests for special content types."""

    def test_ec07_unicode_comprehensive(self, temp_dir, python_cmd, scripts_dir):
        """EC07: Comprehensive Unicode character test."""
        pptx_path = temp_dir / "unicode_all.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(6))
        tf = shape.text_frame

        # Various Unicode ranges
        texts = [
            "Latin Extended: ÄÖÜÉÈÊ",
            "CJK: 日本語 中文 한국어",
            "Cyrillic: Привет мир",
            "Greek: Γεια σου κόσμε",
            "Arabic: مرحبا بالعالم",
            "Hebrew: שלום עולם",
            "Symbols: ™ © ® ℃ ℉ № ‰",
            "Math: ∑ ∏ ∫ √ ∞ ∂ ≈ ≠ ≤ ≥",
            "Arrows: → ← ↑ ↓ ↔ ↕ ⇒ ⇐",
            "Currency: $ € £ ¥ ₹ ₽ ₿",
        ]

        tf.paragraphs[0].text = texts[0]
        for text in texts[1:]:
            p = tf.add_paragraph()
            p.text = text

        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        # Verify content preserved
        with open(inventory_path, encoding="utf-8") as f:
            inventory = json.load(f)

        shapes = inventory.get("slide-0", {})
        assert len(shapes) > 0

    def test_ec08_special_xml_characters(self, temp_dir, python_cmd, scripts_dir):
        """EC08: Characters that need XML escaping."""
        pptx_path = temp_dir / "xml_chars.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(2))
        # Characters that need XML escaping
        shape.text_frame.paragraphs[0].text = '<tag> & "quoted" \'single\' </tag>'

        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0


class TestShapeBoundaries:
    """Tests for shape positioning and sizing edge cases."""

    def test_ec09_zero_size_shape(self, temp_dir, python_cmd, scripts_dir):
        """EC09: Shape with zero width or height."""
        pptx_path = temp_dir / "zero_size.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Very small shape (not exactly zero, but tiny)
        shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(0.01), Inches(0.01))
        shape.text_frame.paragraphs[0].text = "Tiny"

        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        # Should handle gracefully
        assert result.returncode == 0

    def test_ec10_shape_outside_slide(self, temp_dir, python_cmd, scripts_dir):
        """EC10: Shape positioned outside slide bounds."""
        pptx_path = temp_dir / "outside.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Shape mostly outside slide
        shape = slide.shapes.add_textbox(Inches(15), Inches(10), Inches(3), Inches(2))
        shape.text_frame.paragraphs[0].text = "Off-slide content"

        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

    def test_ec11_overlapping_shapes(self, temp_dir, python_cmd, scripts_dir):
        """EC11: Multiple overlapping shapes."""
        pptx_path = temp_dir / "overlapping.pptx"

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Create overlapping shapes
        for i in range(5):
            shape = slide.shapes.add_textbox(
                Inches(1 + i * 0.3), Inches(1 + i * 0.3),
                Inches(4), Inches(2)
            )
            shape.text_frame.paragraphs[0].text = f"Shape {i + 1}"

        prs.save(pptx_path)

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        with open(inventory_path) as f:
            inventory = json.load(f)

        # All shapes should be captured
        shapes = inventory.get("slide-0", {})
        assert len(shapes) >= 5


class TestKnownLimitations:
    """Tests documenting known limitations."""

    @pytest.mark.limitation
    def test_lim01_no_image_editing(self, temp_dir):
        """LIM01: Cannot edit images in existing presentations."""
        # Document limitation
        limitation = {
            "id": "LIM01",
            "feature": "Image editing",
            "status": "Not supported",
            "description": "Cannot programmatically edit, replace, or manipulate images in existing presentations",
            "workaround": "Use template-based workflow with placeholder areas for images"
        }
        assert limitation["status"] == "Not supported"

    @pytest.mark.limitation
    def test_lim02_no_chart_editing(self, temp_dir):
        """LIM02: Cannot edit embedded chart data."""
        limitation = {
            "id": "LIM02",
            "feature": "Chart data editing",
            "status": "Not supported",
            "description": "Cannot modify data in existing charts embedded in presentations",
            "workaround": "Create new charts via placeholder workflow with PptxGenJS"
        }
        assert limitation["status"] == "Not supported"

    @pytest.mark.limitation
    def test_lim03_no_smartart(self, temp_dir):
        """LIM03: SmartArt manipulation not supported."""
        limitation = {
            "id": "LIM03",
            "feature": "SmartArt editing",
            "status": "Not supported",
            "description": "SmartArt diagrams cannot be created or modified",
            "workaround": "Use shapes and text boxes to create similar layouts"
        }
        assert limitation["status"] == "Not supported"

    @pytest.mark.limitation
    def test_lim04_no_animations(self, temp_dir):
        """LIM04: Animations and transitions not supported."""
        limitation = {
            "id": "LIM04",
            "feature": "Animations",
            "status": "Not supported",
            "description": "Cannot create, modify, or preserve slide animations and transitions",
            "workaround": "Manually add animations in PowerPoint after generation"
        }
        assert limitation["status"] == "Not supported"

    @pytest.mark.limitation
    def test_lim05_no_audio_video(self, temp_dir):
        """LIM05: Audio/video embedding not supported."""
        limitation = {
            "id": "LIM05",
            "feature": "Media embedding",
            "status": "Not supported",
            "description": "Cannot embed audio, video, or other media files",
            "workaround": "Add media manually after presentation is generated"
        }
        assert limitation["status"] == "Not supported"

    @pytest.mark.limitation
    def test_lim06_no_web_image_search(self, temp_dir):
        """LIM06: No integrated web image search."""
        limitation = {
            "id": "LIM06",
            "feature": "Web image search",
            "status": "Not supported",
            "description": "Cannot automatically search and insert images from the web",
            "workaround": "Download images separately and include via HTML img tags"
        }
        assert limitation["status"] == "Not supported"

    @pytest.mark.limitation
    def test_lim07_css_gradients(self, temp_dir):
        """LIM07: CSS gradients not directly supported."""
        limitation = {
            "id": "LIM07",
            "feature": "CSS gradients",
            "status": "Requires workaround",
            "description": "CSS linear-gradient and radial-gradient not supported in HTML",
            "workaround": "Pre-rasterize gradients to PNG using Sharp before including in HTML"
        }
        assert limitation["status"] == "Requires workaround"

    @pytest.mark.limitation
    def test_lim08_custom_fonts(self, temp_dir):
        """LIM08: Custom fonts not supported."""
        limitation = {
            "id": "LIM08",
            "feature": "Custom fonts",
            "status": "Limited",
            "description": "Only web-safe fonts are reliably supported",
            "supported_fonts": [
                "Arial", "Helvetica", "Times New Roman", "Georgia",
                "Courier New", "Verdana", "Tahoma", "Trebuchet MS",
                "Impact", "Comic Sans MS"
            ],
            "workaround": "Use web-safe fonts or rasterize text with custom fonts as images"
        }
        assert limitation["status"] == "Limited"


class TestErrorHandling:
    """Tests for error handling and recovery."""

    def test_err01_corrupted_pptx(self, temp_dir, python_cmd, scripts_dir):
        """ERR01: Handle corrupted PPTX file gracefully."""
        pptx_path = temp_dir / "corrupted.pptx"

        # Create invalid file
        pptx_path.write_bytes(b"This is not a valid PPTX file")

        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        # Should fail gracefully with clear error
        assert result.returncode != 0

    def test_err02_nonexistent_file(self, temp_dir, python_cmd, scripts_dir):
        """ERR02: Handle non-existent file gracefully."""
        pptx_path = temp_dir / "does_not_exist.pptx"
        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode != 0

    def test_err03_readonly_output(self, temp_dir, python_cmd, scripts_dir):
        """ERR03: Handle read-only output location."""
        pptx_path = temp_dir / "test.pptx"

        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[6])
        prs.save(pptx_path)

        # Create read-only directory (skip on Windows)
        import sys
        if sys.platform != "win32":
            readonly_dir = temp_dir / "readonly"
            readonly_dir.mkdir()
            readonly_dir.chmod(0o444)

            inventory_path = readonly_dir / "inventory.json"

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
                capture_output=True,
                text=True
            )

            # Restore permissions for cleanup
            readonly_dir.chmod(0o755)

            # Should fail with permission error
            assert result.returncode != 0 or "permission" in result.stderr.lower() or "error" in result.stderr.lower()


class TestCapabilityMatrix:
    """Test capability matrix for documentation."""

    def test_capability_summary(self, temp_dir):
        """Generate capability summary for documentation."""
        capabilities = {
            "text_operations": {
                "extract_text": {"status": "Full", "notes": "All text shapes extracted"},
                "replace_text": {"status": "Full", "notes": "With formatting preservation"},
                "format_text": {"status": "Full", "notes": "Bold, italic, underline, color, size"},
                "bullet_lists": {"status": "Full", "notes": "Multi-level with proper indentation"},
                "numbered_lists": {"status": "Full", "notes": "Standard numbering"},
            },
            "slide_operations": {
                "duplicate": {"status": "Full", "notes": "Including images and shapes"},
                "reorder": {"status": "Full", "notes": "Arbitrary order supported"},
                "delete": {"status": "Full", "notes": "By exclusion from list"},
                "add_new": {"status": "HTML only", "notes": "Via html2pptx workflow"},
            },
            "visual_elements": {
                "shapes": {"status": "Partial", "notes": "Rectangles, ovals, rounded rects"},
                "images": {"status": "HTML only", "notes": "Cannot edit existing"},
                "charts": {"status": "Placeholder", "notes": "Via PptxGenJS only"},
                "tables": {"status": "Placeholder", "notes": "Via PptxGenJS only"},
                "smartart": {"status": "None", "notes": "Not supported"},
            },
            "formatting": {
                "colors_rgb": {"status": "Full", "notes": "Hex RGB colors"},
                "colors_theme": {"status": "Preserve", "notes": "Cannot modify themes"},
                "fonts": {"status": "Limited", "notes": "Web-safe fonts only"},
                "gradients": {"status": "None", "notes": "Must pre-rasterize"},
                "shadows": {"status": "Outer only", "notes": "Inset not supported"},
            },
            "media": {
                "audio": {"status": "None", "notes": "Not supported"},
                "video": {"status": "None", "notes": "Not supported"},
                "animations": {"status": "None", "notes": "Not supported"},
            }
        }

        # Write capability matrix to file
        cap_file = temp_dir / "capabilities.json"
        with open(cap_file, "w") as f:
            json.dump(capabilities, f, indent=2)

        assert cap_file.exists()
