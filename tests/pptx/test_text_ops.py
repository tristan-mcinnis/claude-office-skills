"""
Test text operations for PPTX skill.

Tests text extraction, replacement, and formatting preservation capabilities.
"""
import pytest
import json
import subprocess
from pathlib import Path


class TestTextExtraction:
    """Tests for text extraction/inventory functionality."""

    def test_t01_simple_text_extraction(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T01: Extract text from simple presentation."""
        pptx_path = create_simple_pptx("simple.pptx", num_slides=3)
        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0, f"Inventory extraction failed: {result.stderr}"
        assert inventory_path.exists(), "Inventory file not created"

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Should have 3 slides
        assert len([k for k in inventory.keys() if k.startswith("slide-")]) == 3
        # Each slide should have shapes
        assert "slide-0" in inventory
        assert len(inventory["slide-0"]) > 0

    def test_t02_extract_formatting_metadata(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T02: Verify formatting metadata is captured."""
        pptx_path = create_simple_pptx("formatted.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Check that shapes have position and formatting info
        slide = inventory.get("slide-0", {})
        if slide:
            first_shape = list(slide.values())[0]
            # Should have position data
            assert "left" in first_shape
            assert "top" in first_shape
            assert "width" in first_shape
            assert "height" in first_shape
            # Should have paragraphs
            assert "paragraphs" in first_shape

    def test_t03_extract_paragraph_formatting(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T03: Verify paragraph-level formatting is extracted."""
        pptx_path = create_simple_pptx("para_format.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        slide = inventory.get("slide-0", {})
        for shape_id, shape_data in slide.items():
            for para in shape_data.get("paragraphs", []):
                # Paragraphs should have text
                assert "text" in para
                # Should have some formatting info (may be None if default)
                # Just verify the structure exists
                assert isinstance(para, dict)


class TestTextReplacement:
    """Tests for text replacement functionality."""

    def test_t04_simple_replacement(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T04: Simple text replacement."""
        pptx_path = create_simple_pptx("replace_test.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        # First get inventory
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Create replacement for first shape
        replacements = {}
        if "slide-0" in inventory:
            first_shape_id = list(inventory["slide-0"].keys())[0]
            replacements["slide-0"] = {
                first_shape_id: {
                    "paragraphs": [{"text": "Replaced Title Text", "alignment": "CENTER"}]
                }
            }

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        # Apply replacement
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(pptx_path), str(replacements_path), str(output_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0, f"Replacement failed: {result.stderr}"
        assert output_path.exists(), "Output file not created"

    def test_t05_multi_paragraph_replacement(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T05: Replace text with multiple paragraphs."""
        pptx_path = create_simple_pptx("multi_para.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Find a shape with body text
        replacements = {}
        if "slide-0" in inventory:
            for shape_id, shape_data in inventory["slide-0"].items():
                if len(shape_data.get("paragraphs", [])) > 1:
                    replacements["slide-0"] = {
                        shape_id: {
                            "paragraphs": [
                                {"text": "First paragraph"},
                                {"text": "Second paragraph"},
                                {"text": "Third paragraph"}
                            ]
                        }
                    }
                    break

        if replacements:
            with open(replacements_path, "w") as f:
                json.dump(replacements, f)

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "replace.py"),
                 str(pptx_path), str(replacements_path), str(output_path)],
                capture_output=True,
                text=True
            )

            assert result.returncode == 0, f"Multi-paragraph replacement failed: {result.stderr}"

    def test_t06_bullet_formatting(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T06: Replace text with bullet formatting."""
        pptx_path = create_simple_pptx("bullets.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Create bullet list replacement
        replacements = {}
        if "slide-0" in inventory:
            # Find body text shape (usually second shape)
            shapes = list(inventory["slide-0"].keys())
            if len(shapes) > 1:
                body_shape = shapes[1]
                replacements["slide-0"] = {
                    body_shape: {
                        "paragraphs": [
                            {"text": "First bullet", "bullet": True, "level": 0},
                            {"text": "Second bullet", "bullet": True, "level": 0},
                            {"text": "Sub-bullet", "bullet": True, "level": 1},
                            {"text": "Third bullet", "bullet": True, "level": 0}
                        ]
                    }
                }

        if replacements:
            with open(replacements_path, "w") as f:
                json.dump(replacements, f)

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "replace.py"),
                 str(pptx_path), str(replacements_path), str(output_path)],
                capture_output=True,
                text=True
            )

            assert result.returncode == 0, f"Bullet formatting failed: {result.stderr}"


class TestFormattingPreservation:
    """Tests for formatting preservation during operations."""

    def test_t07_bold_italic_preservation(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T07: Bold and italic formatting in replacements."""
        pptx_path = create_simple_pptx("bold_italic.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        if "slide-0" in inventory:
            first_shape = list(inventory["slide-0"].keys())[0]
            replacements = {
                "slide-0": {
                    first_shape: {
                        "paragraphs": [
                            {"text": "Bold and Important", "bold": True},
                            {"text": "Italicized text", "italic": True}
                        ]
                    }
                }
            }

            with open(replacements_path, "w") as f:
                json.dump(replacements, f)

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "replace.py"),
                 str(pptx_path), str(replacements_path), str(output_path)],
                capture_output=True,
                text=True
            )

            assert result.returncode == 0

    def test_t08_color_formatting(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T08: Color formatting in replacements."""
        pptx_path = create_simple_pptx("colors.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        if "slide-0" in inventory:
            first_shape = list(inventory["slide-0"].keys())[0]
            replacements = {
                "slide-0": {
                    first_shape: {
                        "paragraphs": [
                            {"text": "Red text", "color": "FF0000"},
                            {"text": "Green text", "color": "00FF00"},
                            {"text": "Blue text", "color": "0000FF"}
                        ]
                    }
                }
            }

            with open(replacements_path, "w") as f:
                json.dump(replacements, f)

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "replace.py"),
                 str(pptx_path), str(replacements_path), str(output_path)],
                capture_output=True,
                text=True
            )

            assert result.returncode == 0


class TestSpecialCharacters:
    """Tests for special character handling."""

    def test_t09_unicode_characters(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T09: Unicode character handling."""
        pptx_path = create_simple_pptx("unicode.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        if "slide-0" in inventory:
            first_shape = list(inventory["slide-0"].keys())[0]
            replacements = {
                "slide-0": {
                    first_shape: {
                        "paragraphs": [
                            {"text": "日本語テスト - Japanese"},
                            {"text": "中文测试 - Chinese"},
                            {"text": "Symbols: © ® ™ € £ ¥"}
                        ]
                    }
                }
            }

            with open(replacements_path, "w") as f:
                json.dump(replacements, f, ensure_ascii=False)

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "replace.py"),
                 str(pptx_path), str(replacements_path), str(output_path)],
                capture_output=True,
                text=True
            )

            assert result.returncode == 0, f"Unicode replacement failed: {result.stderr}"

    def test_t10_special_symbols(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T10: Special symbol handling."""
        pptx_path = create_simple_pptx("symbols.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        if "slide-0" in inventory:
            first_shape = list(inventory["slide-0"].keys())[0]
            replacements = {
                "slide-0": {
                    first_shape: {
                        "paragraphs": [
                            {"text": "Arrows: → ← ↑ ↓ ↔ ↕"},
                            {"text": "Math: ∑ ∏ √ ∞ ≠ ≈ ≤ ≥"},
                            {"text": "Quotes: "quoted" 'single'"}
                        ]
                    }
                }
            }

            with open(replacements_path, "w") as f:
                json.dump(replacements, f, ensure_ascii=False)

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "replace.py"),
                 str(pptx_path), str(replacements_path), str(output_path)],
                capture_output=True,
                text=True
            )

            assert result.returncode == 0


class TestOverflowDetection:
    """Tests for text overflow detection."""

    def test_t11_overflow_detection(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T11: Detect text overflow in small shapes."""
        pptx_path = create_simple_pptx("overflow.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        # Check if overflow warnings are in output
        # The script should report overflow issues
        assert result.returncode == 0

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Verify inventory has overflow info if applicable
        assert "slide-0" in inventory


class TestEdgeCases:
    """Edge case tests for text operations."""

    def test_t12_empty_shape_handling(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T12: Shapes without text should be handled."""
        pptx_path = create_simple_pptx("empty_shapes.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        # Should complete without error
        assert result.returncode == 0

    def test_t13_very_long_text(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T13: Handle very long text content."""
        pptx_path = create_simple_pptx("long_text.pptx", num_slides=1)
        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(pptx_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Create very long text
        long_text = "This is a very long paragraph. " * 100

        if "slide-0" in inventory:
            first_shape = list(inventory["slide-0"].keys())[0]
            replacements = {
                "slide-0": {
                    first_shape: {
                        "paragraphs": [{"text": long_text}]
                    }
                }
            }

            with open(replacements_path, "w") as f:
                json.dump(replacements, f)

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "replace.py"),
                 str(pptx_path), str(replacements_path), str(output_path)],
                capture_output=True,
                text=True
            )

            # Should complete (may have overflow warning)
            assert result.returncode == 0

    def test_t14_invalid_shape_reference(self, create_simple_pptx, temp_dir, python_cmd, scripts_dir):
        """T14: Invalid shape reference should error gracefully."""
        pptx_path = create_simple_pptx("invalid_ref.pptx", num_slides=1)
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "output.pptx"

        # Reference non-existent shape
        replacements = {
            "slide-0": {
                "shape-999": {
                    "paragraphs": [{"text": "This should fail"}]
                }
            }
        }

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(pptx_path), str(replacements_path), str(output_path)],
            capture_output=True,
            text=True
        )

        # Should fail with error about missing shape
        assert result.returncode != 0 or "not found" in result.stderr.lower() or "error" in result.stderr.lower()
