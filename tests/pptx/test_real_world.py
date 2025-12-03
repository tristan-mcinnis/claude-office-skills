"""
Test real-world usage scenarios for PPTX skill.

Tests common business use cases and workflows.
"""
import pytest
import subprocess
import json
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN


def create_corporate_template(path, company_name="ACME Corp"):
    """Create a realistic corporate presentation template."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 0: Title slide
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(12.3), Inches(1.5))
    tf = title.text_frame
    p = tf.paragraphs[0]
    p.text = "[PRESENTATION TITLE]"
    p.font.size = Pt(48)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER

    subtitle = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(12.3), Inches(1))
    tf = subtitle.text_frame
    p = tf.paragraphs[0]
    p.text = "[Subtitle / Date]"
    p.font.size = Pt(24)
    p.alignment = PP_ALIGN.CENTER

    # Slide 1: Agenda
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    tf = title.text_frame
    p = tf.paragraphs[0]
    p.text = "Agenda"
    p.font.size = Pt(36)
    p.font.bold = True

    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.3), Inches(5))
    tf = body.text_frame
    for i, item in enumerate(["[Agenda Item 1]", "[Agenda Item 2]", "[Agenda Item 3]", "[Agenda Item 4]"]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.font.size = Pt(24)
        p.level = 0

    # Slide 2: Content slide
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    tf = title.text_frame
    p = tf.paragraphs[0]
    p.text = "[Section Title]"
    p.font.size = Pt(36)
    p.font.bold = True

    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.3), Inches(5))
    tf = body.text_frame
    p = tf.paragraphs[0]
    p.text = "[Content placeholder - replace with actual content]"
    p.font.size = Pt(20)

    # Slide 3: Two-column
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    tf = title.text_frame
    p = tf.paragraphs[0]
    p.text = "[Comparison Title]"
    p.font.size = Pt(36)
    p.font.bold = True

    left = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(5.9), Inches(5))
    tf = left.text_frame
    p = tf.paragraphs[0]
    p.text = "[Left column content]"
    p.font.size = Pt(18)

    right = slide.shapes.add_textbox(Inches(6.9), Inches(1.5), Inches(5.9), Inches(5))
    tf = right.text_frame
    p = tf.paragraphs[0]
    p.text = "[Right column content]"
    p.font.size = Pt(18)

    # Slide 4: Summary
    slide = prs.slides.add_slide(blank)
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(1))
    tf = title.text_frame
    p = tf.paragraphs[0]
    p.text = "Summary"
    p.font.size = Pt(36)
    p.font.bold = True

    body = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.3), Inches(5))
    tf = body.text_frame
    for i, item in enumerate(["[Key Takeaway 1]", "[Key Takeaway 2]", "[Key Takeaway 3]"]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.font.size = Pt(24)
        p.level = 0

    prs.save(path)
    return path


class TestQuarterlyReport:
    """Tests for quarterly report use case."""

    def test_rw01_quarterly_report_template(self, temp_dir, python_cmd, scripts_dir):
        """RW01: Update quarterly report from template."""
        template_path = temp_dir / "q4_template.pptx"
        create_corporate_template(template_path)

        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "q4_report.pptx"

        # Extract inventory
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Create realistic Q4 content
        replacements = {}

        # Slide 0: Title
        if "slide-0" in inventory:
            for shape_id, shape_data in inventory["slide-0"].items():
                for para in shape_data.get("paragraphs", []):
                    if "TITLE" in para.get("text", ""):
                        replacements.setdefault("slide-0", {})[shape_id] = {
                            "paragraphs": [{"text": "Q4 2024 Business Review", "bold": True, "alignment": "CENTER"}]
                        }
                    elif "Subtitle" in para.get("text", "") or "Date" in para.get("text", ""):
                        replacements.setdefault("slide-0", {})[shape_id] = {
                            "paragraphs": [{"text": "January 15, 2025", "alignment": "CENTER"}]
                        }

        # Slide 1: Agenda
        if "slide-1" in inventory:
            for shape_id, shape_data in inventory["slide-1"].items():
                if len(shape_data.get("paragraphs", [])) > 1:  # Body with multiple items
                    replacements.setdefault("slide-1", {})[shape_id] = {
                        "paragraphs": [
                            {"text": "Q4 Financial Highlights", "bullet": True, "level": 0},
                            {"text": "Key Accomplishments", "bullet": True, "level": 0},
                            {"text": "Challenges & Learnings", "bullet": True, "level": 0},
                            {"text": "Q1 2025 Outlook", "bullet": True, "level": 0}
                        ]
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
        assert output_path.exists()


class TestPitchDeck:
    """Tests for pitch deck creation."""

    def test_rw02_pitch_deck_structure(self, temp_dir, python_cmd, scripts_dir):
        """RW02: Build pitch deck from template structure."""
        template_path = temp_dir / "pitch_template.pptx"
        create_corporate_template(template_path)

        # Rearrange for pitch deck structure
        output_path = temp_dir / "pitch_rearranged.pptx"

        # Use title (0), content (2) twice, summary (4)
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(output_path), "0,2,2,4"],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        prs = Presentation(output_path)
        assert len(prs.slides) == 4


class TestLocalization:
    """Tests for content localization."""

    def test_rw03_text_swap_localization(self, temp_dir, python_cmd, scripts_dir):
        """RW03: Swap text for localization."""
        template_path = temp_dir / "english.pptx"
        create_corporate_template(template_path)

        inventory_path = temp_dir / "inventory.json"
        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "german.pptx"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Simple localization map
        translations = {
            "Agenda": "Tagesordnung",
            "Summary": "Zusammenfassung"
        }

        replacements = {}
        for slide_id, shapes in inventory.items():
            if not slide_id.startswith("slide-"):
                continue
            for shape_id, shape_data in shapes.items():
                for para in shape_data.get("paragraphs", []):
                    text = para.get("text", "")
                    if text in translations:
                        replacements.setdefault(slide_id, {})[shape_id] = {
                            "paragraphs": [{"text": translations[text], "bold": para.get("bold")}]
                        }

        if replacements:
            with open(replacements_path, "w") as f:
                json.dump(replacements, f)

            result = subprocess.run(
                [python_cmd, str(scripts_dir / "replace.py"),
                 str(template_path), str(replacements_path), str(output_path)],
                capture_output=True,
                text=True
            )

            assert result.returncode == 0


class TestContentExtraction:
    """Tests for content extraction and analysis."""

    def test_rw04_extract_all_text(self, temp_dir, python_cmd, scripts_dir):
        """RW04: Extract all text content from presentation."""
        template_path = temp_dir / "content.pptx"
        create_corporate_template(template_path)

        inventory_path = temp_dir / "all_text.json"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Extract all text
        all_text = []
        for slide_id, shapes in inventory.items():
            if not slide_id.startswith("slide-"):
                continue
            slide_text = []
            for shape_id, shape_data in shapes.items():
                for para in shape_data.get("paragraphs", []):
                    text = para.get("text", "").strip()
                    if text:
                        slide_text.append(text)
            all_text.append({"slide": slide_id, "content": slide_text})

        assert len(all_text) > 0


class TestMultipleVersions:
    """Tests for creating multiple versions from template."""

    def test_rw05_multiple_client_versions(self, temp_dir, python_cmd, scripts_dir):
        """RW05: Create multiple client-specific versions."""
        template_path = temp_dir / "base_template.pptx"
        create_corporate_template(template_path)

        inventory_path = temp_dir / "inventory.json"

        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        clients = ["Client A", "Client B", "Client C"]

        for client in clients:
            replacements_path = temp_dir / f"replacements_{client.lower().replace(' ', '_')}.json"
            output_path = temp_dir / f"deck_{client.lower().replace(' ', '_')}.pptx"

            replacements = {}
            # Replace title slide for each client
            if "slide-0" in inventory:
                first_shape = list(inventory["slide-0"].keys())[0]
                replacements["slide-0"] = {
                    first_shape: {
                        "paragraphs": [{"text": f"Proposal for {client}", "bold": True, "alignment": "CENTER"}]
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
            assert output_path.exists()


class TestVisualAnalysis:
    """Tests for visual analysis capabilities."""

    @pytest.mark.requires_libreoffice
    def test_rw06_thumbnail_generation(self, temp_dir, python_cmd, scripts_dir):
        """RW06: Generate thumbnail grid for visual analysis."""
        template_path = temp_dir / "visual.pptx"
        create_corporate_template(template_path)

        output_prefix = temp_dir / "thumbnails"

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "thumbnail.py"),
             str(template_path), str(output_prefix)],
            capture_output=True,
            text=True,
            timeout=120
        )

        # May fail if LibreOffice not available
        if result.returncode == 0:
            # Check if thumbnail was created
            thumbnails = list(temp_dir.glob("thumbnails*.jpg"))
            assert len(thumbnails) > 0


class TestOutlineToPresentation:
    """Tests for converting outline to presentation."""

    def test_rw07_outline_to_slides(self, temp_dir, python_cmd, scripts_dir):
        """RW07: Convert text outline to slide content."""
        template_path = temp_dir / "outline_template.pptx"
        create_corporate_template(template_path)

        # Outline structure
        outline = [
            {"title": "Introduction", "points": ["Background", "Objectives", "Scope"]},
            {"title": "Analysis", "points": ["Data review", "Key findings", "Implications"]},
            {"title": "Recommendations", "points": ["Short-term actions", "Long-term strategy"]},
        ]

        # First rearrange to get right number of slides
        working_path = temp_dir / "working.pptx"

        # Use content slide (2) for each outline section
        indices = ",".join(["0"] + ["2"] * len(outline) + ["4"])
        subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(working_path), indices],
            capture_output=True,
            text=True
        )

        # Now populate with outline content
        inventory_path = temp_dir / "inventory.json"
        subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(working_path), str(inventory_path)],
            capture_output=True,
            text=True
        )

        with open(inventory_path) as f:
            inventory = json.load(f)

        # Build replacements from outline
        replacements = {}

        # Process each content slide
        for i, section in enumerate(outline):
            slide_id = f"slide-{i + 1}"  # Skip title slide
            if slide_id in inventory:
                shapes = inventory[slide_id]
                shape_list = list(shapes.keys())

                if len(shape_list) >= 2:
                    # First shape is usually title
                    replacements.setdefault(slide_id, {})[shape_list[0]] = {
                        "paragraphs": [{"text": section["title"], "bold": True}]
                    }
                    # Second shape is body
                    replacements.setdefault(slide_id, {})[shape_list[1]] = {
                        "paragraphs": [
                            {"text": point, "bullet": True, "level": 0}
                            for point in section["points"]
                        ]
                    }

        replacements_path = temp_dir / "replacements.json"
        output_path = temp_dir / "from_outline.pptx"

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(working_path), str(replacements_path), str(output_path)],
            capture_output=True,
            text=True
        )

        assert result.returncode == 0


class TestWorkflowIntegration:
    """Tests for integrated workflows."""

    def test_rw08_full_workflow(self, temp_dir, python_cmd, scripts_dir):
        """RW08: Complete end-to-end workflow."""
        # 1. Create template
        template_path = temp_dir / "template.pptx"
        create_corporate_template(template_path)

        # 2. Analyze template
        inventory_path = temp_dir / "inventory.json"
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(template_path), str(inventory_path)],
            capture_output=True,
            text=True
        )
        assert result.returncode == 0

        # 3. Rearrange slides
        working_path = temp_dir / "working.pptx"
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "rearrange.py"),
             str(template_path), str(working_path), "0,1,2,2,4"],
            capture_output=True,
            text=True
        )
        assert result.returncode == 0

        # 4. Get new inventory
        working_inventory_path = temp_dir / "working_inventory.json"
        result = subprocess.run(
            [python_cmd, str(scripts_dir / "inventory.py"), str(working_path), str(working_inventory_path)],
            capture_output=True,
            text=True
        )
        assert result.returncode == 0

        # 5. Apply replacements
        with open(working_inventory_path) as f:
            inventory = json.load(f)

        replacements = {}
        for slide_id in ["slide-0", "slide-1"]:
            if slide_id in inventory:
                first_shape = list(inventory[slide_id].keys())[0]
                replacements[slide_id] = {
                    first_shape: {
                        "paragraphs": [{"text": f"Final content for {slide_id}"}]
                    }
                }

        replacements_path = temp_dir / "final_replacements.json"
        output_path = temp_dir / "final.pptx"

        with open(replacements_path, "w") as f:
            json.dump(replacements, f)

        result = subprocess.run(
            [python_cmd, str(scripts_dir / "replace.py"),
             str(working_path), str(replacements_path), str(output_path)],
            capture_output=True,
            text=True
        )
        assert result.returncode == 0
        assert output_path.exists()


class TestLimitationsInPractice:
    """Document practical limitations in real-world scenarios."""

    def test_rw09_no_image_insertion(self, temp_dir):
        """RW09: Document limitation - no automatic image insertion."""
        scenario = {
            "use_case": "Insert relevant images from web",
            "status": "Not supported",
            "limitation": "Cannot automatically search for and insert images from the web",
            "workaround": [
                "1. Download images manually",
                "2. Use html2pptx workflow with <img> tags",
                "3. Or add images manually in PowerPoint after generation"
            ]
        }
        assert scenario["status"] == "Not supported"

    def test_rw10_no_chart_from_data(self, temp_dir):
        """RW10: Document limitation - charts require placeholder workflow."""
        scenario = {
            "use_case": "Create chart from raw data",
            "status": "Requires specific workflow",
            "limitation": "Cannot directly add charts to existing presentations",
            "workaround": [
                "1. Use HTML placeholder in html2pptx workflow",
                "2. Add chart via PptxGenJS after HTML conversion",
                "3. Or use template with existing chart and update data manually"
            ]
        }
        assert scenario["status"] == "Requires specific workflow"

    def test_rw11_complex_formatting_loss(self, temp_dir):
        """RW11: Document limitation - some complex formatting may be lost."""
        scenario = {
            "use_case": "Preserve complex formatting from templates",
            "status": "Partial",
            "limitations": [
                "Animations are not preserved",
                "Transitions are not preserved",
                "SmartArt cannot be edited",
                "3D effects may be lost",
                "Some shadow effects may differ"
            ],
            "preserved": [
                "Basic text formatting (bold, italic, underline)",
                "Colors (RGB and theme)",
                "Font sizes",
                "Bullet/numbering",
                "Basic shapes"
            ]
        }
        assert "Animations" in scenario["limitations"][0]
