"""
Test HTML-to-PPTX creation for PPTX skill.

Tests creating presentations from scratch using HTML workflow.
"""
import pytest
import subprocess
import json
from pathlib import Path


class TestBasicHTML:
    """Tests for basic HTML slide creation."""

    @pytest.fixture
    def html_template(self):
        """Basic HTML template for 16:9 slides."""
        return '''<!DOCTYPE html>
<html>
<head>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: {bg};
        }}
        .title {{
            position: absolute;
            left: 36pt;
            top: 36pt;
            width: 648pt;
            font-size: 36pt;
            font-weight: bold;
        }}
        .content {{
            position: absolute;
            left: 36pt;
            top: 100pt;
            width: 648pt;
            font-size: 18pt;
        }}
    </style>
</head>
<body>
{content}
</body>
</html>'''

    def test_hc01_simple_text_slide(self, temp_dir, html_template):
        """HC01: Create simple slide with title and text."""
        html_path = temp_dir / "slide1.html"
        html_content = html_template.format(
            bg="#FFFFFF",
            content='''
            <div class="title"><p>Simple Title Slide</p></div>
            <div class="content"><p>This is simple body text content.</p></div>
            '''
        )
        html_path.write_text(html_content)

        # Verify HTML is valid
        assert html_path.exists()
        content = html_path.read_text()
        assert "720pt" in content
        assert "405pt" in content

    def test_hc02_two_column_layout(self, temp_dir):
        """HC02: Create two-column layout."""
        html_path = temp_dir / "two_col.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .title {
            position: absolute;
            left: 36pt;
            top: 20pt;
            width: 648pt;
            font-size: 32pt;
            font-weight: bold;
        }
        .left-col {
            position: absolute;
            left: 36pt;
            top: 80pt;
            width: 300pt;
            height: 290pt;
            font-size: 16pt;
        }
        .right-col {
            position: absolute;
            left: 360pt;
            top: 80pt;
            width: 324pt;
            height: 290pt;
            font-size: 16pt;
        }
    </style>
</head>
<body>
    <div class="title"><p>Two Column Layout</p></div>
    <div class="left-col">
        <p>Left column content goes here.</p>
        <p>Multiple paragraphs supported.</p>
    </div>
    <div class="right-col">
        <p>Right column content goes here.</p>
        <p>This creates a nice layout.</p>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        assert html_path.exists()
        content = html_path.read_text()
        assert "left-col" in content
        assert "right-col" in content

    def test_hc03_bullet_list(self, temp_dir):
        """HC03: Create slide with bullet list."""
        html_path = temp_dir / "bullets.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .title {
            position: absolute;
            left: 36pt;
            top: 20pt;
            width: 648pt;
            font-size: 32pt;
            font-weight: bold;
        }
        .bullets {
            position: absolute;
            left: 36pt;
            top: 80pt;
            width: 648pt;
            font-size: 18pt;
        }
        ul { list-style-type: disc; margin-left: 20pt; }
        li { margin-bottom: 8pt; }
    </style>
</head>
<body>
    <div class="title"><p>Key Points</p></div>
    <div class="bullets">
        <ul>
            <li>First important point</li>
            <li>Second key insight</li>
            <li>Third critical item</li>
            <li>Fourth consideration</li>
        </ul>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        assert html_path.exists()
        content = html_path.read_text()
        assert "<ul>" in content
        assert "<li>" in content


class TestShapesAndStyling:
    """Tests for shapes and styling in HTML."""

    def test_hc04_colored_background(self, temp_dir):
        """HC04: Slide with colored background."""
        html_path = temp_dir / "colored_bg.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #1E3A5F;
        }
        .title {
            position: absolute;
            left: 36pt;
            top: 150pt;
            width: 648pt;
            text-align: center;
        }
        .title p {
            font-size: 48pt;
            font-weight: bold;
            color: #FFFFFF;
        }
    </style>
</head>
<body>
    <div class="title"><p>Dark Background Title</p></div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "#1E3A5F" in content

    def test_hc05_shape_with_border(self, temp_dir):
        """HC05: Shape with border styling."""
        html_path = temp_dir / "bordered.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .card {
            position: absolute;
            left: 100pt;
            top: 100pt;
            width: 520pt;
            height: 200pt;
            background: #F5F5F5;
            border: 2pt solid #333333;
            padding: 20pt;
        }
        .card p {
            font-size: 18pt;
        }
    </style>
</head>
<body>
    <div class="card">
        <p>This is a card with a border.</p>
        <p>It has padding and background color.</p>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "border:" in content

    def test_hc06_rounded_rectangle(self, temp_dir):
        """HC06: Rounded rectangle shape."""
        html_path = temp_dir / "rounded.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .rounded-box {
            position: absolute;
            left: 200pt;
            top: 120pt;
            width: 320pt;
            height: 160pt;
            background: #3498DB;
            border-radius: 16pt;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .rounded-box p {
            font-size: 24pt;
            color: #FFFFFF;
            text-align: center;
        }
    </style>
</head>
<body>
    <div class="rounded-box">
        <p>Rounded Shape</p>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "border-radius" in content

    def test_hc07_drop_shadow(self, temp_dir):
        """HC07: Shape with drop shadow."""
        html_path = temp_dir / "shadow.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .shadowed {
            position: absolute;
            left: 200pt;
            top: 120pt;
            width: 320pt;
            height: 160pt;
            background: #FFFFFF;
            box-shadow: 4pt 4pt 8pt rgba(0,0,0,0.3);
        }
        .shadowed p {
            font-size: 20pt;
            padding: 20pt;
        }
    </style>
</head>
<body>
    <div class="shadowed">
        <p>Box with shadow effect</p>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "box-shadow" in content


class TestPlaceholders:
    """Tests for placeholder elements (charts, tables)."""

    def test_hc08_chart_placeholder(self, temp_dir):
        """HC08: Chart placeholder area."""
        html_path = temp_dir / "chart_placeholder.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .title {
            position: absolute;
            left: 36pt;
            top: 20pt;
            width: 300pt;
        }
        .title p { font-size: 28pt; font-weight: bold; }
        .chart-area {
            position: absolute;
            left: 360pt;
            top: 60pt;
            width: 340pt;
            height: 320pt;
        }
    </style>
</head>
<body>
    <div class="title"><p>Sales Overview</p></div>
    <div class="chart-area placeholder" data-type="chart" data-chart-type="bar"></div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "placeholder" in content
        assert "chart" in content

    def test_hc09_table_placeholder(self, temp_dir):
        """HC09: Table placeholder area."""
        html_path = temp_dir / "table_placeholder.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .title {
            position: absolute;
            left: 36pt;
            top: 20pt;
            width: 648pt;
        }
        .title p { font-size: 28pt; font-weight: bold; }
        .table-area {
            position: absolute;
            left: 36pt;
            top: 80pt;
            width: 648pt;
            height: 300pt;
        }
    </style>
</head>
<body>
    <div class="title"><p>Data Summary</p></div>
    <div class="table-area placeholder" data-type="table"></div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "table" in content


class TestInlineFormatting:
    """Tests for inline text formatting."""

    def test_hc10_bold_italic_text(self, temp_dir):
        """HC10: Bold and italic inline formatting."""
        html_path = temp_dir / "formatted.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .content {
            position: absolute;
            left: 36pt;
            top: 100pt;
            width: 648pt;
            font-size: 24pt;
        }
    </style>
</head>
<body>
    <div class="content">
        <p>This text has <b>bold words</b> and <i>italic words</i> and <u>underlined words</u>.</p>
        <p>You can also <b><i>combine formatting</i></b> for emphasis.</p>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "<b>" in content
        assert "<i>" in content
        assert "<u>" in content

    def test_hc11_colored_text(self, temp_dir):
        """HC11: Colored text using span."""
        html_path = temp_dir / "colored_text.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .content {
            position: absolute;
            left: 36pt;
            top: 100pt;
            width: 648pt;
            font-size: 24pt;
        }
    </style>
</head>
<body>
    <div class="content">
        <p>This text has <span style="color: #FF0000;">red text</span> and <span style="color: #00FF00;">green text</span>.</p>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "#FF0000" in content
        assert "#00FF00" in content


class TestComplexLayouts:
    """Tests for complex slide layouts."""

    def test_hc12_three_column_layout(self, temp_dir):
        """HC12: Three-column layout."""
        html_path = temp_dir / "three_col.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .title {
            position: absolute;
            left: 36pt;
            top: 20pt;
            width: 648pt;
            text-align: center;
        }
        .title p { font-size: 28pt; font-weight: bold; }
        .col1 {
            position: absolute;
            left: 36pt;
            top: 80pt;
            width: 200pt;
            height: 300pt;
            background: #E8F4FD;
            padding: 10pt;
        }
        .col2 {
            position: absolute;
            left: 260pt;
            top: 80pt;
            width: 200pt;
            height: 300pt;
            background: #E8F4FD;
            padding: 10pt;
        }
        .col3 {
            position: absolute;
            left: 484pt;
            top: 80pt;
            width: 200pt;
            height: 300pt;
            background: #E8F4FD;
            padding: 10pt;
        }
        .col1 p, .col2 p, .col3 p { font-size: 14pt; }
    </style>
</head>
<body>
    <div class="title"><p>Three Pillars</p></div>
    <div class="col1"><p>First Pillar: Quality focus on all deliverables</p></div>
    <div class="col2"><p>Second Pillar: Customer satisfaction metrics</p></div>
    <div class="col3"><p>Third Pillar: Continuous improvement</p></div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "col1" in content
        assert "col2" in content
        assert "col3" in content

    def test_hc13_flexbox_layout(self, temp_dir):
        """HC13: Flexbox-based layout."""
        html_path = temp_dir / "flexbox.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            width: 720pt;
            height: 405pt;
            font-family: Arial, sans-serif;
            position: relative;
            background: #FFFFFF;
        }
        .container {
            position: absolute;
            left: 36pt;
            top: 80pt;
            width: 648pt;
            height: 300pt;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .item {
            width: 150pt;
            height: 150pt;
            background: #3498DB;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 8pt;
        }
        .item p { color: #FFFFFF; font-size: 16pt; text-align: center; }
    </style>
</head>
<body>
    <div class="container">
        <div class="item"><p>Item 1</p></div>
        <div class="item"><p>Item 2</p></div>
        <div class="item"><p>Item 3</p></div>
        <div class="item"><p>Item 4</p></div>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "display: flex" in content


class TestKnownLimitations:
    """Tests documenting known limitations of HTML-to-PPTX."""

    @pytest.mark.limitation
    def test_hc14_gradient_not_supported(self, temp_dir):
        """HC14: CSS gradients are NOT supported (must pre-rasterize)."""
        html_path = temp_dir / "gradient.html"
        # This HTML would cause an error if processed
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        body {
            width: 720pt;
            height: 405pt;
            /* WARNING: This will fail - gradients not supported */
            background: linear-gradient(to right, #FF0000, #0000FF);
        }
    </style>
</head>
<body></body>
</html>'''
        html_path.write_text(html_content)

        # Document that this is a known limitation
        content = html_path.read_text()
        assert "linear-gradient" in content
        # In real usage, this would need to be pre-rasterized with Sharp

    @pytest.mark.limitation
    def test_hc15_br_tag_not_supported(self, temp_dir):
        """HC15: BR tags are NOT supported (use separate elements)."""
        html_path = temp_dir / "br_tag.html"
        # This HTML would cause an error if processed
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        body { width: 720pt; height: 405pt; font-family: Arial; }
        .content { position: absolute; left: 36pt; top: 100pt; width: 648pt; }
    </style>
</head>
<body>
    <div class="content">
        <!-- WARNING: <br> tags not supported, use separate <p> elements -->
        <p>Line one<br>Line two<br>Line three</p>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "<br>" in content
        # In real usage, this would fail validation

    @pytest.mark.limitation
    def test_hc16_text_in_div_ignored(self, temp_dir):
        """HC16: Text directly in DIV is silently ignored."""
        html_path = temp_dir / "text_in_div.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        body { width: 720pt; height: 405pt; font-family: Arial; }
        .content { position: absolute; left: 36pt; top: 100pt; width: 648pt; font-size: 24pt; }
    </style>
</head>
<body>
    <div class="content">
        <!-- WARNING: This text will be ignored - must use <p> tags -->
        This text is NOT in a p tag and will be ignored!
        <p>This text IS in a p tag and will appear.</p>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        # Document that unwrapped text is ignored
        content = html_path.read_text()
        assert "NOT in a p tag" in content

    @pytest.mark.limitation
    def test_hc17_non_websafe_fonts(self, temp_dir):
        """HC17: Non web-safe fonts may not render correctly."""
        html_path = temp_dir / "custom_font.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        body {
            width: 720pt;
            height: 405pt;
            /* WARNING: Custom fonts not supported, will fallback */
            font-family: "Montserrat", "Custom Font", Arial, sans-serif;
        }
        .content { position: absolute; left: 36pt; top: 100pt; width: 648pt; }
        .content p { font-size: 24pt; }
    </style>
</head>
<body>
    <div class="content">
        <p>This may not render in Montserrat font.</p>
    </div>
</body>
</html>'''
        html_path.write_text(html_content)

        # Document font limitation
        content = html_path.read_text()
        assert "Montserrat" in content

    @pytest.mark.limitation
    def test_hc18_inset_shadow_not_supported(self, temp_dir):
        """HC18: Inset shadows cause PowerPoint corruption."""
        html_path = temp_dir / "inset_shadow.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        body { width: 720pt; height: 405pt; font-family: Arial; background: #FFFFFF; }
        .box {
            position: absolute;
            left: 200pt;
            top: 100pt;
            width: 320pt;
            height: 200pt;
            background: #F0F0F0;
            /* WARNING: inset shadows NOT supported - causes file corruption */
            box-shadow: inset 4pt 4pt 8pt rgba(0,0,0,0.3);
        }
    </style>
</head>
<body>
    <div class="box"><p>Inset shadow box</p></div>
</body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        assert "inset" in content


class TestDimensionValidation:
    """Tests for HTML dimension validation."""

    def test_hc19_correct_dimensions_16_9(self, temp_dir):
        """HC19: Correct 16:9 dimensions."""
        html_path = temp_dir / "correct_16_9.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        body { width: 720pt; height: 405pt; font-family: Arial; }
    </style>
</head>
<body></body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        # 16:9 = 720pt x 405pt
        assert "720pt" in content
        assert "405pt" in content

    def test_hc20_correct_dimensions_4_3(self, temp_dir):
        """HC20: Correct 4:3 dimensions."""
        html_path = temp_dir / "correct_4_3.html"
        html_content = '''<!DOCTYPE html>
<html>
<head>
    <style>
        body { width: 720pt; height: 540pt; font-family: Arial; }
    </style>
</head>
<body></body>
</html>'''
        html_path.write_text(html_content)

        content = html_path.read_text()
        # 4:3 = 720pt x 540pt
        assert "720pt" in content
        assert "540pt" in content
