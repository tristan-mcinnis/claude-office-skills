"""
Pytest configuration and fixtures for PPTX capability tests.
"""
import pytest
import os
import sys
import json
import shutil
import tempfile
from pathlib import Path

# Add project root to path for imports
PROJECT_ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(PROJECT_ROOT / "public" / "pptx" / "scripts"))
sys.path.insert(0, str(PROJECT_ROOT / "public" / "pptx" / "ooxml" / "scripts"))

FIXTURES_DIR = Path(__file__).parent / "fixtures"
RESULTS_DIR = Path(__file__).parent / "results"
TEMPLATES_DIR = Path(__file__).parent / "templates"
HTML_DIR = Path(__file__).parent / "html"


@pytest.fixture(scope="session")
def project_root():
    """Return the project root directory."""
    return PROJECT_ROOT


@pytest.fixture(scope="session")
def fixtures_dir():
    """Return the fixtures directory path."""
    FIXTURES_DIR.mkdir(exist_ok=True)
    return FIXTURES_DIR


@pytest.fixture(scope="session")
def results_dir():
    """Return the results directory path."""
    RESULTS_DIR.mkdir(exist_ok=True)
    return RESULTS_DIR


@pytest.fixture(scope="session")
def templates_dir():
    """Return the templates directory path."""
    TEMPLATES_DIR.mkdir(exist_ok=True)
    return TEMPLATES_DIR


@pytest.fixture(scope="session")
def html_dir():
    """Return the html test files directory path."""
    HTML_DIR.mkdir(exist_ok=True)
    return HTML_DIR


@pytest.fixture
def temp_dir():
    """Create a temporary directory for test outputs."""
    tmp = tempfile.mkdtemp(prefix="pptx_test_")
    yield Path(tmp)
    # Cleanup after test
    shutil.rmtree(tmp, ignore_errors=True)


@pytest.fixture
def python_cmd():
    """Return the path to the venv Python executable."""
    return str(PROJECT_ROOT / "venv" / "bin" / "python")


@pytest.fixture
def scripts_dir():
    """Return the path to the PPTX scripts directory."""
    return PROJECT_ROOT / "public" / "pptx" / "scripts"


@pytest.fixture
def ooxml_scripts_dir():
    """Return the path to the OOXML scripts directory."""
    return PROJECT_ROOT / "public" / "pptx" / "ooxml" / "scripts"


# --- Test Data Fixtures ---

@pytest.fixture
def simple_text_content():
    """Simple text content for basic tests."""
    return {
        "title": "Test Presentation Title",
        "subtitle": "A subtitle for testing",
        "body": "This is body text for testing purposes.",
        "bullets": [
            "First bullet point",
            "Second bullet point",
            "Third bullet point"
        ]
    }


@pytest.fixture
def complex_text_content():
    """Complex text content with formatting."""
    return {
        "title": "Quarterly Business Review",
        "subtitle": "Q4 2024 Performance Analysis",
        "sections": [
            {
                "heading": "Key Highlights",
                "bullets": [
                    {"text": "Revenue grew 15% YoY", "bold": True},
                    {"text": "Customer satisfaction at 92%", "color": "00AA00"},
                    {"text": "Market share increased to 28%"}
                ]
            },
            {
                "heading": "Challenges",
                "bullets": [
                    {"text": "Supply chain disruptions", "level": 0},
                    {"text": "Increased raw material costs", "level": 1},
                    {"text": "Logistics delays", "level": 1},
                    {"text": "Competitive pressure", "level": 0}
                ]
            }
        ]
    }


@pytest.fixture
def unicode_text_content():
    """Unicode and special character content."""
    return {
        "languages": [
            {"lang": "Japanese", "text": "日本語テスト"},
            {"lang": "Chinese", "text": "中文测试"},
            {"lang": "Korean", "text": "한국어 테스트"},
            {"lang": "Arabic", "text": "اختبار عربي"},
            {"lang": "Russian", "text": "Русский тест"},
            {"lang": "Greek", "text": "Ελληνική δοκιμή"}
        ],
        "symbols": "© ® ™ € £ ¥ → ← ↑ ↓ • ◦ ▪ ★ ☆",
        "math": "∑ ∏ √ ∞ ≠ ≈ ≤ ≥ ± × ÷",
        "emoji": "Not supported - use images instead"
    }


@pytest.fixture
def html_slide_template():
    """Basic HTML slide template."""
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
            background: {background};
        }}
    </style>
</head>
<body>
{content}
</body>
</html>'''


@pytest.fixture
def replacement_json_template():
    """Template for replacement JSON structure."""
    return {
        "slide-0": {
            "shape-0": {
                "paragraphs": [
                    {"text": "Placeholder text", "alignment": "CENTER"}
                ]
            }
        }
    }


# --- Helper Functions ---

def create_test_pptx(output_path, num_slides=5, include_images=False, include_charts=False):
    """
    Create a simple test PPTX file programmatically.

    Args:
        output_path: Path to save the PPTX
        num_slides: Number of slides to create
        include_images: Whether to include image placeholders
        include_charts: Whether to include chart placeholders
    """
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Get blank layout
    blank_layout = prs.slide_layouts[6]

    for i in range(num_slides):
        slide = prs.slides.add_slide(blank_layout)

        # Add title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(1))
        title_frame = title_box.text_frame
        title_para = title_frame.paragraphs[0]
        title_para.text = f"Slide {i + 1} Title"
        title_para.font.size = Pt(44)
        title_para.font.bold = True
        title_para.alignment = PP_ALIGN.CENTER

        # Add body text
        body_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(12), Inches(4))
        body_frame = body_box.text_frame
        body_para = body_frame.paragraphs[0]
        body_para.text = f"This is body text for slide {i + 1}."
        body_para.font.size = Pt(24)

        # Add bullets
        for j in range(3):
            para = body_frame.add_paragraph()
            para.text = f"Bullet point {j + 1} on slide {i + 1}"
            para.font.size = Pt(20)
            para.level = 0

    prs.save(output_path)
    return output_path


@pytest.fixture
def create_simple_pptx(temp_dir):
    """Factory fixture to create simple test PPTX files."""
    def _create(name="test.pptx", num_slides=5, **kwargs):
        path = temp_dir / name
        return create_test_pptx(path, num_slides, **kwargs)
    return _create


# --- Markers ---

def pytest_configure(config):
    """Register custom markers."""
    config.addinivalue_line("markers", "slow: marks tests as slow (deselect with '-m \"not slow\"')")
    config.addinivalue_line("markers", "requires_libreoffice: marks tests that require LibreOffice")
    config.addinivalue_line("markers", "requires_node: marks tests that require Node.js")
    config.addinivalue_line("markers", "integration: marks integration tests")
    config.addinivalue_line("markers", "edge_case: marks edge case tests")
    config.addinivalue_line("markers", "limitation: marks tests for known limitations")


# --- Test Results Tracking ---

class TestResultsCollector:
    """Collect and summarize test results for capability reporting."""

    def __init__(self):
        self.results = {
            "passed": [],
            "failed": [],
            "skipped": [],
            "limitations": []
        }

    def add_result(self, test_id, status, description, notes=None):
        self.results[status].append({
            "test_id": test_id,
            "description": description,
            "notes": notes
        })

    def save_report(self, path):
        with open(path, "w") as f:
            json.dump(self.results, f, indent=2)


@pytest.fixture(scope="session")
def results_collector():
    """Session-scoped results collector."""
    return TestResultsCollector()
