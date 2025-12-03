"""
Test SKILL.md instruction variations and their impact on output quality.

This is the PRIMARY test file for understanding how different instruction
configurations in SKILL.md affect Claude Code's PowerPoint editing behavior.

The hypothesis: More specific, constraint-based instructions lead to better
layout matching and fewer empty placeholder issues.
"""
import pytest
import json
from pathlib import Path
from dataclasses import dataclass
from typing import List, Dict, Optional


# ============================================================================
# SKILL INSTRUCTION VARIATIONS
# ============================================================================
# These represent different versions of instructions that could go in SKILL.md
# Each variation tests a hypothesis about instruction effectiveness

@dataclass
class SkillVariation:
    """Represents a variation of SKILL.md instructions."""
    name: str
    description: str
    instructions: str
    expected_impact: str


# Baseline: Current minimal instructions
VARIATION_BASELINE = SkillVariation(
    name="baseline",
    description="Minimal instructions, maximum flexibility",
    instructions="""
## Template-Based Workflow

1. Analyze template slides
2. Map content to appropriate slides
3. Apply replacements
    """,
    expected_impact="May select wrong layouts, leave empty placeholders"
)

# Enhanced: Add layout matching rules
VARIATION_ENHANCED = SkillVariation(
    name="enhanced",
    description="Added layout matching constraints",
    instructions="""
## Template-Based Workflow

1. Analyze template slides and their layout structures
2. Map content to appropriate slides

### Layout Matching Rules
- Single-column layouts: Use for unified narrative or single topic
- Two-column layouts: Use ONLY when you have exactly 2 distinct items
- Three-column layouts: Use ONLY when you have exactly 3 distinct items
- Never use layouts with more placeholders than you have content

3. Apply replacements ensuring all placeholders are filled
    """,
    expected_impact="Better layout matching, fewer empty placeholders"
)

# Strict: Add forbidden patterns and decision tree
VARIATION_STRICT = SkillVariation(
    name="strict",
    description="Explicit forbidden patterns and decision tree",
    instructions="""
## Template-Based Workflow

1. Analyze template slides and count placeholders per slide
2. Count your content items BEFORE selecting a layout

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

3. Apply replacements - verify ALL shapes have content
    """,
    expected_impact="Highest constraint satisfaction, may be more rigid"
)

# Example-driven: Show correct and incorrect examples
VARIATION_EXAMPLES = SkillVariation(
    name="examples",
    description="Examples of correct and incorrect layout choices",
    instructions="""
## Template-Based Workflow

1. Analyze template slides
2. Match content to layouts using these examples:

### Correct Layout Choices
- 2 benefits to compare → 2-column layout ✓
- 3 strategic pillars → 3-column layout ✓
- 5 bullet points → single-column bullet list ✓
- CEO quote with name → quote layout ✓

### Incorrect Layout Choices
- 2 items in 3-column layout ✗ (empty column)
- 4 items squeezed into 3-column ✗ (overflow)
- Image layout without image ✗ (empty placeholder)
- Quote layout for emphasis text ✗ (no attribution)

3. Apply replacements
    """,
    expected_impact="Clear mental models, easier to follow patterns"
)

# Checklist: Pre-flight checklist before applying
VARIATION_CHECKLIST = SkillVariation(
    name="checklist",
    description="Pre-application checklist to verify choices",
    instructions="""
## Template-Based Workflow

1. Analyze template slides
2. Map content to slides

### Pre-Application Checklist (VERIFY BEFORE PROCEEDING)
□ Content item count matches placeholder count?
□ No image placeholders without images?
□ Quote layouts only for attributed quotes?
□ All placeholders will have content?
□ No content overflow expected?

If ANY checkbox fails → choose different layout

3. Apply replacements only after checklist passes
    """,
    expected_impact="Forces explicit verification step"
)

ALL_VARIATIONS = [
    VARIATION_BASELINE,
    VARIATION_ENHANCED,
    VARIATION_STRICT,
    VARIATION_EXAMPLES,
    VARIATION_CHECKLIST
]


# ============================================================================
# TEST SCENARIOS
# ============================================================================
# Each scenario tests a specific layout decision challenge

@dataclass
class TestScenario:
    """A test scenario for evaluating skill instruction effectiveness."""
    id: str
    name: str
    description: str
    content_items: int
    content_type: str
    available_layouts: List[str]
    correct_choice: str
    incorrect_choices: List[str]
    why_matters: str


SCENARIOS = [
    TestScenario(
        id="S01",
        name="two_items_comparison",
        description="User has exactly 2 items to compare",
        content_items=2,
        content_type="comparison",
        available_layouts=["1-column", "2-column", "3-column"],
        correct_choice="2-column",
        incorrect_choices=["3-column (empty column)"],
        why_matters="3-column leaves visible empty space, looks unprofessional"
    ),
    TestScenario(
        id="S02",
        name="three_items_pillars",
        description="User has exactly 3 strategic pillars",
        content_items=3,
        content_type="pillars",
        available_layouts=["1-column", "2-column", "3-column"],
        correct_choice="3-column",
        incorrect_choices=["2-column (content overflow)"],
        why_matters="2-column can't fit 3 equal items, loses visual hierarchy"
    ),
    TestScenario(
        id="S03",
        name="five_items_list",
        description="User has 5 bullet points",
        content_items=5,
        content_type="bullet_list",
        available_layouts=["1-column", "2-column", "3-column"],
        correct_choice="1-column (bullet list)",
        incorrect_choices=["3-column (cramming)", "2-column (uneven split)"],
        why_matters="Forcing 5 items into columns creates cramped, hard-to-read slides"
    ),
    TestScenario(
        id="S04",
        name="no_image_available",
        description="Text content only, no images to insert",
        content_items=1,
        content_type="text_description",
        available_layouts=["text-only", "image+text"],
        correct_choice="text-only",
        incorrect_choices=["image+text (empty image area)"],
        why_matters="Empty image placeholder looks like a mistake"
    ),
    TestScenario(
        id="S05",
        name="emphasis_not_quote",
        description="Key message for emphasis (not a real quote)",
        content_items=1,
        content_type="emphasis",
        available_layouts=["content", "quote"],
        correct_choice="content",
        incorrect_choices=["quote (no attribution)"],
        why_matters="Quote layouts expect speaker name, missing attribution looks wrong"
    ),
    TestScenario(
        id="S06",
        name="real_quote_with_speaker",
        description="Actual quote from a person with attribution",
        content_items=1,
        content_type="quote",
        available_layouts=["content", "quote"],
        correct_choice="quote",
        incorrect_choices=["content (loses visual impact)"],
        why_matters="Quote formatting adds credibility and visual distinction"
    ),
    TestScenario(
        id="S07",
        name="four_items_awkward",
        description="4 items - doesn't fit cleanly into 2 or 3 columns",
        content_items=4,
        content_type="features",
        available_layouts=["1-column", "2-column", "3-column"],
        correct_choice="1-column OR 2x2 grid OR 2 slides",
        incorrect_choices=["3-column (one item orphaned)", "2-column (2 items cramped per column)"],
        why_matters="4 items is the trickiest - needs creative solution"
    ),
    TestScenario(
        id="S08",
        name="single_key_message",
        description="One powerful statement to emphasize",
        content_items=1,
        content_type="key_message",
        available_layouts=["content", "title-only", "full-bleed"],
        correct_choice="title-only or full-bleed",
        incorrect_choices=["content with empty body area"],
        why_matters="Single message deserves visual prominence, not empty body"
    ),
]


# ============================================================================
# TEST CLASSES
# ============================================================================

class TestSkillVariationDefinitions:
    """Verify skill variations are properly defined."""

    def test_sv01_all_variations_have_instructions(self):
        """SV01: All variations must have non-empty instructions."""
        for var in ALL_VARIATIONS:
            assert len(var.instructions.strip()) > 50, f"{var.name} has insufficient instructions"
            assert var.description, f"{var.name} missing description"
            assert var.expected_impact, f"{var.name} missing expected impact"

    def test_sv02_variations_differ_meaningfully(self):
        """SV02: Variations should be meaningfully different."""
        instruction_sets = [var.instructions for var in ALL_VARIATIONS]
        # All should be unique
        assert len(set(instruction_sets)) == len(instruction_sets), "Duplicate variations found"

    def test_sv03_baseline_is_minimal(self):
        """SV03: Baseline should be minimal (fewest constraints)."""
        baseline_len = len(VARIATION_BASELINE.instructions)
        for var in ALL_VARIATIONS:
            if var.name != "baseline":
                assert len(var.instructions) > baseline_len, \
                    f"{var.name} should have more instructions than baseline"


class TestScenarioCoverage:
    """Verify test scenarios cover key decision points."""

    def test_sc01_scenarios_cover_all_content_counts(self):
        """SC01: Scenarios should cover 1, 2, 3, 4, 5+ item counts."""
        counts = set(s.content_items for s in SCENARIOS)
        assert 1 in counts, "Missing single-item scenario"
        assert 2 in counts, "Missing two-item scenario"
        assert 3 in counts, "Missing three-item scenario"
        assert 4 in counts, "Missing four-item scenario"
        assert any(c >= 5 for c in counts), "Missing 5+ item scenario"

    def test_sc02_scenarios_have_clear_correct_choice(self):
        """SC02: Each scenario must have a clear correct choice."""
        for scenario in SCENARIOS:
            assert scenario.correct_choice, f"{scenario.id} missing correct choice"
            assert scenario.incorrect_choices, f"{scenario.id} missing incorrect choices"
            assert scenario.why_matters, f"{scenario.id} missing explanation"

    def test_sc03_scenarios_cover_content_types(self):
        """SC03: Scenarios should cover various content types."""
        content_types = set(s.content_type for s in SCENARIOS)
        expected_types = {"comparison", "bullet_list", "quote", "emphasis"}
        covered = content_types & expected_types
        assert len(covered) >= 3, f"Need more content type coverage, only have: {content_types}"


class TestVariationEffectiveness:
    """Test expected effectiveness of each variation."""

    @pytest.mark.parametrize("scenario", SCENARIOS, ids=lambda s: s.id)
    def test_ve01_scenario_analysis(self, scenario):
        """VE01: Analyze each scenario for instruction requirements."""
        analysis = {
            "scenario_id": scenario.id,
            "content_count": scenario.content_items,
            "baseline_likely_success": scenario.content_items <= 2,  # Simple cases
            "needs_explicit_constraint": scenario.content_items >= 4 or scenario.content_type in ["quote", "emphasis"],
            "key_instruction_needed": None
        }

        # Determine what instruction is needed
        if scenario.content_items == 2 and "3-column" in str(scenario.incorrect_choices):
            analysis["key_instruction_needed"] = "2 items → 2-column ONLY"
        elif scenario.content_items >= 4:
            analysis["key_instruction_needed"] = "4+ items → bullet list or multiple slides"
        elif "image" in scenario.content_type.lower() or "image" in str(scenario.incorrect_choices):
            analysis["key_instruction_needed"] = "image layout requires actual image"
        elif scenario.content_type == "emphasis":
            analysis["key_instruction_needed"] = "quote layout requires attribution"

        # This test documents the analysis
        assert analysis["scenario_id"] == scenario.id

    def test_ve02_variation_constraint_coverage(self):
        """VE02: Check which constraints each variation covers."""
        constraints_to_check = [
            "2 items → 2-column only",
            "3 items → 3-column only",
            "no empty placeholders",
            "image layout requires image",
            "quote layout requires attribution",
            "4+ items special handling"
        ]

        coverage = {}
        for var in ALL_VARIATIONS:
            coverage[var.name] = {
                constraint: constraint.lower().split()[0] in var.instructions.lower()
                for constraint in constraints_to_check
            }

        # Enhanced and above should cover more than baseline
        baseline_coverage = sum(coverage["baseline"].values())
        for var in ["enhanced", "strict", "examples", "checklist"]:
            var_coverage = sum(coverage[var].values())
            assert var_coverage >= baseline_coverage, \
                f"{var} should have more constraint coverage than baseline"


class TestExpectedBehavior:
    """Document expected Claude Code behavior per variation."""

    def test_eb01_baseline_expected_issues(self):
        """EB01: Document expected issues with baseline instructions."""
        expected_issues = {
            "variation": "baseline",
            "likely_issues": [
                "May use 3-column for 2 items (empty column)",
                "May use image layout without image",
                "May use quote layout for emphasis text",
                "May cram 5+ items into 3-column"
            ],
            "success_cases": [
                "Simple 1-item content",
                "Clear 3-item content matching 3-column",
                "Straightforward replacements"
            ]
        }
        assert len(expected_issues["likely_issues"]) > 0

    def test_eb02_enhanced_expected_behavior(self):
        """EB02: Document expected behavior with enhanced instructions."""
        expected_behavior = {
            "variation": "enhanced",
            "improvements_over_baseline": [
                "Should match 2 items to 2-column",
                "Should match 3 items to 3-column",
                "Should avoid overfilling layouts"
            ],
            "remaining_issues": [
                "May still miss image/quote edge cases",
                "4-item case still ambiguous"
            ]
        }
        assert len(expected_behavior["improvements_over_baseline"]) > 0

    def test_eb03_strict_expected_behavior(self):
        """EB03: Document expected behavior with strict instructions."""
        expected_behavior = {
            "variation": "strict",
            "improvements": [
                "Explicit forbidden patterns should prevent common errors",
                "Decision tree provides clear fallback logic",
                "4+ items explicitly handled"
            ],
            "potential_downsides": [
                "May be too rigid for creative layouts",
                "More instructions = more tokens",
                "Could conflict with user preferences"
            ]
        }
        assert len(expected_behavior["improvements"]) > 0


class TestInstructionMetrics:
    """Metrics for comparing instruction variations."""

    def test_im01_instruction_length_comparison(self):
        """IM01: Compare instruction lengths (token proxy)."""
        lengths = {var.name: len(var.instructions) for var in ALL_VARIATIONS}

        # Document lengths
        assert lengths["baseline"] < lengths["enhanced"]
        assert lengths["enhanced"] < lengths["strict"]

    def test_im02_constraint_density(self):
        """IM02: Measure constraint density (constraints per 100 chars)."""
        constraint_keywords = ["ONLY", "NEVER", "must", "exactly", "always", "forbidden"]

        densities = {}
        for var in ALL_VARIATIONS:
            text = var.instructions.lower()
            count = sum(text.count(kw.lower()) for kw in constraint_keywords)
            density = count / (len(var.instructions) / 100)
            densities[var.name] = round(density, 2)

        # Strict should have highest density
        assert densities["strict"] >= densities["baseline"]

    def test_im03_example_count(self):
        """IM03: Count concrete examples in instructions."""
        example_indicators = ["→", "✓", "✗", "e.g.", "example", "for instance"]

        example_counts = {}
        for var in ALL_VARIATIONS:
            text = var.instructions.lower()
            count = sum(text.count(ind.lower()) for ind in example_indicators)
            example_counts[var.name] = count

        # Examples variation should have most examples
        assert example_counts["examples"] >= example_counts["baseline"]


class TestRecommendations:
    """Generate recommendations based on test analysis."""

    def test_rec01_variation_recommendation_matrix(self):
        """REC01: Generate recommendation matrix."""
        matrix = {
            "simple_content": {
                "description": "1-3 items, clear structure",
                "recommended": "enhanced",
                "reason": "Sufficient constraints without overhead"
            },
            "complex_content": {
                "description": "4+ items, mixed types",
                "recommended": "strict",
                "reason": "Needs explicit decision tree"
            },
            "mixed_media": {
                "description": "Has images, quotes, charts",
                "recommended": "strict or checklist",
                "reason": "Multiple placeholder types need explicit rules"
            },
            "new_users": {
                "description": "First time using skill",
                "recommended": "examples",
                "reason": "Learn by pattern matching"
            }
        }

        # All scenarios should have recommendations
        for scenario, rec in matrix.items():
            assert rec["recommended"] in [v.name for v in ALL_VARIATIONS]
            assert rec["reason"]

    def test_rec02_generate_optimal_skill_instructions(self):
        """REC02: Combine best elements into optimal instructions."""
        optimal_instructions = """
## Template-Based Workflow

### Step 1: Analyze Template
- Count slides and identify layout types
- Note placeholder counts per slide (1-col, 2-col, 3-col, image, quote)

### Step 2: Content Analysis (DO THIS FIRST)
- Count your content items
- Identify content types (comparison, pillars, list, quote, etc.)

### Step 3: Layout Matching

RULES:
- 2 items → 2-column layout ONLY
- 3 items → 3-column layout ONLY
- 4+ items → bullet list OR multiple slides
- Quote with attribution → quote layout
- No image → skip image layouts

FORBIDDEN (these create visual problems):
- 2 items in 3-column (empty column)
- Image layout without actual image
- Quote layout without speaker attribution

### Step 4: Verification Checklist
Before applying replacements, verify:
□ Placeholder count matches content count?
□ No empty placeholders will remain?
□ Content fits without overflow?

### Step 5: Apply Replacements
Only proceed after checklist passes.
        """

        # Optimal should incorporate elements from multiple variations
        assert "RULES:" in optimal_instructions
        assert "FORBIDDEN" in optimal_instructions
        assert "Checklist" in optimal_instructions
        assert "→" in optimal_instructions  # Has examples


# ============================================================================
# OUTPUT: Skill Variation Test Summary
# ============================================================================

def generate_test_summary():
    """Generate a summary document of all variations and scenarios."""
    summary = {
        "variations": [
            {
                "name": var.name,
                "description": var.description,
                "instruction_length": len(var.instructions),
                "expected_impact": var.expected_impact
            }
            for var in ALL_VARIATIONS
        ],
        "scenarios": [
            {
                "id": s.id,
                "name": s.name,
                "content_items": s.content_items,
                "correct_choice": s.correct_choice,
                "why_matters": s.why_matters
            }
            for s in SCENARIOS
        ],
        "recommendation": "Use 'strict' or 'checklist' variation for best layout matching"
    }
    return summary


class TestSummaryGeneration:
    """Generate test summary for documentation."""

    def test_generate_summary(self, temp_dir):
        """Generate summary JSON file."""
        summary = generate_test_summary()

        output_path = temp_dir / "skill_variation_summary.json"
        with open(output_path, "w") as f:
            json.dump(summary, f, indent=2)

        assert output_path.exists()
        assert len(summary["variations"]) == len(ALL_VARIATIONS)
        assert len(summary["scenarios"]) == len(SCENARIOS)
