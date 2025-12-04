# PowerPoint Skill Manual Test Checklist

This checklist contains prompts for manually testing Claude Code's PowerPoint capabilities. For each test, provide a PPTX file path and the prompt to Claude Code, then mark as passed/failed.

**How to use:**
1. Have a PPTX file ready (any corporate template, pitch deck, or presentation)
2. Copy the prompt and replace `[PATH]` with your actual file path
3. Run the prompt in Claude Code
4. Mark the checkbox if it completes successfully

---

## Section 1: Text Extraction & Analysis

### 1.1 Basic Text Extraction
- [ ] **Extract all text**
  ```
  Extract all text from [PATH] and show me what's on each slide.
  ```

- [ ] **Summarize presentation**
  ```
  Read [PATH] and give me a summary of what this presentation is about.
  ```

- [ ] **Count slides and content**
  ```
  How many slides are in [PATH]? List the title of each slide.
  ```

### 1.2 Detailed Analysis
- [ ] **Analyze structure**
  ```
  Analyze the structure of [PATH]. What slide layouts are used? How is content organized?
  ```

- [ ] **Find specific content**
  ```
  Search [PATH] for any mentions of [KEYWORD]. Which slides contain it?
  ```

- [ ] **Extract bullet points**
  ```
  Extract all bullet points from [PATH] and organize them by slide.
  ```

---

## Section 2: Template Analysis & Thumbnails

### 2.1 Visual Analysis
- [ ] **Generate thumbnail grid**
  ```
  Create a thumbnail grid of [PATH] so I can see all slides at once.
  ```

- [ ] **Analyze template design**
  ```
  Analyze the design of [PATH]. What colors, fonts, and layouts does this template use?
  ```

- [ ] **Identify slide types**
  ```
  Look at [PATH] and categorize each slide by type (title slide, content slide, divider, etc.).
  ```

### 2.2 Template Inventory
- [ ] **Create template inventory**
  ```
  Create a detailed inventory of [PATH] listing every slide with its index, layout type, and what placeholders/shapes it contains.
  ```

- [ ] **Map reusable slides**
  ```
  Which slides in [PATH] would be good templates to reuse? Create a mapping of slide indices to their purpose.
  ```

---

## Section 3: Text Replacement (Template-Based Editing)

### 3.1 Simple Replacements
- [ ] **Replace title text**
  ```
  In [PATH], change the title on slide 1 to "New Company Name" and save as a new file.
  ```

- [ ] **Replace multiple text elements**
  ```
  In [PATH], replace:
  - The main title with "Q1 2025 Report"
  - The subtitle with "Financial Overview"
  Save as a new file.
  ```

- [ ] **Update bullet points**
  ```
  On slide 2 of [PATH], replace the bullet points with:
  - Revenue increased 15%
  - Customer base grew to 10,000
  - Launched 3 new products
  Save as a new file.
  ```

### 3.2 Batch Text Replacement
- [ ] **Find and replace across presentation**
  ```
  In [PATH], replace all instances of "2024" with "2025" and save as a new file.
  ```

- [ ] **Localization test**
  ```
  Take [PATH] and replace the English text with German translations (you can use placeholder German). Save as a new file.
  ```

### 3.3 Formatted Replacements
- [ ] **Replace with formatting**
  ```
  On slide 1 of [PATH], replace the title with "Important Announcement" in bold red text. Save as a new file.
  ```

- [ ] **Multi-paragraph with bullets**
  ```
  Replace the content area on slide 2 of [PATH] with:

  Key Highlights:
  • First major point (bold)
  • Second point with details
  • Third concluding point

  Save as a new file.
  ```

---

## Section 4: Slide Manipulation

### 4.1 Slide Rearrangement
- [ ] **Reorder slides**
  ```
  Take [PATH] and reorder the slides so slide 3 comes first, then slide 1, then slide 2. Save as a new file.
  ```

- [ ] **Duplicate a slide**
  ```
  In [PATH], duplicate slide 2 three times and save as a new file.
  ```

- [ ] **Remove slides**
  ```
  Create a version of [PATH] with only slides 1, 3, and 5. Save as a new file.
  ```

### 4.2 Complex Rearrangement
- [ ] **Build presentation from template slides**
  ```
  Using [PATH] as a template, create a new presentation with:
  - Slide 0 (title slide)
  - Slide 2 duplicated twice
  - Slide 4
  - Slide 0 again at the end
  Save as a new file.
  ```

- [ ] **Extract subset**
  ```
  Extract slides 5-10 from [PATH] into a new presentation.
  ```

---

## Section 5: Creating Presentations from Scratch (HTML Workflow)

### 5.1 Simple Slides
- [ ] **Create title slide**
  ```
  Create a new presentation with a single title slide that says "Welcome to Our Company" with subtitle "Annual Meeting 2025". Use a professional blue color scheme.
  ```

- [ ] **Create bullet slide**
  ```
  Create a presentation with one slide containing:
  Title: "Our Goals"
  Bullets:
  - Increase revenue by 20%
  - Expand to 3 new markets
  - Launch mobile app
  - Improve customer satisfaction

  Use a clean, modern design.
  ```

### 5.2 Multi-Slide Presentations
- [ ] **Create 3-slide presentation**
  ```
  Create a presentation with:
  1. Title slide: "Project Proposal"
  2. Content slide: "The Problem" with 3 bullet points about challenges
  3. Content slide: "Our Solution" with 3 bullet points about how we solve it

  Use a consistent professional design.
  ```

- [ ] **Create presentation with columns**
  ```
  Create a slide with a two-column layout:
  Left column: "Pros" with 3 bullet points
  Right column: "Cons" with 3 bullet points

  Use contrasting colors for each column.
  ```

### 5.3 Advanced Layouts
- [ ] **Three-column layout**
  ```
  Create a slide showing 3 product tiers:
  - Basic ($9/mo): 3 features
  - Pro ($29/mo): 5 features
  - Enterprise ($99/mo): 8 features

  Use a three-column layout with distinct styling for each tier.
  ```

- [ ] **Create from outline**
  ```
  Create a 5-slide presentation from this outline:

  1. Title: "Marketing Strategy 2025"
  2. Market Analysis (3 key findings)
  3. Target Audience (demographics breakdown)
  4. Campaign Plan (4 initiatives)
  5. Budget & Timeline

  Design it professionally.
  ```

---

## Section 6: Template + Content Workflow

### 6.1 Apply Content to Template
- [ ] **Fill template with new content**
  ```
  Use [PATH] as a template. Keep slides 0, 2, and 4. Fill them with content about a fictional software product launch:
  - Slide 0: Product name and tagline
  - Slide 2: Key features (3-4 points)
  - Slide 4: Call to action

  Save as a new file.
  ```

- [ ] **Create client version**
  ```
  Using [PATH] as a template, create a version for "Acme Corporation" by:
  1. Replacing company name references
  2. Updating the title slide
  3. Keeping the same structure but with Acme-specific content

  Save as a new file.
  ```

### 6.2 Complex Template Operations
- [ ] **Multi-version generation**
  ```
  Using [PATH] as a template, create 2 different versions:
  1. Version for "Tech Startup" audience
  2. Version for "Enterprise" audience

  Same structure, different messaging. Save both as separate files.
  ```

---

## Section 7: Edge Cases & Limitations

### 7.1 Special Characters
- [ ] **Unicode text**
  ```
  Create a slide with text in multiple languages:
  - English: "Hello World"
  - Japanese: "こんにちは"
  - Chinese: "你好世界"
  - Symbols: © ® ™ € £ ¥
  ```

- [ ] **Special formatting**
  ```
  In [PATH], add a slide with:
  - Superscript: E=mc²
  - Arrows: → ← ↑ ↓
  - Math symbols: ≤ ≥ ≠ ±
  ```

### 7.2 Known Limitations (Expected to have issues)
- [ ] **Image insertion** (LIMITATION)
  ```
  Add an image to slide 2 of [PATH].
  ```
  *Expected: Should explain this is not directly supported*

- [ ] **Chart creation** (LIMITATION)
  ```
  Create a bar chart in [PATH] showing Q1-Q4 sales data.
  ```
  *Expected: Should explain charts are limited to placeholders*

- [ ] **Animation** (LIMITATION)
  ```
  Add a fade-in animation to the title on slide 1 of [PATH].
  ```
  *Expected: Should explain animations are not supported*

### 7.3 Error Handling
- [ ] **Invalid file path**
  ```
  Extract text from /nonexistent/path/fake.pptx
  ```
  *Expected: Should error gracefully*

- [ ] **Overflow detection**
  ```
  On slide 1 of [PATH], replace the title with a 500-word paragraph.
  ```
  *Expected: Should warn about text overflow*

---

## Section 8: Real-World Scenarios

### 8.1 Business Use Cases
- [ ] **Quarterly report update**
  ```
  Take [PATH] (quarterly report template) and update it for Q1 2025:
  - Update all date references
  - Change revenue figure to $4.2M
  - Update customer count to 15,000
  Save as a new file.
  ```

- [ ] **Pitch deck customization**
  ```
  Customize [PATH] (pitch deck) for a meeting with "GlobalTech Ventures":
  - Add their name to the title slide
  - Tailor the problem statement to their industry
  - Keep the same structure
  Save as a new file.
  ```

### 8.2 Content Workflows
- [ ] **Extract and rebuild**
  ```
  Extract all content from [PATH], analyze it, then rebuild it as a cleaner, more focused 5-slide presentation. Save as a new file.
  ```

- [ ] **Merge presentations**
  ```
  Take slides 1-3 from [PATH1] and slides 2-4 from [PATH2] and combine them into a single presentation.
  ```

### 8.3 Batch Operations
- [ ] **Update branding**
  ```
  In [PATH], replace all instances of the old company name "OldCorp" with "NewCorp Inc." and update the copyright year to 2025. Save as a new file.
  ```

---

## Test Results Summary

| Section | Passed | Failed | Total |
|---------|--------|--------|-------|
| 1. Text Extraction | | | 6 |
| 2. Template Analysis | | | 5 |
| 3. Text Replacement | | | 7 |
| 4. Slide Manipulation | | | 5 |
| 5. Creating from Scratch | | | 6 |
| 6. Template + Content | | | 3 |
| 7. Edge Cases | | | 6 |
| 8. Real-World Scenarios | | | 5 |
| **TOTAL** | | | **43** |

---

## Notes & Observations

*Record any issues, unexpected behaviors, or feedback here:*

### Issues Found:


### Suggestions for Improvement:


### Working Well:

