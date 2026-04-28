---
name: mck-ppt-design
description: >-
  Create professional, consultant-grade PowerPoint presentations from scratch
  using MckEngine, a python-pptx-based presentation engine with McKinsey-style
  design discipline. Use this skill for pitch decks, board presentations,
  strategy reviews, quarterly business reviews, investment materials, operating
  updates, and other executive presentations. The engine provides high-level
  methods such as eng.cover(), eng.toc(), eng.big_number(), eng.timeline(),
  eng.grouped_bar(), eng.table_insight(), and eng.donut(), along with strict
  typography, anti-corruption XML cleanup, BLOCK_ARC native circular charts,
  spacing and overflow guard rails, and structured layout guidance for
  consultant-style storytelling. This is a cleaned, maintainable skill spec that
  preserves the important behaviors of the original while removing duplication,
  translating non-English documentation to English, adding explicit
  security/privacy controls for outbound integrations, and incorporating Pyramid
  Principle communication rules for answer-first executive messaging.
---

# McKinsey PPT Design Framework

> **Copyright © 2024-2026 Kaku Li (https://github.com/likaku).** Licensed under Apache 2.0. See `NOTICE`. remixed by albertojb to translate from non-english and to add Barbara Minto Pyramid Principle instructions.

> **Document type**: Cleaned skill specification  
> **Intent**: Preserve the original framework's important behaviors while removing repeated examples, duplicate explanations, and mixed-language documentation  
> **Base version reviewed**: 2.2.0 with later version-history notes through 2.3.0  
> **Required tools**: Read, Write, Bash  
> **Runtime requirements**: python3, pip

## What this version is

This document is a **cleaned skill spec**, not a full-fidelity archival rewrite. It keeps the behaviors and rules that matter for actually using the framework well:

- Message-led executive storytelling.
- Layout selection logic.
- Design system rules.
- MckEngine usage pattern.
- File-integrity and anti-corruption guidance.
- Production guard rails.
- Important security and privacy constraints.

It intentionally removes or collapses:

- Repeated examples that illustrate the same rule more than once.
- Duplicated troubleshooting explanations.
- Long version-history details that do not change current usage.
- Repetitive API descriptions where methods follow the same pattern.
- Mixed Chinese/English documentation outside necessary CJK compatibility guidance.

Use this document as the maintainable specification. Use the original source only when historical detail or edge-case rationale is needed.

## Purpose

This skill defines a framework for generating professional business presentations with consultant-grade structure, visual consistency, and implementation discipline. It combines three capabilities:

1. **Communication architecture** for executive-ready, answer-first messaging.
2. **Design system** for typography, color, spacing, and slide consistency.
3. **Presentation engine guidance** for safe, repeatable `.pptx` generation with MckEngine.

The framework is appropriate when the goal is not only to create slides that look polished, but also to communicate a clear recommendation, argument, or decision.

## When to use this skill

Use this skill when users ask to:

1. Create business presentations, strategy decks, board materials, investment decks, quarterly reviews, or executive updates.
2. Generate PowerPoint files programmatically with `python-pptx` through a higher-level engine.
3. Apply consulting-style slide design with flat visuals, clean hierarchy, and minimal decorative noise.
4. Build structured slides such as executive summaries, data dashboards, timelines, comparisons, matrices, roadmaps, and case studies.
5. Fix presentation quality issues such as file corruption, inconsistent fonts, spacing defects, poor chart rendering, or weak title structure.
6. Turn raw business content into a logically sequenced, message-led deck.

## Core principles

### Communication first

A strong presentation starts with the message, not the layout. Before selecting slide patterns, define the audience, the decision to be made, the governing thought, and the minimum evidence required to support it.

### Answer first

The deck should lead with the answer and then support it. Do not begin with background unless the audience needs that context in order to interpret the recommendation.

### One idea per slide

Each slide should communicate one governing thought. Supporting evidence can appear beneath it, but the slide should not try to make several unrelated arguments at once.

### Minimalist, not empty

Minimalism means removing visual clutter such as shadows, gradients, decorative icons, and unnecessary color blocks. It does **not** mean under-filling the slide or replacing analysis with whitespace.

### Consistency over novelty

Consistent title treatment, spacing, color use, and text hierarchy matter more than clever layout variation. Variety is useful, but only when it supports comprehension.

## Communication architecture

### Pyramid Principle

This skill should produce presentations that follow Barbara Minto's Pyramid Principle:

1. **Start with the answer.** Lead with the recommendation, conclusion, or core takeaway.
2. **Group supporting points.** Organize the next layer into two to four distinct supporting ideas.
3. **Order the points logically.** Use a clear order such as priority, chronology, causation, or issue tree logic.
4. **Support with evidence.** Use charts, examples, tables, and facts to prove the assertion.
5. **Drive to action.** Make clear what the audience should approve, decide, fund, prioritize, or monitor.

### MECE grouping

Where possible, supporting points should be **mutually exclusive, collectively exhaustive**. In practice, this means parallel buckets should not overlap, and together they should address the question fully.

### Conclusion-led headlines

Slide titles must state a conclusion, not just name a topic.

- Good: `Margin pressure is concentrated in two product lines`
- Weak: `Margin analysis`
- Good: `Spain is the best near-term expansion market`
- Weak: `Expansion options`

### Storyboarding rule

Before generating slides, create a storyboard with one sentence per slide. That sequence should read like a short executive memo. Only after the storyline is stable should layout selection begin.

### Recommended executive flow

For most business presentations, use this top-down sequence:

1. Recommendation or governing thought.
2. Why this conclusion is correct.
3. Evidence and implications.
4. Risks, trade-offs, or alternatives.
5. Implementation path and next steps.

## Presentation planning

### Standard deck templates

#### Standard presentation (10-12 slides)

1. Cover slide.
2. Table of contents.
3. Executive summary or governing thought.
4. Supporting argument 1.
5. Supporting argument 2.
6. Supporting argument 3.
7. Evidence or case study.
8. Data exhibit or comparison.
9. Synthesis or implications.
10. Recommendation and next steps.
11. Optional roadmap or implementation slide.
12. Closing slide.

#### Short presentation (6-8 slides)

1. Cover slide.
2. Executive summary.
3. Argument 1.
4. Argument 2.
5. Evidence or case study.
6. Synthesis or recommendation.
7. Next steps.
8. Closing slide.

### Mandatory planning rules

- Minimum slide count for any substantive topic: **8 slides**.
- Generate the full deck in one script run; do not stop after a partial sequence.
- The table of contents must enumerate the major sections actually present in the deck.
- Slides 2-5 should use strong editorial or insight-led layouts, not plain text.
- Every major section should answer a visible audience question.

### Content density rules

Each content slide should include:

- A conclusion-led title.
- At least three meaningful visual zones, such as a title bar, a body area, and an insight panel.
- Enough content to use at least half of the practical content area.
- Source attribution when data or external facts appear.
- A clear implication, recommendation, or takeaway.

## Layout selection logic

Choose layouts based on the communication task, not on aesthetic preference alone.

| Content need | Preferred layout types | Avoid |
|---|---|---|
| Single critical statistic | Big Number, Metric Card, Key Takeaway | Plain paragraph |
| Two options or scenarios | Side-by-Side, Before/After, Comparison row | Generic two-column text |
| Three or four parallel ideas | Table+Insight, Four-Column, Metric Cards, Icon Grid | Bullet list |
| Process or implementation path | Process Chevron, Vertical Steps, Value Chain, Timeline | Long numbered text |
| Time series or trend data | Grouped Bar, Line Chart, Stacked Bar, Stacked Area | Table-only slide |
| Composition or share | Donut, Pie, Stacked Bar | Bullet list |
| Case study or example | Case Study, Case Study with Image, Content + Image | Dense text block |
| Summary or synthesis | Executive Summary, Key Takeaway, Table+Insight | Topic-only title + bullets |
| Risk or prioritization | 2x2 Matrix, Risk Matrix, SWOT, Harvey Ball table | Plain text |

### Opening slide priority

For the first few content slides after the cover and table of contents, strongly prefer high-impact layouts such as:

1. `table_insight`
2. `big_number`
3. `key_takeaway`
4. `executive_summary`

These patterns establish the deck's message architecture early and improve executive readability.

### Layout diversity rule

Neighboring content slides should not reuse the same layout unless there is a clear analytical reason to do so. Repetition makes the deck feel mechanical and reduces visual energy.

## MckEngine quick start

MckEngine is the preferred runtime path. Do not write raw shape-coordinate code when a supported engine method already exists.

### Setup

```bash
pip install python-pptx lxml
```

If optional local image processing is used:

```bash
pip install pillow numpy rembg
```

If optional cloud cover-image generation is enabled explicitly:

```bash
pip install tencentcloud-sdk-python pillow numpy rembg
```

### Import pattern

```python
import sys, os
sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/mck-ppt-design'))
from mck_ppt import MckEngine
from mck_ppt.constants import *
from pptx.util import Inches
```

### Complete generation pattern

```python
import sys, os
sys.path.insert(0, os.path.expanduser('~/.workbuddy/skills/mck-ppt-design'))
from mck_ppt import MckEngine
from mck_ppt.constants import *
from pptx.util import Inches

eng = MckEngine(total_slides=10)

eng.cover(
    title='Q1 2026 strategy review',
    subtitle='Board update',
    author='Strategy team',
    date='March 2026'
)

eng.toc(items=[
    ('1', 'Executive summary', 'Recommendation and key evidence'),
    ('2', 'Market dynamics', 'Where pressure and opportunity are shifting'),
    ('3', 'Strategic actions', 'Three moves to accelerate growth'),
])

eng.executive_summary(
    title='Revenue can return to double-digit growth with three targeted actions',
    headline='Growth is available, but it is concentrated in two channels and one product tier',
    items=[
        ('1', 'Shift mix toward premium bundles', 'Premium bundles drive higher margin with limited cost-to-serve impact'),
        ('2', 'Expand in underpenetrated channels', 'Two distributor channels remain underdeveloped relative to peers'),
        ('3', 'Fund growth through operating simplification', 'Back-office complexity can be reduced without harming service levels'),
    ],
    source='Source: internal analysis, Q1 2026'
)

eng.grouped_bar(
    title='Revenue growth is strongest in premium and partner-led channels',
    categories=['Q1', 'Q2', 'Q3', 'Q4'],
    series=[('Premium', NAVY), ('Partner', ACCENT_BLUE)],
    data=[[120, 80], [145, 95], [160, 110], [180, 130]],
    max_val=200,
    source='Source: finance team'
)

eng.table_insight(
    title='Three actions improve growth while protecting margin',
    headers=['Action', 'Mechanism', 'Expected impact'],
    rows=[
        ['Premium bundles', 'Raise mix in high-value segments', 'Revenue +12%'],
        ['Channel expansion', 'Add distributor reach in two verticals', 'Revenue +8%'],
        ['Cost simplification', 'Reduce duplication across support functions', 'Margin +3pp'],
    ],
    insights=[
        'The actions are complementary rather than sequential.',
        'The first two accelerate growth; the third funds execution.',
        'Each action is feasible within the next planning cycle.',
    ],
    source='Source: strategy team'
)

eng.timeline(
    title='A phased rollout reduces execution risk and speeds learning',
    milestones=[
        ('Q1', 'Design and prioritization'),
        ('Q2', 'Pilot launch'),
        ('Q3', 'Scale-up'),
        ('Q4', 'Measure and refine'),
    ],
    source='Source: PMO'
)

eng.closing(title='Thank you', message='Discussion and decision points')
eng.save('output/q1_strategy_review.pptx')
```

### Engine rules

1. One engine method should produce one slide.
2. `eng.save()` should handle cleanup and sanitization automatically.
3. Page numbers should be driven by `total_slides`.
4. Use constants from `mck_ppt.constants` rather than ad hoc formatting values.
5. Prefer engine methods over low-level shape construction.

## Security and privacy

### Security posture

This skill may support optional outbound integrations, including cloud image generation and messaging-channel delivery. Those features are acceptable only when they are explicitly enabled and clearly appropriate for the user's confidentiality context.

### Default rule

**Outbound features must be disabled by default.** The base assumption for this skill is local generation only.

### Approved outbound categories

Outbound actions may be allowed only when all of the following are true:

1. The integration is explicitly enabled in configuration.
2. The user has requested or approved the behavior.
3. The content is appropriate for transmission to that service.
4. The runtime environment is permitted to use networked services.

### Prohibited default behavior

Do **not** assume that any of the following are allowed automatically:

- Sending prompts, titles, or business context to a cloud image-generation API.
- Sending generated `.pptx` files to a chat, messaging channel, or collaboration system.
- Uploading deck contents, notes, or metadata to third-party services.
- Enabling network-dependent helper functions without explicit configuration.

### Local-only mode

Provide or respect a `local_only` style configuration wherever possible. In local-only mode:

- Cloud cover-image generation is disabled.
- Channel delivery helpers are disabled.
- Only local rendering, local image assets, and local file writes are allowed.

### External integrations documentation rule

Every external integration must be documented with:

- Service name.
- Data sent.
- Purpose of transmission.
- Credentials required.
- Expected output.
- Failure mode and fallback behavior.

### Subprocess and file-transfer guidance

If a helper such as `deliver_to_channel()` exists, it must:

- Be opt-in.
- Validate output paths.
- Avoid shell interpolation patterns that increase command-injection risk.
- Fail safely and save locally if the delivery path is unavailable.
- Be disabled in sensitive or regulated workflows.

### Cover image guidance

If cloud cover-image generation is enabled, the prompt should avoid client-identifying data unless the user explicitly accepts that risk. Prefer abstract or generic subject descriptions over sensitive business language.

### Dependency hygiene

Use pinned dependencies where possible and review core packages periodically, especially:

- `python-pptx`
- `lxml`
- `pillow`
- `numpy`
- `rembg`
- Any cloud SDKs used for outbound calls

## Design system

### Design philosophy

1. **Extreme minimalism**: remove decorative noise and keep attention on content.
2. **Consistency**: repeat visual patterns across slides.
3. **Hierarchy**: make the message path obvious through size, weight, and contrast.
4. **Flat design**: no shadows, gradients, or 3D effects.

### Color palette

| Color | Hex | Usage |
|---|---|---|
| NAVY | `#051C2C` | Primary titles, circles, key emphasis |
| WHITE | `#FFFFFF` | Backgrounds and reverse text |
| BLACK | `#000000` | Rules, separators, title text |
| DARK_GRAY | `#333333` | Body text |
| MED_GRAY | `#666666` | Secondary text and labels |
| LINE_GRAY | `#CCCCCC` | Table and chart separators |
| BG_GRAY | `#F2F2F2` | Insight panels and neutral fills |

### Accent colors

Use accent colors only when multiple parallel items require clear differentiation.

| Accent | Hex | Typical use |
|---|---|---|
| ACCENT_BLUE | `#006BA6` | First comparative item |
| ACCENT_GREEN | `#007A53` | Second comparative item |
| ACCENT_ORANGE | `#D46A00` | Third comparative item |
| ACCENT_RED | `#C62828` | Warning, negative movement, or fourth item |

Accent colors should not replace NAVY as the primary design anchor.

## Typography

### Font policy

- English headings: **Georgia**
- English body text: **Arial**
- CJK text: use an East Asian font such as **KaiTi**, with **SimSun** or another platform-appropriate fallback if needed

### Font sizes

| Size | Use |
|---|---|
| 44pt | Cover title |
| 28pt | Section header |
| 24pt | Cover subtitle |
| 22pt | Action title |
| 18pt | Sub-header |
| 16pt | Emphasis text |
| 14pt | Body text |
| 9pt | Source and footnote text |

No other font sizes should be used unless a layout explicitly requires controlled shrinkage for overflow protection.

### CJK implementation guidance

If the output contains Chinese, Japanese, or Korean text, apply East Asian font settings at the run level to avoid fallback rendering issues in PowerPoint.

```python
def set_ea_font(run, typeface='KaiTi'):
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', typeface)
```

This skill is English-first, but it remains compatible with bilingual output when needed.

## Line treatment

### Rendering rules

1. All lines must be flat.
2. Do not use theme effects or shadowed connectors.
3. Use solid color only.
4. Use consistent widths by context.

### Standard line widths

| Usage | Width | Color |
|---|---|---|
| Title separator | 0.5pt | BLACK |
| Section divider | 0.5pt | BLACK |
| Table header line | 1.0pt | BLACK |
| Table row separator | 0.5pt | LINE_GRAY |
| Timeline guide line | 0.75pt | LINE_GRAY |
| Cover accent line | 2.0pt | NAVY |

### Do not use connectors

Do not use `slide.shapes.add_connector()` for line drawing. Connector shapes can carry theme-style references that create unwanted effects and may contribute to file-corruption behavior in some presentation outputs.

Instead, use thin rectangles for lines:

```python
def add_hline(slide, x, y, length, color=BLACK, thickness=Pt(0.5)):
    from pptx.util import Emu
    h = max(int(thickness), Emu(6350))
    return add_rect(slide, x, y, length, h, color)
```

## Shape and text treatment

### Shape styling

All shapes should have:

- Solid fills.
- No gradient fills.
- No shadow or reflection effects.
- No visible border unless the layout specifically requires one.
- Immediate shape cleanup where required by the engine.

### Text box padding

Apply consistent padding so text does not touch shape edges.

```python
bodyPr = tf._txBody.find(qn('a:bodyPr'))
for attr in ['lIns', 'tIns', 'rIns', 'bIns']:
    bodyPr.set(attr, '45720')
```

### Vertical anchoring

- Action titles should sit low enough to feel anchored to the title separator.
- Body text should typically anchor to the top.
- Small labels can be center-aligned as needed.

## File integrity and cleanup

### Anti-corruption approach

The framework uses a three-part strategy:

1. Prefer safe shape patterns that avoid problematic style references.
2. Clean shapes inline where needed.
3. Run post-save XML cleanup through the engine.

### Full cleanup pattern

```python
import zipfile, os
from lxml import etree

def full_cleanup(outpath):
    tmppath = outpath + '.tmp'
    with zipfile.ZipFile(outpath, 'r') as zin:
        with zipfile.ZipFile(tmppath, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.endswith('.xml'):
                    root = etree.fromstring(data)
                    ns_p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
                    ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
                    for style in root.findall(f'.//{{{ns_p}}}style'):
                        style.getparent().remove(style)
                    if 'theme' in item.filename.lower():
                        for tag in ['outerShdw', 'innerShdw', 'scene3d', 'sp3d']:
                            for el in root.findall(f'.//{{{ns_a}}}{tag}'):
                                el.getparent().remove(el)
                    data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                zout.writestr(item, data)
    os.replace(tmppath, outpath)
```

When using MckEngine, prefer `eng.save()` rather than manually repeating cleanup logic.

## Production guard rails

These rules are mandatory because they address recurring defects in real presentation generation.

### Rule 1: Bottom-bar spacing

Maintain at least `0.15"` between the last content block and any bottom summary bar.

### Rule 2: Overflow protection

Every element must stay inside content margins. Text inside a filled box must be inset rather than placed edge-to-edge.

### Rule 3: Bottom whitespace control

Avoid leaving large empty gaps between charts or content blocks and bottom insight bars. The space should feel deliberate, not accidental.

### Rule 4: Legend color consistency

Chart legends must use colored squares that match the plotted series. Do not fake legends with monochrome text glyphs.

### Rule 5: Title style consistency

Use `add_action_title()` as the default title pattern for content slides. Reserve alternate title bars for explicit user requests only.

### Rule 6: Axis label centering

In matrix and grid charts, center axis labels on the full span of the axis rather than placing them at arbitrary offsets.

### Rule 7: Visual-relief slide requirement

For decks with eight or more slides, include at least one image-containing slide or image placeholder slide to break up walls of text and charts.

### Rule 8: Dynamic sizing for variable-count layouts

Whenever the number of rows, stages, or parallel cards changes, compute dimensions dynamically rather than relying on fixed widths or heights.

### Rule 9: BLOCK_ARC for circular charts

Use `BLOCK_ARC` for donut, pie, and gauge charts rather than simulating arcs with hundreds of tiny rectangles.

### Rule 10: Horizontal overflow protection

For horizontal item layouts, compute item widths dynamically so gaps never become negative and shapes never acquire invalid dimensions.

## BLOCK_ARC chart guidance

For circular charts, use native PowerPoint `BLOCK_ARC` shapes. This reduces file size, improves rendering quality, and avoids jagged charts built from hundreds of small rectangles.

### Angle convention

PowerPoint measures angles clockwise from 12 o'clock:

- Top = `0°`
- Right = `90°`
- Bottom = `180°`
- Left = `270°`

### Helper pattern

```python
from pptx.oxml.ns import qn

def add_block_arc(slide, left, top, width, height, start_deg, end_deg, inner_ratio, color):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.BLOCK_ARC, left, top, width, height)
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    _clean_shape(sh)

    sp = sh._element.find(qn('p:spPr'))
    prstGeom = sp.find(qn('a:prstGeom'))
    if prstGeom is not None:
        avLst = prstGeom.find(qn('a:avLst'))
        if avLst is None:
            avLst = prstGeom.makeelement(qn('a:avLst'), {})
            prstGeom.append(avLst)
        for gd in avLst.findall(qn('a:gd')):
            avLst.remove(gd)
        avLst.append(avLst.makeelement(qn('a:gd'), {'name': 'adj1', 'fmla': f'val {int(start_deg * 60000)}'}))
        avLst.append(avLst.makeelement(qn('a:gd'), {'name': 'adj2', 'fmla': f'val {int(end_deg * 60000)}'}))
        avLst.append(avLst.makeelement(qn('a:gd'), {'name': 'adj3', 'fmla': f'val {inner_ratio}'}))
    return sh
```

## Mandatory slide elements

Every content slide except Cover and Closing should include:

- Action title.
- Title separator line.
- Main content area.
- Source attribution.
- Page number.

## API usage guidance

This cleaned spec does **not** reproduce the original's full method-by-method catalog. Instead, it keeps the method families and the usage logic that matter most for maintenance and generation quality.

### Structure and narrative

- `eng.cover(...)`
- `eng.toc(...)`
- `eng.section_divider(...)`
- `eng.executive_summary(...)`
- `eng.key_takeaway(...)`
- `eng.table_insight(...)`
- `eng.closing(...)`

Use these first when the goal is to establish storyline, argument, and synthesis.

### Data and comparison

- `eng.big_number(...)`
- `eng.two_stat(...)`
- `eng.three_stat(...)`
- `eng.data_table(...)`
- `eng.side_by_side(...)`
- `eng.before_after(...)`
- `eng.pros_cons(...)`
- `eng.scorecard(...)`
- `eng.rag_status(...)`

Use these to support claims with measurable evidence.

### Process and implementation

- `eng.timeline(...)`
- `eng.vertical_steps(...)`
- `eng.process_chevron(...)`
- `eng.value_chain(...)`
- `eng.decision_tree(...)`
- `eng.agenda(...)`

Use these when the audience needs to understand sequence, dependencies, or execution logic.

### Charts

- `eng.grouped_bar(...)`
- `eng.stacked_bar(...)`
- `eng.horizontal_bar(...)`
- `eng.line_chart(...)`
- `eng.waterfall(...)`
- `eng.pareto(...)`
- `eng.stacked_area(...)`
- `eng.donut(...)`
- `eng.pie(...)`
- `eng.gauge(...)`

Use charts when data contains time periods, comparisons, composition, ranking, or trend movement.

### Visual and image layouts

- `eng.content_right_image(...)`
- `eng.three_images(...)`
- `eng.image_four_points(...)`
- `eng.full_width_image(...)`
- `eng.case_study_image(...)`
- `eng.two_col_image_grid(...)`

Use these when a visual materially improves the narrative.

## Cover image policy

### Preferred approach

Default to either:

- A clean text-only cover slide, or
- A locally supplied image asset.

### Optional cloud mode

Cloud-generated cover images may be used only if explicitly enabled and appropriate for the confidentiality context.

If `cover_image='auto'` is supported in the runtime:

- Treat it as optional.
- Do not assume it is allowed.
- Avoid sending sensitive text or client names in prompts.
- Provide a local fallback whenever possible.

## Best practices

1. Use MckEngine rather than raw coordinate-heavy code.
2. Build the full storyline before building slides.
3. Use conclusion-led titles on every content slide.
4. Match the layout to the analytical task.
5. Prefer charts over text when data is quantitative.
6. Include at least one visual-relief slide in longer decks.
7. Keep typography and spacing disciplined.
8. Default to local-only behavior unless outbound integrations are explicitly approved.
9. Use `eng.save()` as the canonical output path.

## Common mistakes to avoid

- Topic-only slide titles.
- Bullet-heavy opening slides.
- Two-column text repeated across several slides.
- Circular charts built from tiny rectangles.
- Legends whose colors do not match the chart.
- Text boxes that touch shape edges.
- Fixed-size row or stage layouts for variable item counts.
- Automatic cloud calls or file delivery in confidential contexts.
- Mixing several unrelated messages on one slide.

## Example slide-title rewrites

| Weak title | Strong title |
|---|---|
| Market overview | Share is shifting to low-cost entrants while demand remains resilient |
| Financial analysis | Margin erosion is driven by fixed overhead rather than labor |
| Strategic options | The best path is to expand in Spain while simplifying support operations |
| Customer feedback | Customers value speed and reliability more than feature breadth |
| Implementation plan | A phased rollout reduces execution risk and accelerates learning |

## Maintenance notes

This cleaned skill spec should be the primary file for day-to-day use and future editing. If future maintainers need to restore more edge-case detail, the best approach is to add concise appendices selectively rather than rebuilding the original duplication-heavy structure.

This cleaned skill spec should be the primary file for day-to-day use and future editing. If future maintainers need to restore more edge-case detail, the best approach is to add concise appendices selectively rather than rebuilding the original duplication-heavy structure.
