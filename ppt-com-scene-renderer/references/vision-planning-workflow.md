# Vision Planning Workflow

Use this reference when the user provides one or more images to recreate or combine, or when a generated PPT preview image needs QA.

Vision is an analysis and planning step. It converts visual inputs into structured observations and rendering decisions. It does not operate PowerPoint, replace `scripts/Render-PptScene.ps1`, or guarantee exact coordinates.

## Required Flow

1. Inspect each supplied image before writing `scene.json`.
2. Produce `vision-analysis.md` for reference-image recreation or multi-image composition.
3. Map visual observations to native PPT shapes, text boxes, lines, curves, groups, components, and local image/vector assets.
4. Mark complex details that should be generated or inserted as PNG/SVG/EMF assets.
5. After rendering with `-ExportPreview`, inspect `preview.png` and list corrections before rerendering when fidelity matters.

## Vision Analysis Template

Save this as `output/<task-name>/vision-analysis.md` when images are part of the task.

```markdown
# Vision Analysis

## Inputs
- Source images:
- Task type:
- Target output:

## Overall Structure
- Canvas/aspect ratio:
- Main regions:
- Reading order:
- Alignment/grid:
- Visual hierarchy:

## Element Inventory
| ID | Source | Type | Approx position | Content | PPT strategy |
| --- | --- | --- | --- | --- | --- |
| element-1 | image-1 | rect/text/line/icon/etc. | top-left/center/etc. | label or visual role | native shape/text/image/vector |

## Text and Labels
| Text | Approx location | Priority | Notes |
| --- | --- | --- | --- |

## Connectors and Relationships
| From | To | Connector type | Arrow/dash/style | Notes |
| --- | --- | --- | --- | --- |

## Style Extraction
- Background:
- Palette and color roles:
- Stroke widths:
- Typography:
- Corner radius:
- Fill/opacity/gradient:
- Icon/image style:

## Native PPT Mapping
- Rectangles/cards:
- Circles/ellipses:
- Lines/arrows:
- Curves/freeforms:
- Groups/components:
- Theme tokens:

## Complex Asset Candidates
| Asset | Why native PPT is insufficient | Preferred format | Placement |
| --- | --- | --- | --- |

## Multi-Image Composition Notes
- Elements selected from each source:
- Style conflicts:
- Unified visual system:
- Elements intentionally omitted:

## Fidelity Risks
- Hard-to-read text:
- Dense areas:
- Ambiguous arrows:
- Complex pictograms:
- Likely manual refinements:
```

## Preview QA Template

Append this section to `vision-analysis.md` or save it as `preview-qa.md` after `-ExportPreview`.

```markdown
# Preview QA

## Preview Checked
- Preview path:
- Compared against:

## Pass/Fail Summary
- Overall:
- Highest-priority fixes:

## Issues
| Area | Expected | Actual | Correction |
| --- | --- | --- | --- |

## Scene-Level Fixes
- Coordinate/layout changes:
- Text changes:
- Connector changes:
- Color/style changes:
- Asset changes:
```

## Mapping Rules

- Use Vision to identify structure first, then choose precise coordinates in `scene.json`.
- Prefer native PPT shapes for information-bearing elements so the final deck stays editable.
- Use `freeform`/`bezier` for curved arrows, loops, and hand-drawn-like connectors.
- Use `group` and reusable `component` elements for repeated visual grammar.
- Use generated local PNG/SVG/EMF only for complex pictograms, equipment details, or illustrations that would be inefficient as native PPT.
- Keep source provenance in IDs for multi-image tasks, such as `img1-process-card` or `img2-loop-arrow`.

## Preview QA Rules

- Check missing elements, text mismatches, arrow direction, ordering, relative spacing, color drift, and image placement.
- Treat preview QA as a correction pass, not a reason to copy pixels.
- If a mismatch is caused by renderer limitations, record the limitation and choose the closest editable PPT-native representation.
- If an image-only element is more faithful than an editable approximation, use it only for non-critical pictograms or decorative details.
