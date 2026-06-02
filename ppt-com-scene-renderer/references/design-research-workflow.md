# Description-Driven Figure Design Workflow

Use this reference when the user asks for an original PPT figure from a text description, asks for Nature/Science-style inspiration, or wants a figure designed before rendering.

For tasks with supplied images, read `vision-planning-workflow.md` first and use the resulting `vision-analysis.md` as an input to the design plan. For text-only tasks, Vision is not required unless a rendered preview image is being checked.

## Required Flow

1. Clarify only high-impact missing intent. Otherwise proceed with reasonable assumptions.
2. If images are supplied, complete Vision analysis before design planning.
3. Research visual references when the request calls for inspiration, benchmarking, current examples, or journal/design style.
4. Extract abstract patterns, not source content.
5. Write a `Figure Design Plan`.
6. Wait for user confirmation before rendering unless the user explicitly asks to generate directly.
7. Convert the approved design into `scene.json`, local assets, PPTX, and preview.

## Reference Research

Default reference scope:

- Scholarly figures: Nature, Nature sub-journals, Science, Science sub-journals, and related publisher pages.
- Design references: credible design galleries, data visualization portfolios, editorial/scientific illustration examples, product/system diagram examples, and design-system documentation.

For each useful reference, record:

- URL or citation-like source label.
- What abstract pattern was useful.
- What will not be copied.
- How the final design will differ.

Good abstract takeaways:

- "Three-tier system framing with a top actor layer, middle decision layer, and bottom execution layer."
- "Modular card matrix with category color coding and compact bullet hierarchy."
- "Closed-loop workflow arrows that make feedback and iteration visible."
- "Low-saturation background with one high-contrast accent path for the primary story."

Do not reuse:

- Source images, screenshots, icons, or illustrations.
- Exact layout geometry, figure composition, or distinctive annotation paths.
- Identical color palettes, typography pairings, or decorative motifs.
- Source labels, captions, data, or claims unless the user provides permission and citation context.

## Originality Guardrails

- Treat references as mood and structure inputs, not reusable assets.
- Combine multiple abstract patterns and adapt them to the user's content.
- Change spatial structure enough that the figure is independently authored.
- Generate new icons or use generic PPT-native symbols when a pictogram is needed.
- Prefer editable PPT primitives for all information-bearing structure.
- Use raster or vector assets only when native PPT cannot represent the detail at the requested quality.

## Figure Design Plan Template

Save this as `output/<task-name>/design-plan.md` when the task is description-driven.

```markdown
# Figure Design Plan

## Goal
- Figure purpose:
- Audience:
- Output format:

## Core Message
- Main story:
- Secondary messages:

## Reference-Informed Direction
- Abstract patterns to use:
- Patterns intentionally avoided:
- Originality/differentiation notes:

## Vision-Derived Structure
- Source images analyzed:
- Main detected structure:
- Elements to preserve:
- Elements to reinterpret:

## Layout
- Canvas and aspect ratio:
- Main regions:
- Reading order:
- Alignment/grid:

## Visual System
- Background:
- Primary colors and roles:
- Font system:
- Stroke widths:
- Icon/image style:

## Element Inventory
- Native PPT shapes:
- Lines/arrows/curves:
- Text labels:
- Reusable components:
- Generated or inserted assets:

## Rendering Plan
- Scene JSON path:
- Asset root:
- PPTX path:
- Preview path:
- Known risks:

## Preview QA Findings
- Preview image checked:
- Issues found:
- Scene-level corrections:
```

## Inspiration Notes Template

Save this as `output/<task-name>/inspiration-notes.md`. Keep it internal unless the user asks for a source appendix.

```markdown
# Inspiration Notes

## Reference Summary
| Source | Abstract takeaway | Not reused | Differentiation |
| --- | --- | --- | --- |
| URL or source label | Pattern only | Content avoided | How final design differs |

## Final Originality Check
- No source image, icon, screenshot, or proprietary diagram element is reused.
- Layout and styling are newly composed for the user's content.
- Any generated image asset was created for this task and stored locally.
```

## Handoff to Scene JSON

- If `vision-analysis.md` exists, use it as the structural source of truth before external references.
- Start from the approved design plan, not directly from external references.
- Map each layout region to `group` elements when it improves editability.
- Use theme tokens for repeated colors, fonts, and stroke widths.
- Use built-in components for repeated step circles, skill cards, legends, brackets, and flow arrows.
- Use `freeform`/`bezier` for smooth connectors and curved callouts.
- Use `svg`/`emf` for crisp vector icons when available; use `image` for generated raster assets.
- Export a preview with `-ExportPreview` and compare it against the approved design plan. If a reference image or prior plan exists, run Vision QA and record corrections.

## Network-Unavailable Fallback

If web access is unavailable or blocked:

- State that live reference research was not performed.
- Continue using the user's description and local design knowledge.
- Still produce `design-plan.md` and `inspiration-notes.md`; mark references as "not searched".
- Do not invent source URLs.
