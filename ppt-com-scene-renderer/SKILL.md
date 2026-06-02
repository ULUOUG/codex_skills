---
name: ppt-com-scene-renderer
description: Create editable PowerPoint diagrams from rough reference images, mixed visual sources, text-only figure descriptions, preview images, or JSON scene descriptions using Vision-assisted planning and Microsoft PowerPoint COM automation. Use when the user wants to analyze a reference image, recreate an image in PPT, combine elements from multiple images into one slide, plan an original scientific or design-style figure from a description, research visual inspiration without copying, generate grid/network/structure diagrams, insert generated raster icons where PPT native shapes are insufficient, QA a generated PPT preview image, or watch PowerPoint draw elements live.
---

# PPT COM Scene Renderer

## Overview

Use this skill to turn visual intent into an editable `.pptx` by writing a JSON scene and rendering it with PowerPoint COM. Prefer native PPT shapes for geometry, lines, arrows, labels, and grids; insert local PNG/JPG/SVG/EMF assets only for details that are impractical to draw with PowerPoint primitives.

Supported input modes:

- Single-image recreation: reproduce the supplied reference as an editable PPT figure.
- Multi-image composition: take selected elements from several images and unify them into one figure.
- Description-driven original design: start from a text description, research visual inspiration, create a figure design plan, then render the final editable PPT.
- Preview QA: inspect a rendered preview image and identify corrections before finalizing the PPT.

Bundled resources:

- `scripts/Render-PptScene.ps1`: JSON scene to PPTX renderer.
- `references/scene-json-v1.md`: supported JSON schema and examples.
- `references/design-research-workflow.md`: description-driven planning, reference research, originality, and handoff templates.
- `references/vision-planning-workflow.md`: Vision-assisted image decomposition, planning, and preview QA templates.
- `assets/sample/`: sample scene and PNG asset for smoke testing.

## Workflow

1. Classify the input mode.
   - For a supplied image, use single-image recreation.
   - For multiple supplied images, use multi-image composition and identify which elements come from each image.
   - For a text-only request or a request for an original figure, use description-driven original design and read `references/design-research-workflow.md`.
   - For a rendered preview image, use preview QA.
2. Run Vision analysis when images are involved.
   - For single-image recreation, read `references/vision-planning-workflow.md` and create `vision-analysis.md` before writing `scene.json`.
   - For multi-image composition, create one Vision analysis per source image, then summarize selected elements, style conflicts, and the unified composition.
   - For preview QA, compare `preview.png` against the reference image, `vision-analysis.md`, or `design-plan.md`; list concrete fixes before rerendering.
   - Vision only analyzes and plans. It does not replace `scripts/Render-PptScene.ps1` or any PPTX renderer.
3. For description-driven original design, produce a `Figure Design Plan` before writing `scene.json` unless the user explicitly asks to generate the PPT directly.
   - Search scholarly and design references when the user asks for research, current inspiration, Nature/Science-style figures, or broad visual benchmarking.
   - Extract only abstract design patterns: information hierarchy, layout grammar, visual rhythm, color role, annotation strategy, and storytelling flow.
   - Do not copy, trace, screenshot, reuse, or closely imitate external figures, icons, illustrations, palettes, or text placement.
   - Keep an internal `inspiration-notes.md` with reference URLs, abstract takeaways, and how the final design differs.
4. Preserve editability by default.
   - Use PowerPoint-native shapes, text, lines, curves, groups, components, gradients, transparency, and theme references wherever practical.
   - Generate or insert local PNG/SVG/EMF assets only for complex pictograms, equipment images, textures, and details that are not practical as native PPT shapes.
5. Create a task folder in the current workspace, such as `output/<task-name>/`.
   - Put Vision planning notes at `output/<task-name>/vision-analysis.md` when applicable.
   - Put design planning notes at `output/<task-name>/design-plan.md` when applicable.
   - Put inspiration notes at `output/<task-name>/inspiration-notes.md` when applicable.
   - Put scene JSON at `output/<task-name>/scene.json`.
   - Put generated or extracted assets under `output/<task-name>/assets/`.
   - Put the final deck under `output/<task-name>/<task-name>.pptx`.
   - Put preview exports under `output/<task-name>/preview.png` or `output/<task-name>/preview/`.
6. Write `scene.json` using the v1 schema.
   - Use logical canvas coordinates, normally `1280 x 720`; for portrait figures set both `canvas` and `page` to the desired ratio.
   - Draw elements in z-order: lower layers first, foreground labels and icons last.
   - Use `rect`, `ellipse`, `line`, `polyline`, `freeform`, `path`, and `text` for editable PPT content.
   - Use `svg`, `emf`, or `vector` for local vector files; use `image` for raster files.
   - Use `group`, `zIndex`, theme references, and built-in `component` elements for repeated diagram structure.
   - Use `gradient`, `fillOpacity`, `strokeOpacity`, `rotation`, and `shadow` when visual fidelity requires them.
   - Save JSON as UTF-8; the renderer reads UTF-8 strictly to avoid Chinese text, bullets, and symbols becoming mojibake.
7. Render with PowerShell:

```powershell
.\scripts\Render-PptScene.ps1 `
  -SceneJson .\output\example\scene.json `
  -OutputPptx .\output\example\example.pptx `
  -AssetRoot .\output\example
```

By default the renderer suppresses PowerPoint alerts, marks the presentation saved, closes the generated deck, and quits the COM PowerPoint instance to avoid AutoSave/Save prompts. When running from outside the skill directory, call the bundled script by absolute path or copy it into the workspace. If PowerPoint COM fails in a sandboxed command with a logon/session error, rerun the same command with escalated permissions.

8. Use live preview when the user wants to see PowerPoint draw the diagram:

```powershell
.\scripts\Render-PptScene.ps1 `
  -SceneJson .\output\example\scene.json `
  -OutputPptx .\output\example\example-live.pptx `
  -AssetRoot .\output\example `
  -LivePreview `
  -StepDelayMs 120
```

`-LivePreview` shows the drawing process but still closes PowerPoint when done. Add `-KeepPowerPointOpen` only when the user explicitly wants the final deck left open for inspection.

Export a preview in the same COM session when visual QA is needed:

```powershell
.\scripts\Render-PptScene.ps1 `
  -SceneJson .\output\example\scene.json `
  -OutputPptx .\output\example\example.pptx `
  -AssetRoot .\output\example `
  -ExportPreview `
  -PreviewPath .\output\example\preview.png `
  -PreviewWidth 1920 `
  -PreviewHeight 1080
```

9. Verify the output.
   - Confirm the PPTX exists.
   - Inspect the PPTX package for `ppt/slides/slide1.xml`.
   - If using images, confirm `ppt/media/*` exists.
   - Prefer `-ExportPreview` over a second PowerPoint COM run for previews.
   - When a preview image is available, use Vision QA to identify missing elements, text mismatches, wrong arrows, layout drift, color drift, and raster/native-element mistakes.

## Vision Analysis & Planning

When the user supplies one or more images, or when a PPT preview image is available:

- Read `references/vision-planning-workflow.md`.
- Produce `vision-analysis.md` before writing `scene.json` for image recreation or image composition tasks.
- Identify visual elements, spatial layout, text labels, connectors, hierarchy, style, native-PPT mapping, generated asset candidates, and fidelity risks.
- For multi-image composition, keep source provenance for each selected element and then define one unified visual system.
- For preview QA, compare the preview against the planned structure and produce specific scene-level fixes.
- Do not treat Vision output as exact geometry. Use it as structured planning, then refine coordinates and styles in `scene.json`.

## Description-Driven Design

When the user gives only a description, or asks for a new figure inspired by scientific/design examples:

- Read `references/design-research-workflow.md`.
- Produce a concise `Figure Design Plan` with goal, audience, core message, layout, visual hierarchy, color/font system, element inventory, native-PPT elements, generated/inserted assets, and render plan.
- Wait for user confirmation before PPT rendering by default. If the user says "directly generate", "no confirmation", or equivalent, proceed and record that assumption.
- If web access is unavailable, state that reference research was not performed and continue from the user's description plus local knowledge.
- Use Vision only when the task also includes image inputs or when inspecting a rendered preview.
- Do not place reference URLs in the final PPT unless the user asks for an appendix or audit trail.

## Image Composition Guidance

For rough reference-image recreation:

- Convert simple visual structures into editable PPT primitives.
- Use `freeform`/`path` for curved arrows, smooth brackets, loops, and Bezier-like connector shapes.
- Prefer EMF/SVG insertion for complex icons that should remain crisp; fall back to generated PNG when Office vector import is unreliable.
- Match layout, relative proportions, color families, labels, arrows, and hierarchy.
- Avoid tracing every pixel; create a clean PowerPoint diagram that communicates the same structure.
- Use generated local image assets for complex pictograms or decorative fragments, then insert them with `image`.

For multi-image composition:

- Normalize all source elements to one canvas coordinate system.
- Resolve style conflicts deliberately: one palette, one font system, consistent stroke widths.
- Keep provenance in element ids when helpful, such as `img1-grid`, `img2-arrow-style`, `generated-icon-reactor`.

For repeated figure grammar:

- Use `component.stepCircle` for numbered workflow steps.
- Use `component.skillCard` for repeated skill/library cards.
- Use `component.legendItem` for legend rows.
- Use `component.bracket` and `component.flowArrow` for recurring framework connectors.

## JSON Reference

Read `references/scene-json-v1.md` when writing or debugging scene JSON. Load it especially when adding a new scene type, handling image assets, or diagnosing PowerPoint COM range/style errors.

## Smoke Test

To test the skill's bundled renderer from the skill directory:

```powershell
.\scripts\Render-PptScene.ps1 `
  -SceneJson .\assets\sample\scene.basic.json `
  -OutputPptx .\assets\sample\output\basic-grid.pptx `
  -AssetRoot .\assets\sample
```

Expected result: a one-slide editable PowerPoint grid diagram with an inserted PNG icon.
