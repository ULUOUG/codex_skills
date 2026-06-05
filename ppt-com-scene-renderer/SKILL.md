---
name: ppt-com-scene-renderer
description: Use whenever the user wants an editable PowerPoint (.pptx) figure, diagram, or slide — including "给我/输出/生成/做一个可编辑的 PPT"、"画一张 PPT 图"、"PPT 示意图/流程图/架构图/网络图/框架图/概念图"、"editable PPT/PPTX"、"make a PowerPoint diagram"、"draw this in PPT". Produces a real .pptx via JSON scene + PowerPoint COM, so every shape, arrow, and label stays editable in PowerPoint. Also covers recreating a reference image as a PPT figure, composing elements from multiple images into one slide, designing an original scientific or design-style figure from a text description (with optional reference research when the user explicitly asks for it), inserting generated raster/SVG/EMF assets where native PPT shapes fall short, QA-ing a rendered preview image, and live-drawing mode where the user watches PowerPoint build the slide. SKIP for plain text-only decks with no diagram intent, for read-only image export (PNG/SVG only with no .pptx requested), or when the user explicitly asks to edit an existing .pptx file in place.
---

# PPT COM Scene Renderer

## Overview

Use this skill to turn visual intent into an editable `.pptx` by writing a JSON scene and rendering it with PowerPoint COM. Prefer native PPT shapes for geometry, lines, arrows, labels, and grids; insert local PNG/JPG/SVG/EMF assets only for details that are impractical to draw with PowerPoint primitives.

Supported input modes:

- Single-image recreation: reproduce the supplied reference as an editable PPT figure.
- Multi-image composition: take selected elements from several images and unify them into one figure.
- Description-driven original design: start from a text description, create a Figure Design Plan, optionally research visual inspiration **only when the user explicitly asks for it**, then render the final editable PPT.
- Preview QA: inspect a rendered preview image and identify corrections before finalizing the PPT.

Bundled resources:

- `scripts/Render-PptScene.ps1`: JSON scene to PPTX renderer.
- `references/scene-json-v1.md`: supported JSON schema and examples.
- `references/design-research-workflow.md`: description-driven planning, reference research, originality, and handoff templates.
- `references/vision-planning-workflow.md`: Vision-assisted image decomposition, planning, and preview QA templates.
- `assets/sample/`: sample scene and PNG asset for smoke testing.

## Workflow

1. Classify the input mode.
   - Supplied image → single-image recreation.
   - Multiple supplied images → multi-image composition.
   - Text-only request / original figure → description-driven design.
   - Rendered preview image → preview QA.
2. If images are involved, run Vision analysis first.
   - Single-image / multi-image / preview QA: read `references/vision-planning-workflow.md` and produce `vision-analysis.md` (one analysis per source image for multi-image tasks) before writing `scene.json`.
   - Vision only analyzes and plans; it does not render PPTX.
3. If the task is description-driven, read `references/design-research-workflow.md` and produce a `Figure Design Plan` (`design-plan.md`) before writing `scene.json`.
   - **Reference research (Nature / Science / design galleries) is opt-in, not default.** Trigger it only when the user explicitly says one of: "参考 / 灵感 / Nature 风 / Science 风 / 设计网站 / benchmark / find inspiration / look up references" or equivalent. Otherwise skip search and design from the user's description plus local knowledge.
   - When research is triggered, extract only abstract patterns (hierarchy, layout grammar, color role, annotation strategy); never copy figures, icons, palettes, or layouts. Keep `inspiration-notes.md` with URLs and how the final design differs.
   - If the user opted into a design plan, **do not call the renderer until the plan is approved**, unless the user said "直接生成 / no confirmation / skip plan".
4. Preserve editability by default.
   - Use PowerPoint-native shapes, text, lines, curves, groups, components, gradients, transparency, and theme references wherever practical.
   - Generate or insert local PNG/SVG/EMF assets only for complex pictograms, equipment images, textures, and details that are not practical as native PPT shapes.
5. Create a task folder under the project's `output/` directory.
   - Path: `output/<slug>/`, where `<slug>` is a short kebab-case name derived from the user's request (e.g. `agent-workflow`, `chemcrow-figure`). Do **not** place artifacts in the project root.
   - `output/<slug>/vision-analysis.md` (when applicable)
   - `output/<slug>/design-plan.md` (when applicable)
   - `output/<slug>/inspiration-notes.md` (when applicable)
   - `output/<slug>/scene.json`
   - `output/<slug>/assets/` for generated or extracted assets
   - `output/<slug>/<slug>.pptx` for the final deck
   - `output/<slug>/preview.png` (or `preview/` for multi-page exports)
6. Write `scene.json` using the v1 schema (`references/scene-json-v1.md`).
   - Default canvas: `1280 x 720`. For portrait figures set both `canvas` and `page` to the desired ratio.
   - Draw elements in z-order: lower layers first, foreground labels and icons last.
   - Use `rect`, `ellipse`, `line`, `polyline`, `freeform`, `path`, `text` for editable content; `svg`/`emf`/`vector` for local vector files; `image` for raster.
   - Use `group`, `zIndex`, theme references, and built-in `component` elements for repeated structure.
   - **Follow the Typography Scale rules in `references/scene-json-v1.md` — font sizes must match canvas size; do not default to small fonts.**
   - Save JSON as UTF-8; the renderer reads UTF-8 strictly to avoid Chinese text, bullets, and symbols becoming mojibake.
7. Render with PowerShell (only after step 3's confirmation gate, if applicable):

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
   - Confirm the `.pptx` file exists at the expected path.
   - Prefer running with `-ExportPreview` and visually inspecting `preview.png` over zip inspection.
   - Only when rendering fails or you suspect a packaging issue, inspect the PPTX package:
     ```powershell
     Add-Type -AssemblyName System.IO.Compression.FileSystem
     [IO.Compression.ZipFile]::OpenRead('output\<slug>\<slug>.pptx').Entries | Select-Object FullName
     ```
     Expect `ppt/slides/slide1.xml` (and `ppt/media/*` if images were used).
   - When a preview image is available, run Vision QA against the reference image or design plan: list missing elements, text mismatches, wrong arrows, layout drift, color drift, and font-size mismatches before rerendering.

## Reference Pointers

Detailed rules live in the references; only consult them on demand:

- `references/vision-planning-workflow.md` — image / preview tasks (analysis template, mapping rules, preview QA).
- `references/design-research-workflow.md` — description-driven tasks (Figure Design Plan template, opt-in reference-research rules, originality guardrails).
- `references/scene-json-v1.md` — JSON schema, coordinate system, **typography scale**, component catalog, error diagnosis.

Cross-cutting rules:

- Reference research (Nature/Science/design galleries) is **opt-in**; trigger keywords listed in step 3 above.
- Do not place reference URLs in the final PPT unless the user asks for an appendix.
- Vision outputs are structural planning, not exact geometry — refine coordinates in `scene.json`.
- For description-driven tasks, the renderer runs **after** plan approval (step 7 gate).

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

## Smoke Test

From the skill directory (`skills/ppt-com-scene-renderer/`):

```powershell
.\scripts\Render-PptScene.ps1 `
  -SceneJson .\assets\sample\scene.basic.json `
  -OutputPptx .\assets\sample\output\basic-grid.pptx `
  -AssetRoot .\assets\sample
```

Expected result: a one-slide editable PowerPoint grid diagram with an inserted PNG icon. From any other working directory, call the script by absolute path.
