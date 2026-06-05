# Scene JSON v1 Reference

Use this reference when authoring scenes for `scripts/Render-PptScene.ps1`.

## Top-Level Shape

```json
{
  "canvas": { "width": 1280, "height": 720 },
  "page": { "width": 960, "height": 540 },
  "theme": {
    "colors": {
      "primary": "#006BD6",
      "accent": "#D000FF"
    },
    "fonts": {
      "body": "Arial"
    }
  },
  "slides": [
    {
      "background": "#F7F9FC",
      "elements": []
    }
  ]
}
```

- `canvas`: logical coordinate space. Default to `1280 x 720` for 16:9 slides.
- `page`: optional PowerPoint page size in points. Default is `960 x 540`; use this for portrait or paper-like figures, for example `{ "width": 675, "height": 950 }`.
- `theme`: optional reusable colors, fonts, and stroke widths. Use references like `theme.colors.primary`, `theme.fonts.body`, or `theme.strokeWidths.normal`.
- `slides`: array of slides.
- `background`: optional slide background color.
- `elements`: array of drawable objects. Elements render in order; later elements appear above earlier ones.

Scene files must be UTF-8. Colors must be `#RRGGBB`, `none`, or `transparent`.

## Coordinate System

- **Origin: top-left** of the canvas (PowerPoint convention).
- **X grows to the right; Y grows downward.**
- All `x`, `y`, `w`, `h`, `x1/y1/x2/y2`, and `points[].x/y` are in **logical canvas units** (the values in `canvas.width` / `canvas.height`).
- The renderer maps logical units → PowerPoint points using the `page` size at render time, so layouts written for `1280 × 720` keep their proportions when `page` is `960 × 540` (default) or anything else.
- `fontSize` is in **points** (PowerPoint native), not logical units. See *Typography Scale* below for the scaling rule when canvas size changes.

## Typography Scale

Default canvas is `1280 × 720`. The following font sizes are the **floors** at that canvas; do not go smaller without a specific reason. Picking a smaller value than the floor is the most common defect — when in doubt, choose the higher end of each range.

| Role | fontSize (pt) | Notes |
| --- | --- | --- |
| Slide title | **36 – 44** | One per slide, bold |
| Section / region title | **24 – 30** | Subgroup headers, swimlane titles |
| Node / card label | **18 – 22** | Text inside flow boxes, cards, chips |
| Body / description | **16 – 18** | Paragraph text, longer captions |
| Annotation / legend | **14 – 16** | **Lower bound — do not go below 14** |
| Source / footer | **12 – 14** | Citations, page footers only |

Rules that override the table:

- **If a node box has `w < 200` or `h < 60`, its label fontSize must be ≥ 18.** Small boxes still need readable labels.
- **If text would overflow, enlarge the container first; do not shrink the font.**
- **Scaling for non-default canvas sizes:** multiply the table values by `canvas.width / 1280`. For example, on a `1920 × 1080` canvas a node label is `18 × 1.5 = 27` pt minimum.
- For portrait/poster canvases, use the longer side as the scaling reference (`max(canvas.width, canvas.height) / 1280`).
- `bold: true` for titles and one-word emphasis only; do not bold body text by default.

For a portrait figure, use matching ratios:

```json
{
  "canvas": { "width": 675, "height": 950 },
  "page": { "width": 675, "height": 950 },
  "slides": []
}
```

## Common Style Fields

- `fill`: shape fill color. Use `none` or `transparent` for no fill.
- `fillOpacity`: number from `0` to `1`.
- `stroke`: outline/line color. Use `none` or `transparent` for no stroke.
- `strokeWidth`: line width in logical pixels.
- `dash`: `solid`, `dash`, `dot`, `dashDot`, `dashDotDot`, or `longDash`.
- `arrowEnd`: `true`, `triangle`, `open`, `stealth`, `diamond`, `oval`, or `none`.
- `arrowBegin`: same values as `arrowEnd`.
- `fontSize`, `fontFace`, `color`, `bold`, `align`, `verticalAlign`, `margin`: text styling fields.
- `zIndex`: optional numeric draw order. Lower values draw first; ties preserve source order.
- `rotation`: optional degrees.
- `gradient`: optional fill gradient `{ "from": "#FFFFFF", "to": "#006BD6", "direction": "horizontal" }`.
- `strokeOpacity`: number from `0` to `1`.
- `shadow`: `true` or `{ "color": "#000000", "opacity": 0.25, "blur": 4, "offsetX": 2, "offsetY": 2 }`.

Text alignment:

- `align`: `left`, `center`, `right`, or `justify`.
- `verticalAlign`: `top`, `middle`, or `bottom`.

## Element Types

### `rect`

```json
{
  "id": "box-1",
  "type": "rect",
  "x": 100,
  "y": 120,
  "w": 180,
  "h": 80,
  "fill": "#EAF2FF",
  "stroke": "#3A6EA5",
  "strokeWidth": 2,
  "radius": 10,
  "text": "Input A",
  "fontSize": 20,
  "bold": true,
  "color": "#1F3B57"
}
```

Use `radius > 0` for rounded rectangles. `fontSize` follows the *Typography Scale* table above (node labels: 18–22 pt at 1280×720).

### `ellipse`

```json
{
  "id": "node-1",
  "type": "ellipse",
  "x": 720,
  "y": 160,
  "w": 110,
  "h": 110,
  "fill": "#FFF3D6",
  "stroke": "#C88719",
  "strokeWidth": 2,
  "text": "Node"
}
```

Equal `w` and `h` creates a circle; unequal values create an ellipse.

### `line`

```json
{
  "id": "arrow-1",
  "type": "line",
  "x1": 270,
  "y1": 191,
  "x2": 300,
  "y2": 191,
  "stroke": "#1F2933",
  "strokeWidth": 2.5,
  "arrowEnd": true
}
```

Use `arrowBegin` and `arrowEnd` only on line-like elements.

### `polyline`

```json
{
  "id": "feedback",
  "type": "polyline",
  "points": [
    { "x": 555, "y": 346 },
    { "x": 555, "y": 408 },
    { "x": 195, "y": 408 },
    { "x": 195, "y": 346 }
  ],
  "stroke": "#9A6A00",
  "strokeWidth": 2,
  "dash": "dash",
  "arrowEnd": true
}
```

`points` must contain at least two points.

### `freeform`, `path`, or `bezier`

Use this for editable curved paths and Bezier-like lines.

```json
{
  "id": "smooth-curve",
  "type": "freeform",
  "nodes": [
    { "x": 100, "y": 200 },
    {
      "type": "curve",
      "c1": { "x": 180, "y": 120 },
      "c2": { "x": 260, "y": 280 },
      "x": 340,
      "y": 200
    },
    { "type": "line", "x": 420, "y": 220 }
  ],
  "stroke": "theme.colors.primary",
  "strokeWidth": 2,
  "fill": "none"
}
```

Set `"closed": true` to close and optionally fill a path.

### `text`

```json
{
  "id": "title",
  "type": "text",
  "x": 70,
  "y": 34,
  "w": 800,
  "h": 44,
  "text": "Grid Structure Demo",
  "fontSize": 40,
  "fontFace": "Arial",
  "bold": true,
  "color": "#1F2933",
  "align": "left"
}
```

Use explicit `w` and `h` large enough for the text; PowerPoint will wrap text inside the text box. Slide titles should be 36–44 pt at 1280×720 — see *Typography Scale* above.

### `image`

```json
{
  "id": "sample-icon",
  "type": "image",
  "x": 898,
  "y": 166,
  "w": 120,
  "h": 120,
  "path": "assets/generated/sample-icon.png"
}
```

Relative image paths resolve against `-AssetRoot`. For example, with `-AssetRoot .\output\example`, the path above resolves to `.\output\example\assets\generated\sample-icon.png`.

### `svg`, `emf`, or `vector`

These are aliases for image insertion and are intended for vector files such as `.svg` and `.emf`.

```json
{
  "id": "vector-icon",
  "type": "svg",
  "x": 100,
  "y": 100,
  "w": 80,
  "h": 80,
  "path": "assets/icon.svg"
}
```

PowerPoint's native import support decides whether the inserted object remains vector-editable. EMF is generally the safer choice for classic editable Office vectors.

### `group`

Groups nested elements and applies an optional local `x` / `y` offset.

```json
{
  "id": "legend-group",
  "type": "group",
  "x": 40,
  "y": 40,
  "elements": [
    { "id": "box", "type": "rect", "x": 0, "y": 0, "w": 16, "h": 16, "fill": "#FFFFFF" },
    { "id": "label", "type": "text", "x": 24, "y": -2, "w": 180, "h": 22, "text": "Legend item" }
  ]
}
```

Use `"ungrouped": true` if local offset is useful but a PowerPoint group is not wanted.

### `component`

Built-in reusable components reduce boilerplate for repeated diagram parts:

```json
{ "id": "step-1", "type": "component", "name": "stepCircle", "x": 40, "y": 100, "size": 22, "text": "1" }
```

Supported `name` values:

- `stepCircle`: numbered circle.
- `skillCard`: rounded card with `title` and `body`.
- `legendItem`: swatch plus label.
- `bracket`: left/right/top/bottom bracket.
- `flowArrow`: line or polyline arrow.

## Error Diagnosis

- AutoSave or Save prompts after generation: use the current renderer, which closes by default; do not pass `-KeepPowerPointOpen` unless the user wants the deck left open.
- `SceneJson was not found`: wrong scene path.
- `SceneJson must be valid UTF-8`: resave the JSON as UTF-8.
- `AssetRoot was not found`: wrong asset root path.
- `Image file for element ... was not found`: fix `image.path` or `-AssetRoot`.
- `unsupported type`: use supported types such as `rect`, `ellipse`, `line`, `polyline`, `freeform`, `text`, `image`, `svg`, `emf`, `group`, or `component`.
- PowerPoint says a value is out of range: check negative dimensions, invalid colors, extreme coordinates, or arrow fields on non-line objects.
- Generated PPTX has no `ppt/slides/slide1.xml`: check that `slides` is an array and the JSON parses correctly.

## Current Limits

The renderer supports freeform Bezier-style paths, grouping, theme references, gradients, transparency, shadows, and SVG/EMF insertion. It does not parse raw SVG path `d` strings; use an SVG/EMF file for complex vector icons or express paths as `freeform.nodes`.

## Useful Renderer Flags

- `-LivePreview`: show PowerPoint while drawing each shape.
- `-StepDelayMs`: delay between drawn elements during live preview.
- `-KeepPowerPointOpen`: leave PowerPoint open after saving. Avoid this in batch runs because it can create Save/AutoSave prompts later.
- `-ExportPreview`: export slide 1 to PNG during the same COM session.
- `-PreviewPath`, `-PreviewWidth`, `-PreviewHeight`: control preview export.
