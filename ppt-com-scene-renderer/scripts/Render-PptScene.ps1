[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SceneJson,

    [Parameter(Mandatory = $true)]
    [string]$OutputPptx,

    [string]$AssetRoot,

    [switch]$LivePreview,

    [int]$StepDelayMs = 120,

    [switch]$KeepPowerPointOpen,

    [switch]$ExportPreview,

    [string]$PreviewPath,

    [int]$PreviewWidth = 1920,

    [int]$PreviewHeight = 1080
)

$ErrorActionPreference = "Stop"

function Get-JsonProp {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = $null
    )

    if ($null -eq $Object) {
        return $Default
    }

    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) {
        return $Default
    }

    return $prop.Value
}

function Test-JsonProp {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $Object) {
        return $false
    }

    return $null -ne $Object.PSObject.Properties[$Name]
}

function Get-RequiredProp {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Context
    )

    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) {
        throw "$Context is missing required property '$Name'."
    }

    return $prop.Value
}

function ConvertTo-Array {
    param($Value)

    if ($null -eq $Value) {
        return ,([object[]]@())
    }

    return ,([object[]]@($Value))
}

function ConvertTo-RgbInt {
    param(
        $Color,
        [string]$Default = "#000000"
    )

    if ($null -eq $Color -or [string]::IsNullOrWhiteSpace([string]$Color)) {
        $Color = $Default
    }

    $text = Resolve-ThemeValue ([string]$Color)
    $text = ([string]$text).Trim()
    if ($text -eq "none" -or $text -eq "transparent") {
        return $null
    }

    if ($text -notmatch '^#?([0-9a-fA-F]{6})$') {
        throw "Invalid color '$text'. Use #RRGGBB, 'none', or 'transparent'."
    }

    $hex = $Matches[1]
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)

    return $r + ($g * 256) + ($b * 65536)
}

function Resolve-ThemeValue {
    param($Value)

    if ($null -eq $Value) {
        return $Value
    }

    $text = [string]$Value
    if (-not $text.StartsWith("theme.")) {
        return $Value
    }

    $parts = $text.Split(".")
    if ($parts.Count -lt 3) {
        throw "Invalid theme reference '$text'. Use theme.colors.name, theme.fonts.name, or theme.strokeWidths.name."
    }

    $current = $script:Theme
    for ($i = 1; $i -lt $parts.Count; $i++) {
        if ($null -eq $current -or $null -eq $current.PSObject.Properties[$parts[$i]]) {
            throw "Theme reference '$text' was not found."
        }
        $current = $current.PSObject.Properties[$parts[$i]].Value
    }

    return $current
}

function ConvertTo-PptX {
    param([double]$Value)
    return ($Value + $script:OffsetX) * $script:ScaleX
}

function ConvertTo-PptY {
    param([double]$Value)
    return ($Value + $script:OffsetY) * $script:ScaleY
}

function ConvertTo-PptW {
    param([double]$Value)
    return $Value * $script:ScaleX
}

function ConvertTo-PptH {
    param([double]$Value)
    return $Value * $script:ScaleY
}

function ConvertTo-PptLineWeight {
    param($Value, [double]$Default = 1)

    if ($null -eq $Value) {
        $Value = $Default
    }
    else {
        $Value = Resolve-ThemeValue $Value
    }

    return [double]$Value * $script:StrokeScale
}

function ConvertTo-GradientStyle {
    param($Direction)

    $text = if ($null -eq $Direction) { "horizontal" } else { ([string]$Direction).ToLowerInvariant() }
    switch ($text) {
        "vertical" { return 2 }
        "diagonalup" { return 3 }
        "diagonal-up" { return 3 }
        "diagonaldown" { return 4 }
        "diagonal-down" { return 4 }
        default { return 1 }
    }
}

function ConvertTo-ParagraphAlignment {
    param($Align)

    $text = if ($null -eq $Align) { "left" } else { ([string]$Align).ToLowerInvariant() }
    switch ($text) {
        "center" { return 2 }
        "right" { return 3 }
        "justify" { return 4 }
        default { return 1 }
    }
}

function ConvertTo-DashStyle {
    param($Dash)

    if ($null -eq $Dash -or $Dash -eq $false) {
        return 1
    }

    if ($Dash -eq $true) {
        return 4
    }

    switch (([string]$Dash).ToLowerInvariant()) {
        "solid" { return 1 }
        "squareDot" { return 2 }
        "squaredot" { return 2 }
        "dot" { return 3 }
        "roundDot" { return 3 }
        "rounddot" { return 3 }
        "dash" { return 4 }
        "dashDot" { return 5 }
        "dashdot" { return 5 }
        "dashDotDot" { return 6 }
        "dashdotdot" { return 6 }
        "longDash" { return 7 }
        "longdash" { return 7 }
        default { return 4 }
    }
}

function ConvertTo-ArrowheadStyle {
    param($Arrow)

    if ($null -eq $Arrow -or $Arrow -eq $false -or ([string]$Arrow).ToLowerInvariant() -eq "none") {
        return 1
    }

    if ($Arrow -eq $true) {
        return 2
    }

    switch (([string]$Arrow).ToLowerInvariant()) {
        "triangle" { return 2 }
        "open" { return 3 }
        "stealth" { return 4 }
        "diamond" { return 5 }
        "oval" { return 6 }
        default { return 2 }
    }
}

function Apply-FillStyle {
    param(
        [Parameter(Mandatory = $true)]$Shape,
        [Parameter(Mandatory = $true)]$Element
    )

    $gradient = Get-JsonProp $Element "gradient" $null
    if ($null -ne $gradient) {
        $from = ConvertTo-RgbInt (Get-JsonProp $gradient "from" "#FFFFFF") "#FFFFFF"
        $to = ConvertTo-RgbInt (Get-JsonProp $gradient "to" "#FFFFFF") "#FFFFFF"
        if ($null -eq $from -or $null -eq $to) {
            $Shape.Fill.Visible = 0
            return
        }

        $Shape.Fill.Visible = -1
        $Shape.Fill.ForeColor.RGB = $from
        $Shape.Fill.BackColor.RGB = $to
        $Shape.Fill.TwoColorGradient((ConvertTo-GradientStyle (Get-JsonProp $gradient "direction" "horizontal")), [int](Get-JsonProp $gradient "variant" 1))
        $opacity = Get-JsonProp $Element "fillOpacity" (Get-JsonProp $gradient "opacity" 1)
        $Shape.Fill.Transparency = 1 - [double]$opacity
        return
    }

    $fill = Get-JsonProp $Element "fill" "#FFFFFF"
    $rgb = ConvertTo-RgbInt $fill "#FFFFFF"

    if ($null -eq $rgb) {
        $Shape.Fill.Visible = 0
        return
    }

    $Shape.Fill.Visible = -1
    $Shape.Fill.ForeColor.RGB = $rgb
    $opacity = Get-JsonProp $Element "fillOpacity" 1
    $Shape.Fill.Transparency = 1 - [double]$opacity
}

function Apply-LineStyle {
    param(
        [Parameter(Mandatory = $true)]$LineFormat,
        [Parameter(Mandatory = $true)]$Element
    )

    $strokeWidth = Get-JsonProp $Element "strokeWidth" 1
    $stroke = Get-JsonProp $Element "stroke" "#000000"
    $rgb = ConvertTo-RgbInt $stroke "#000000"

    if ($null -eq $rgb -or [double]$strokeWidth -le 0) {
        $LineFormat.Visible = 0
        return
    }

    $LineFormat.Visible = -1
    $LineFormat.ForeColor.RGB = $rgb
    $LineFormat.Weight = ConvertTo-PptLineWeight $strokeWidth
    $LineFormat.DashStyle = ConvertTo-DashStyle (Get-JsonProp $Element "dash" $null)
    $strokeOpacity = Get-JsonProp $Element "strokeOpacity" 1
    try {
        $LineFormat.Transparency = 1 - [double]$strokeOpacity
    }
    catch {
        Write-Warning "Could not apply strokeOpacity: $($_.Exception.Message)"
    }

    if ($null -ne $Element.PSObject.Properties["arrowEnd"]) {
        $LineFormat.EndArrowheadStyle = ConvertTo-ArrowheadStyle (Get-JsonProp $Element "arrowEnd" $false)
    }
    if ($null -ne $Element.PSObject.Properties["arrowBegin"]) {
        $LineFormat.BeginArrowheadStyle = ConvertTo-ArrowheadStyle (Get-JsonProp $Element "arrowBegin" $false)
    }
}

function Apply-ShapeCommon {
    param(
        [Parameter(Mandatory = $true)]$Shape,
        [Parameter(Mandatory = $true)]$Element
    )

    if (Test-JsonProp $Element "rotation") {
        try {
            $Shape.Rotation = [double](Get-JsonProp $Element "rotation" 0)
        }
        catch {
            Write-Warning "Could not apply rotation: $($_.Exception.Message)"
        }
    }

    $shadow = Get-JsonProp $Element "shadow" $null
    if ($null -ne $shadow -and $shadow -ne $false) {
        try {
            $Shape.Shadow.Visible = -1
            if ($shadow -isnot [bool]) {
                if (Test-JsonProp $shadow "color") {
                    $Shape.Shadow.ForeColor.RGB = ConvertTo-RgbInt (Get-JsonProp $shadow "color" "#000000") "#000000"
                }
                if (Test-JsonProp $shadow "opacity") {
                    $Shape.Shadow.Transparency = 1 - [double](Get-JsonProp $shadow "opacity" 0.25)
                }
                if (Test-JsonProp $shadow "blur") {
                    $Shape.Shadow.Blur = [double](Get-JsonProp $shadow "blur" 4)
                }
                if (Test-JsonProp $shadow "offsetX") {
                    $Shape.Shadow.OffsetX = ConvertTo-PptW ([double](Get-JsonProp $shadow "offsetX" 2))
                }
                if (Test-JsonProp $shadow "offsetY") {
                    $Shape.Shadow.OffsetY = ConvertTo-PptH ([double](Get-JsonProp $shadow "offsetY" 2))
                }
            }
        }
        catch {
            Write-Warning "Could not apply shadow: $($_.Exception.Message)"
        }
    }
}

function Apply-TextStyle {
    param(
        [Parameter(Mandatory = $true)]$Shape,
        [Parameter(Mandatory = $true)]$Element,
        [Parameter(Mandatory = $true)]$Text,
        [string]$DefaultAlign = "center",
        [string]$DefaultVerticalAlign = "middle"
    )

    $Shape.TextFrame.TextRange.Text = [string]$Text
    $Shape.TextFrame.WordWrap = -1

    $font = $Shape.TextFrame.TextRange.Font
    $font.Name = [string](Resolve-ThemeValue (Get-JsonProp $Element "fontFace" "Arial"))
    $font.Size = [double](Get-JsonProp $Element "fontSize" 14) * $script:TextScale
    $font.Bold = if ([bool](Get-JsonProp $Element "bold" $false)) { -1 } else { 0 }

    $color = ConvertTo-RgbInt (Get-JsonProp $Element "color" "#1F2933") "#1F2933"
    if ($null -ne $color) {
        $font.Color.RGB = $color
    }

    $align = Get-JsonProp $Element "align" $DefaultAlign
    $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = ConvertTo-ParagraphAlignment $align

    $verticalAlign = ([string](Get-JsonProp $Element "verticalAlign" $DefaultVerticalAlign)).ToLowerInvariant()
    switch ($verticalAlign) {
        "top" { $Shape.TextFrame.VerticalAnchor = 1 }
        "bottom" { $Shape.TextFrame.VerticalAnchor = 4 }
        default { $Shape.TextFrame.VerticalAnchor = 3 }
    }

    $margin = [double](Get-JsonProp $Element "margin" 4) * $script:StrokeScale
    $Shape.TextFrame.MarginLeft = $margin
    $Shape.TextFrame.MarginRight = $margin
    $Shape.TextFrame.MarginTop = $margin
    $Shape.TextFrame.MarginBottom = $margin
}

function Invoke-LiveStep {
    param($Shape)

    if (-not $LivePreview) {
        return
    }

    if ($null -ne $Shape) {
        try {
            $Shape.Select()
        }
        catch {
            # Some objects cannot be selected during fast redraw; drawing still succeeded.
        }
    }

    if ($StepDelayMs -gt 0) {
        Start-Sleep -Milliseconds $StepDelayMs
    }
}

function Resolve-AssetPath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ElementId
    )

    $candidate = if ([System.IO.Path]::IsPathRooted($Path)) {
        $Path
    }
    else {
        Join-Path $script:ResolvedAssetRoot $Path
    }

    if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        throw "Image file for element '$ElementId' was not found: $candidate"
    }

    return (Resolve-Path -LiteralPath $candidate).Path
}

function Add-RectangleElement {
    param($Slide, $Element, [string]$Context)

    $x = ConvertTo-PptX ([double](Get-RequiredProp $Element "x" $Context))
    $y = ConvertTo-PptY ([double](Get-RequiredProp $Element "y" $Context))
    $w = ConvertTo-PptW ([double](Get-RequiredProp $Element "w" $Context))
    $h = ConvertTo-PptH ([double](Get-RequiredProp $Element "h" $Context))
    $radius = [double](Get-JsonProp $Element "radius" 0)
    $shapeType = if ($radius -gt 0) { 5 } else { 1 }

    $shape = $Slide.Shapes.AddShape($shapeType, $x, $y, $w, $h)
    Apply-FillStyle $shape $Element
    Apply-LineStyle $shape.Line $Element
    Apply-ShapeCommon $shape $Element

    if ($radius -gt 0) {
        try {
            $maxRadius = [Math]::Max([Math]::Min([double](Get-RequiredProp $Element "w" $Context), [double](Get-RequiredProp $Element "h" $Context)), 1)
            $shape.Adjustments.Item(1) = [Math]::Min($radius / $maxRadius, 0.5)
        }
        catch {
            # Rounded rectangle adjustment support varies across Office versions.
        }
    }

    $text = Get-JsonProp $Element "text" $null
    if ($null -ne $text -and [string]$text -ne "") {
        Apply-TextStyle $shape $Element $text "center" "middle"
    }

    return $shape
}

function Add-EllipseElement {
    param($Slide, $Element, [string]$Context)

    $x = ConvertTo-PptX ([double](Get-RequiredProp $Element "x" $Context))
    $y = ConvertTo-PptY ([double](Get-RequiredProp $Element "y" $Context))
    $w = ConvertTo-PptW ([double](Get-RequiredProp $Element "w" $Context))
    $h = ConvertTo-PptH ([double](Get-RequiredProp $Element "h" $Context))

    $shape = $Slide.Shapes.AddShape(9, $x, $y, $w, $h)
    Apply-FillStyle $shape $Element
    Apply-LineStyle $shape.Line $Element
    Apply-ShapeCommon $shape $Element

    $text = Get-JsonProp $Element "text" $null
    if ($null -ne $text -and [string]$text -ne "") {
        Apply-TextStyle $shape $Element $text "center" "middle"
    }

    return $shape
}

function Add-LineElement {
    param($Slide, $Element, [string]$Context)

    $x1 = ConvertTo-PptX ([double](Get-RequiredProp $Element "x1" $Context))
    $y1 = ConvertTo-PptY ([double](Get-RequiredProp $Element "y1" $Context))
    $x2 = ConvertTo-PptX ([double](Get-RequiredProp $Element "x2" $Context))
    $y2 = ConvertTo-PptY ([double](Get-RequiredProp $Element "y2" $Context))

    $shape = $Slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    Apply-LineStyle $shape.Line $Element
    Apply-ShapeCommon $shape $Element
    return $shape
}

function Add-PolylineElement {
    param($Slide, $Element, [string]$Context)

    $points = ConvertTo-Array (Get-RequiredProp $Element "points" $Context)
    if ($points.Count -lt 2) {
        throw "$Context polyline requires at least two points."
    }

    $first = $points[0]
    $x0 = ConvertTo-PptX ([double](Get-RequiredProp $first "x" "$Context point 0"))
    $y0 = ConvertTo-PptY ([double](Get-RequiredProp $first "y" "$Context point 0"))
    $builder = $Slide.Shapes.BuildFreeform(1, $x0, $y0)

    for ($i = 1; $i -lt $points.Count; $i++) {
        $point = $points[$i]
        $x = ConvertTo-PptX ([double](Get-RequiredProp $point "x" "$Context point $i"))
        $y = ConvertTo-PptY ([double](Get-RequiredProp $point "y" "$Context point $i"))
        $builder.AddNodes(0, 1, $x, $y)
    }

    $shape = $builder.ConvertToShape()
    $shape.Fill.Visible = 0
    Apply-LineStyle $shape.Line $Element
    Apply-ShapeCommon $shape $Element
    return $shape
}

function Get-PathCoordinate {
    param($Node, [string]$Name, [string]$Context)

    if (Test-JsonProp $Node $Name) {
        return [double](Get-JsonProp $Node $Name 0)
    }

    throw "$Context is missing path coordinate '$Name'."
}

function Add-FreeformElement {
    param($Slide, $Element, [string]$Context)

    $nodes = ConvertTo-Array (Get-RequiredProp $Element "nodes" $Context)
    if ($nodes.Count -lt 2) {
        throw "$Context freeform/path requires at least two nodes."
    }

    $first = $nodes[0]
    $x0 = ConvertTo-PptX (Get-PathCoordinate $first "x" "$Context node 0")
    $y0 = ConvertTo-PptY (Get-PathCoordinate $first "y" "$Context node 0")
    $editingType = switch (([string](Get-JsonProp $Element "editing" "corner")).ToLowerInvariant()) {
        "smooth" { 2 }
        "symmetric" { 3 }
        "auto" { 0 }
        default { 1 }
    }
    $builder = $Slide.Shapes.BuildFreeform($editingType, $x0, $y0)

    for ($i = 1; $i -lt $nodes.Count; $i++) {
        $node = $nodes[$i]
        $nodeType = ([string](Get-JsonProp $node "type" (Get-JsonProp $node "command" "line"))).ToLowerInvariant()
        switch ($nodeType) {
            { $_ -in @("curve", "bezier", "c") } {
                $c1 = Get-JsonProp $node "c1" $null
                $c2 = Get-JsonProp $node "c2" $null
                $c1x = if ($null -ne $c1) { Get-PathCoordinate $c1 "x" "$Context node $i c1" } else { Get-PathCoordinate $node "c1x" "$Context node $i" }
                $c1y = if ($null -ne $c1) { Get-PathCoordinate $c1 "y" "$Context node $i c1" } else { Get-PathCoordinate $node "c1y" "$Context node $i" }
                $c2x = if ($null -ne $c2) { Get-PathCoordinate $c2 "x" "$Context node $i c2" } else { Get-PathCoordinate $node "c2x" "$Context node $i" }
                $c2y = if ($null -ne $c2) { Get-PathCoordinate $c2 "y" "$Context node $i c2" } else { Get-PathCoordinate $node "c2y" "$Context node $i" }
                $x = Get-PathCoordinate $node "x" "$Context node $i"
                $y = Get-PathCoordinate $node "y" "$Context node $i"
                $builder.AddNodes(1, 2, (ConvertTo-PptX $c1x), (ConvertTo-PptY $c1y), (ConvertTo-PptX $c2x), (ConvertTo-PptY $c2y), (ConvertTo-PptX $x), (ConvertTo-PptY $y))
            }
            default {
                $x = Get-PathCoordinate $node "x" "$Context node $i"
                $y = Get-PathCoordinate $node "y" "$Context node $i"
                $builder.AddNodes(0, 1, (ConvertTo-PptX $x), (ConvertTo-PptY $y))
            }
        }
    }

    if ([bool](Get-JsonProp $Element "closed" $false)) {
        $builder.AddNodes(0, 1, $x0, $y0)
    }

    $shape = $builder.ConvertToShape()
    if ([bool](Get-JsonProp $Element "closed" $false) -or (Test-JsonProp $Element "fill") -or (Test-JsonProp $Element "gradient")) {
        Apply-FillStyle $shape $Element
    }
    else {
        $shape.Fill.Visible = 0
    }
    Apply-LineStyle $shape.Line $Element
    Apply-ShapeCommon $shape $Element
    return $shape
}

function Add-TextElement {
    param($Slide, $Element, [string]$Context)

    $x = ConvertTo-PptX ([double](Get-RequiredProp $Element "x" $Context))
    $y = ConvertTo-PptY ([double](Get-RequiredProp $Element "y" $Context))
    $w = ConvertTo-PptW ([double](Get-RequiredProp $Element "w" $Context))
    $h = ConvertTo-PptH ([double](Get-RequiredProp $Element "h" $Context))
    $text = Get-RequiredProp $Element "text" $Context

    $shape = $Slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    Apply-TextStyle $shape $Element $text "left" "top"
    Apply-ShapeCommon $shape $Element
    return $shape
}

function Add-ImageElement {
    param($Slide, $Element, [string]$Context)

    $x = ConvertTo-PptX ([double](Get-RequiredProp $Element "x" $Context))
    $y = ConvertTo-PptY ([double](Get-RequiredProp $Element "y" $Context))
    $w = ConvertTo-PptW ([double](Get-RequiredProp $Element "w" $Context))
    $h = ConvertTo-PptH ([double](Get-RequiredProp $Element "h" $Context))
    $path = [string](Get-RequiredProp $Element "path" $Context)
    $id = [string](Get-JsonProp $Element "id" $Context)
    $resolvedPath = Resolve-AssetPath $path $id

    $shape = $Slide.Shapes.AddPicture($resolvedPath, 0, -1, $x, $y, $w, $h)
    Apply-ShapeCommon $shape $Element
    return $shape
}

function Sort-SceneElements {
    param($Elements)

    $array = ConvertTo-Array $Elements
    $indexed = for ($i = 0; $i -lt $array.Count; $i++) {
        [pscustomobject]@{
            Index = $i
            Z = [double](Get-JsonProp $array[$i] "zIndex" 0)
            Element = $array[$i]
        }
    }
    return ,([object[]](($indexed | Sort-Object -Property Z, Index) | ForEach-Object { $_.Element }))
}

function Group-ShapesByName {
    param($Slide, [object[]]$ShapeNames, [string]$Context)

    $names = @($ShapeNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($names.Count -eq 0) {
        return $null
    }
    if ($names.Count -eq 1) {
        return $Slide.Shapes.Item($names[0])
    }

    try {
        return $Slide.Shapes.Range([object[]]$names).Group()
    }
    catch {
        Write-Warning "Could not group shapes for ${Context}: $($_.Exception.Message)"
        return $Slide.Shapes.Item($names[-1])
    }
}

function Add-GroupElement {
    param($Slide, $Element, [string]$Context)

    $children = Sort-SceneElements (Get-RequiredProp $Element "elements" $Context)
    $oldOffsetX = $script:OffsetX
    $oldOffsetY = $script:OffsetY
    $script:OffsetX = $script:OffsetX + [double](Get-JsonProp $Element "x" 0)
    $script:OffsetY = $script:OffsetY + [double](Get-JsonProp $Element "y" 0)
    $names = @()
    try {
        for ($i = 0; $i -lt $children.Count; $i++) {
            $child = $children[$i]
            $childId = Get-JsonProp $child "id" "child-$($i + 1)"
            $shape = Add-SceneElement $Slide $child 0 ($i + 1)
            if ($null -ne $shape) {
                try {
                    $shape.Name = "$Context/$childId"
                }
                catch {}
                $names += $shape.Name
            }
        }
    }
    finally {
        $script:OffsetX = $oldOffsetX
        $script:OffsetY = $oldOffsetY
    }

    if ([bool](Get-JsonProp $Element "ungrouped" $false)) {
        if ($names.Count -gt 0) {
            return $Slide.Shapes.Item($names[-1])
        }
        return $null
    }

    $group = Group-ShapesByName $Slide $names $Context
    if ($null -ne $group) {
        Apply-ShapeCommon $group $Element
    }
    return $group
}

function New-Element {
    param([hashtable]$Properties)

    return [pscustomobject]$Properties
}

function Add-ComponentElement {
    param($Slide, $Element, [string]$Context)

    $name = ([string](Get-RequiredProp $Element "name" $Context)).ToLowerInvariant()
    switch ($name) {
        "stepcircle" {
            $size = [double](Get-JsonProp $Element "size" 24)
            $circle = New-Element @{
                id = "circle"; type = "ellipse"; x = [double](Get-JsonProp $Element "x" 0); y = [double](Get-JsonProp $Element "y" 0); w = $size; h = $size;
                fill = Get-JsonProp $Element "fill" "theme.colors.primary"; stroke = Get-JsonProp $Element "stroke" "theme.colors.primary"; strokeWidth = Get-JsonProp $Element "strokeWidth" 1;
                text = [string](Get-JsonProp $Element "text" "1"); fontSize = Get-JsonProp $Element "fontSize" 10; bold = Get-JsonProp $Element "bold" $true; color = Get-JsonProp $Element "color" "#FFFFFF"
            }
            return Add-EllipseElement $Slide $circle $Context
        }
        "flowarrow" {
            if (Test-JsonProp $Element "points") {
                $arrow = New-Element @{
                    id = "arrow"; type = "polyline"; points = Get-JsonProp $Element "points" @(); stroke = Get-JsonProp $Element "stroke" "theme.colors.primary";
                    strokeWidth = Get-JsonProp $Element "strokeWidth" 2; dash = Get-JsonProp $Element "dash" $null; arrowEnd = Get-JsonProp $Element "arrowEnd" $true; arrowBegin = Get-JsonProp $Element "arrowBegin" $false
                }
                return Add-PolylineElement $Slide $arrow $Context
            }
            $line = New-Element @{
                id = "arrow"; type = "line"; x1 = Get-JsonProp $Element "x1" 0; y1 = Get-JsonProp $Element "y1" 0; x2 = Get-JsonProp $Element "x2" 0; y2 = Get-JsonProp $Element "y2" 0;
                stroke = Get-JsonProp $Element "stroke" "theme.colors.primary"; strokeWidth = Get-JsonProp $Element "strokeWidth" 2; dash = Get-JsonProp $Element "dash" $null;
                arrowEnd = Get-JsonProp $Element "arrowEnd" $true; arrowBegin = Get-JsonProp $Element "arrowBegin" $false
            }
            return Add-LineElement $Slide $line $Context
        }
        "bracket" {
            $x = [double](Get-JsonProp $Element "x" 0)
            $y = [double](Get-JsonProp $Element "y" 0)
            $w = [double](Get-JsonProp $Element "w" 20)
            $h = [double](Get-JsonProp $Element "h" 100)
            $side = ([string](Get-JsonProp $Element "side" "left")).ToLowerInvariant()
            $points = switch ($side) {
                "right" { @((New-Element @{x=$x; y=$y}), (New-Element @{x=($x+$w); y=$y}), (New-Element @{x=($x+$w); y=($y+$h)}), (New-Element @{x=$x; y=($y+$h)})) }
                "top" { @((New-Element @{x=$x; y=($y+$h)}), (New-Element @{x=$x; y=$y}), (New-Element @{x=($x+$w); y=$y}), (New-Element @{x=($x+$w); y=($y+$h)})) }
                "bottom" { @((New-Element @{x=$x; y=$y}), (New-Element @{x=$x; y=($y+$h)}), (New-Element @{x=($x+$w); y=($y+$h)}), (New-Element @{x=($x+$w); y=$y})) }
                default { @((New-Element @{x=($x+$w); y=$y}), (New-Element @{x=$x; y=$y}), (New-Element @{x=$x; y=($y+$h)}), (New-Element @{x=($x+$w); y=($y+$h)})) }
            }
            $bracket = New-Element @{ id="bracket"; type="polyline"; points=$points; stroke=Get-JsonProp $Element "stroke" "theme.colors.primary"; strokeWidth=Get-JsonProp $Element "strokeWidth" 2 }
            return Add-PolylineElement $Slide $bracket $Context
        }
        "legenditem" {
            $x = [double](Get-JsonProp $Element "x" 0)
            $y = [double](Get-JsonProp $Element "y" 0)
            $s = [double](Get-JsonProp $Element "swatchSize" 16)
            $group = New-Element @{
                id="legend"; type="group"; x=0; y=0; elements=@(
                    (New-Element @{ id="swatch"; type="rect"; x=$x; y=$y; w=$s; h=$s; fill=Get-JsonProp $Element "swatch" (Get-JsonProp $Element "fill" "#FFFFFF"); stroke=Get-JsonProp $Element "stroke" "#000000"; strokeWidth=Get-JsonProp $Element "strokeWidth" 1 }),
                    (New-Element @{ id="label"; type="text"; x=($x+$s+8); y=($y-2); w=Get-JsonProp $Element "textW" 220; h=($s+6); text=Get-JsonProp $Element "text" ""; fontSize=Get-JsonProp $Element "fontSize" 12; fontFace=Get-JsonProp $Element "fontFace" "theme.fonts.body"; color=Get-JsonProp $Element "color" "#000000"; align="left" })
                )
            }
            return Add-GroupElement $Slide $group $Context
        }
        "skillcard" {
            $x = [double](Get-JsonProp $Element "x" 0)
            $y = [double](Get-JsonProp $Element "y" 0)
            $w = [double](Get-JsonProp $Element "w" 150)
            $h = [double](Get-JsonProp $Element "h" 100)
            $group = New-Element @{
                id="skill-card"; type="group"; x=0; y=0; elements=@(
                    (New-Element @{ id="box"; type="rect"; x=$x; y=$y; w=$w; h=$h; fill=Get-JsonProp $Element "fill" "#FFFFFF"; stroke=Get-JsonProp $Element "stroke" "theme.colors.primary"; strokeWidth=Get-JsonProp $Element "strokeWidth" 1.5; radius=Get-JsonProp $Element "radius" 6 }),
                    (New-Element @{ id="title"; type="text"; x=($x+10); y=($y+8); w=($w-20); h=20; text=Get-JsonProp $Element "title" ""; fontSize=Get-JsonProp $Element "titleSize" 11; fontFace=Get-JsonProp $Element "fontFace" "theme.fonts.body"; bold=$true; color=Get-JsonProp $Element "titleColor" (Get-JsonProp $Element "stroke" "theme.colors.primary"); align="left" }),
                    (New-Element @{ id="body"; type="text"; x=($x+10); y=($y+32); w=($w-20); h=($h-38); text=Get-JsonProp $Element "body" ""; fontSize=Get-JsonProp $Element "fontSize" 9; fontFace=Get-JsonProp $Element "fontFace" "theme.fonts.body"; bold=Get-JsonProp $Element "bold" $false; color=Get-JsonProp $Element "color" "#000000"; align="left" })
                )
            }
            return Add-GroupElement $Slide $group $Context
        }
        default {
            throw "$Context uses unsupported component '$name'."
        }
    }
}

function Add-SceneElement {
    param($Slide, $Element, [int]$SlideNumber, [int]$ElementNumber)

    $type = ([string](Get-RequiredProp $Element "type" "slide $SlideNumber element $ElementNumber")).ToLowerInvariant()
    $id = Get-JsonProp $Element "id" "element-$ElementNumber"
    $context = "slide $SlideNumber element '$id'"

    switch ($type) {
        "rect" { return Add-RectangleElement $Slide $Element $context }
        "ellipse" { return Add-EllipseElement $Slide $Element $context }
        "line" { return Add-LineElement $Slide $Element $context }
        "polyline" { return Add-PolylineElement $Slide $Element $context }
        "freeform" { return Add-FreeformElement $Slide $Element $context }
        "path" { return Add-FreeformElement $Slide $Element $context }
        "bezier" { return Add-FreeformElement $Slide $Element $context }
        "text" { return Add-TextElement $Slide $Element $context }
        "image" { return Add-ImageElement $Slide $Element $context }
        "svg" { return Add-ImageElement $Slide $Element $context }
        "emf" { return Add-ImageElement $Slide $Element $context }
        "vector" { return Add-ImageElement $Slide $Element $context }
        "group" { return Add-GroupElement $Slide $Element $context }
        "component" { return Add-ComponentElement $Slide $Element $context }
        default { throw "$context has unsupported type '$type'." }
    }
}

function Resolve-OutputPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = if ([System.IO.Path]::IsPathRooted($Path)) {
        [System.IO.Path]::GetFullPath($Path)
    }
    else {
        [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
    }

    $parent = Split-Path -Parent $fullPath
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    return $fullPath
}

function Set-PowerPointQuietMode {
    param($PowerPoint)

    try {
        # ppAlertsNone = 1. Suppresses Save/AutoSave prompts during COM cleanup.
        $PowerPoint.DisplayAlerts = 1
    }
    catch {
        Write-Warning "Could not set PowerPoint.DisplayAlerts to ppAlertsNone: $($_.Exception.Message)"
    }
}

function Complete-PresentationSaveState {
    param($Presentation)

    if ($null -eq $Presentation) {
        return
    }

    try {
        $Presentation.Saved = -1
    }
    catch {
        Write-Warning "Could not mark presentation as saved: $($_.Exception.Message)"
    }
}

if (-not (Test-Path -LiteralPath $SceneJson -PathType Leaf)) {
    throw "SceneJson was not found: $SceneJson"
}

$resolvedSceneJson = (Resolve-Path -LiteralPath $SceneJson).Path
$sceneDir = Split-Path -Parent $resolvedSceneJson

if ([string]::IsNullOrWhiteSpace($AssetRoot)) {
    $script:ResolvedAssetRoot = $sceneDir
}
else {
    if (-not (Test-Path -LiteralPath $AssetRoot -PathType Container)) {
        throw "AssetRoot was not found: $AssetRoot"
    }
    $script:ResolvedAssetRoot = (Resolve-Path -LiteralPath $AssetRoot).Path
}

$resolvedOutputPptx = Resolve-OutputPath $OutputPptx

$resolvedPreviewPath = $null
if ($ExportPreview) {
    if ([string]::IsNullOrWhiteSpace($PreviewPath)) {
        $resolvedPreviewPath = [System.IO.Path]::ChangeExtension($resolvedOutputPptx, ".png")
    }
    else {
        $resolvedPreviewPath = Resolve-OutputPath $PreviewPath
    }

    if ($PreviewWidth -le 0 -or $PreviewHeight -le 0) {
        throw "PreviewWidth and PreviewHeight must be positive integers."
    }
}

$utf8Strict = [System.Text.UTF8Encoding]::new($false, $true)
try {
    $sceneJsonText = [System.IO.File]::ReadAllText($resolvedSceneJson, $utf8Strict)
}
catch {
    throw "SceneJson must be valid UTF-8: $resolvedSceneJson. $($_.Exception.Message)"
}
$convertFromJsonCommand = Get-Command ConvertFrom-Json
if ($convertFromJsonCommand.Parameters.ContainsKey("Depth")) {
    $scene = $sceneJsonText | ConvertFrom-Json -Depth 100
}
else {
    $scene = $sceneJsonText | ConvertFrom-Json
}
$defaultTheme = [pscustomobject]@{
    colors = [pscustomobject]@{
        primary = "#006BD6"
        accent = "#D000FF"
        success = "#49A046"
        warning = "#FF7F00"
        text = "#000000"
        muted = "#666666"
    }
    fonts = [pscustomobject]@{
        body = "Arial"
        heading = "Arial"
    }
    strokeWidths = [pscustomobject]@{
        hairline = 1
        normal = 2
        heavy = 3
    }
}
$userTheme = Get-JsonProp $scene "theme" $null
if ($null -ne $userTheme) {
    foreach ($sectionName in @("colors", "fonts", "strokeWidths")) {
        $section = Get-JsonProp $userTheme $sectionName $null
        if ($null -ne $section) {
            foreach ($prop in $section.PSObject.Properties) {
                if ($null -eq $defaultTheme.$sectionName.PSObject.Properties[$prop.Name]) {
                    $defaultTheme.$sectionName | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value
                }
                else {
                    $defaultTheme.$sectionName.PSObject.Properties[$prop.Name].Value = $prop.Value
                }
            }
        }
    }
}
$script:Theme = $defaultTheme
$canvas = Get-JsonProp $scene "canvas" $null
$canvasWidth = [double](Get-JsonProp $canvas "width" 1280)
$canvasHeight = [double](Get-JsonProp $canvas "height" 720)

if ($canvasWidth -le 0 -or $canvasHeight -le 0) {
    throw "canvas.width and canvas.height must be positive numbers."
}

$slides = ConvertTo-Array (Get-RequiredProp $scene "slides" "scene")
if ($slides.Count -eq 0) {
    throw "scene.slides must contain at least one slide."
}

$page = Get-JsonProp $scene "page" $null
$slideWidth = [double](Get-JsonProp $page "width" 960)
$slideHeight = [double](Get-JsonProp $page "height" 540)
if ($slideWidth -le 0 -or $slideHeight -le 0) {
    throw "page.width and page.height must be positive numbers when provided."
}
$script:ScaleX = $slideWidth / $canvasWidth
$script:ScaleY = $slideHeight / $canvasHeight
$script:StrokeScale = ($script:ScaleX + $script:ScaleY) / 2
$script:TextScale = $script:ScaleY
$script:OffsetX = 0
$script:OffsetY = 0

$ppt = $null
$presentation = $null
$shouldClosePowerPoint = -not $KeepPowerPointOpen

try {
    try {
        $ppt = New-Object -ComObject PowerPoint.Application
    }
    catch {
        throw "PowerPoint COM automation is not available. Install Microsoft PowerPoint desktop and try again. $($_.Exception.Message)"
    }

    if ($LivePreview -or $KeepPowerPointOpen) {
        $ppt.Visible = -1
    }
    Set-PowerPointQuietMode $ppt

    $presentation = $ppt.Presentations.Add()
    $presentation.PageSetup.SlideWidth = $slideWidth
    $presentation.PageSetup.SlideHeight = $slideHeight

    for ($slideIndex = 0; $slideIndex -lt $slides.Count; $slideIndex++) {
        $slideSpec = $slides[$slideIndex]
        $slide = $presentation.Slides.Add($slideIndex + 1, 12)

        $background = Get-JsonProp $slideSpec "background" $null
        if ($null -ne $background) {
            $rgb = ConvertTo-RgbInt $background "#FFFFFF"
            if ($null -ne $rgb) {
                $slide.Background.Fill.ForeColor.RGB = $rgb
            }
        }

        $elements = Sort-SceneElements (Get-JsonProp $slideSpec "elements" @())
        for ($elementIndex = 0; $elementIndex -lt $elements.Count; $elementIndex++) {
            $element = $elements[$elementIndex]
            $elementId = Get-JsonProp $element "id" "element-$($elementIndex + 1)"
            try {
                $shape = Add-SceneElement $slide $element ($slideIndex + 1) ($elementIndex + 1)
                if ($null -ne $shape) {
                    try {
                        $shape.Name = [string]$elementId
                    }
                    catch {
                        Write-Warning "Could not set PowerPoint shape name for '$elementId': $($_.Exception.Message)"
                    }
                }
            }
            catch {
                throw "Failed to render slide $($slideIndex + 1) element '$elementId': $($_.Exception.Message)"
            }
            Invoke-LiveStep $shape
        }
    }

    $presentation.SaveAs($resolvedOutputPptx)
    Complete-PresentationSaveState $presentation
    Write-Host "Saved PPTX: $resolvedOutputPptx"

    if ($ExportPreview) {
        $presentation.Slides.Item(1).Export($resolvedPreviewPath, "PNG", $PreviewWidth, $PreviewHeight) | Out-Null
        Complete-PresentationSaveState $presentation
        Write-Host "Saved preview: $resolvedPreviewPath"
    }

    if ($shouldClosePowerPoint) {
        Complete-PresentationSaveState $presentation
        $presentation.Close()
        $ppt.Quit()
    }
    elseif ($KeepPowerPointOpen) {
        Write-Host "KeepPowerPointOpen is enabled; PowerPoint was left open for inspection."
    }
}
catch {
    Write-Error $_.Exception.Message

    if ($null -ne $presentation -and $shouldClosePowerPoint) {
        try { Complete-PresentationSaveState $presentation } catch {}
        try { $presentation.Close() } catch {}
    }
    if ($null -ne $ppt -and $shouldClosePowerPoint) {
        try { $ppt.Quit() } catch {}
    }

    exit 1
}
finally {
    if ($shouldClosePowerPoint) {
        if ($null -ne $presentation) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
        }
        if ($null -ne $ppt) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
