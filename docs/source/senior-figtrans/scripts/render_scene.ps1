# Copyright (c) 2026 NCEPU-Senior. Author: XiaoN.
<#
.SYNOPSIS
Render a deliberately small, explicit scene into native PowerPoint and/or Visio.
.DESCRIPTION
Windows PowerShell 5.1 plus installed Microsoft PowerPoint/Visio are required.
This helper is a primitive scaffold, not an image recognizer. It supports rect,
roundRect, ellipse, text, line, polyline, and attached straight/elbow connector.
Coordinates, font sizes and stroke widths are points; origin is the top left.
head_size selects a small/medium/large arrowhead in both applications. Optional
head_width and head_length independently override PowerPoint arrowhead dimensions
using those same size names. Visio uses head_size only; it has no independent
width/length controls in this scaffold and emits a warning for those overrides.
No pictures, equation substitution, charts, groups, or silent fallback are used.
Unknown object types and existing output files are rejected before COM startup.
The source is saved, closed, reopened, checked and exported to PNG. Only documents
created by this invocation are closed. No application-wide Quit is performed.
Visio reuses a running instance when available. A cold start is initialized by a
separate Windows PowerShell 5.1 worker, then reattached after that worker exits.
Only uniquely identified task-owned guards and the new task document are closed.
An additional guard is created only when needed to avoid a zero-document UI
transition between saving and reopening. Existing documents are never closed.
Visio also adds one native white full-page background (sft_canvas_background)
so drawing-bounds export retains the scene canvas and its outer margins.
Each target has independent error handling. A progress receipt is written before
COM stages; partial deliveries remain recorded and any failure returns exit 1.
Only rejected/busy COM calls are retried, at most three times (250/500/1000 ms).
This cannot interrupt a COM call that never returns; inspect the receipt's stage
and live process handle before deciding whether an external timeout is terminal.
For complex diagrams extend a task-local copy or use native application tools.
.EXAMPLE
powershell -NoProfile -File render_scene.ps1 -ScenePath native-smoke.json -Target both -OutputDir D:\figures\smoke
.EXAMPLE
powershell -NoProfile -File render_scene.ps1 -ScenePath native-smoke.json -Target both -OutputDir D:\figures\smoke -ValidateOnly
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$ScenePath,
    [Parameter(Mandatory=$true)][ValidateSet('pptx','visio','both')][string]$Target,
    [Parameter(Mandatory=$true)][string]$OutputDir,
    [switch]$ValidateOnly,
    [switch]$ExportEmf
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Invariant = [Globalization.CultureInfo]::InvariantCulture
$CanvasBackgroundId = 'sft_canvas_background'

function Value($Object, [string]$Name, $Default) {
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}
function Num($Object, [string]$Name, [double]$Default) {
    $value = [double](Value $Object $Name $Default)
    if ([double]::IsNaN($value) -or [double]::IsInfinity($value)) { throw "Non-finite number: $Name" }
    return $value
}
function NumberText([double]$Value) { return $Value.ToString('0.########', $Invariant) }
function Flush-Receipt {
    if($null -ne $script:Receipt) {
        $script:Receipt.updated_utc=[DateTime]::UtcNow.ToString('o')
        $script:Receipt | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $receiptPath -Encoding UTF8
    }
}
function Stage([string]$Name) {
    $script:CurrentEntry.current_stage=$Name
    $null=$script:CurrentEntry.stage_history.Add([ordered]@{stage=$Name;utc=[DateTime]::UtcNow.ToString('o')})
    Flush-Receipt
}
function Com-ErrorCode($ErrorRecord) {
    $exception=$ErrorRecord.Exception
    $codes=New-Object System.Collections.ArrayList
    while($null -ne $exception) {
        $null=$codes.Add(('{0:X8}' -f [int]$exception.HResult))
        $exception=$exception.InnerException
    }
    return @($codes)
}
function Invoke-ComCall([string]$StageName, [scriptblock]$Operation) {
    Stage $StageName
    for($attempt=0;$attempt -le 3;$attempt++) {
        try {
            $result=& $Operation
            return ,$result
        } catch {
            $codes=Com-ErrorCode $_
            $retryable=('80010001' -in $codes -or '8001010A' -in $codes)
            if(-not $retryable -or $attempt -eq 3) { throw }
            $null=$script:CurrentEntry.com_retries.Add([ordered]@{stage=$StageName;attempt=($attempt+1);hresults=$codes;message=$_.Exception.Message})
            Flush-Receipt
            Start-Sleep -Milliseconds ([int](250*[Math]::Pow(2,$attempt)))
        }
    }
}
function Record-AppIdentity($Application, [string]$Format) {
    Stage 'read_application_identity'
    try {
        if(-not ('SeniorFigTrans.NativeWindow' -as [type])) {
            Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; namespace SeniorFigTrans { public static class NativeWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid); } }'
        }
        if($Format -eq 'vsdx') {
            $script:CurrentEntry.application.visio_instance_id=[long]$Application.ProcessID
            $hwnd=[long]$Application.WindowHandle32
        } else { $hwnd=[long]$Application.HWND }
        $script:CurrentEntry.application.hwnd=$hwnd
        [uint32]$applicationPid=0
        $threadId=[SeniorFigTrans.NativeWindow]::GetWindowThreadProcessId([IntPtr]$hwnd,[ref]$applicationPid)
        if($threadId -ne 0) { $script:CurrentEntry.application.windows_pid=$applicationPid }
    } catch {
        $script:CurrentEntry.application.identity_warning=$_.Exception.Message
    }
    Flush-Receipt
}
function Failed-Target($ErrorRecord) {
    $script:CurrentEntry.status='failed'
    $script:CurrentEntry.error=[ordered]@{
        stage=$script:CurrentEntry.current_stage
        message=$ErrorRecord.Exception.Message
        hresults=(Com-ErrorCode $ErrorRecord)
        script_stack=$ErrorRecord.ScriptStackTrace
    }
    Flush-Receipt
    Write-Warning ($script:CurrentEntry.format+': '+$ErrorRecord.Exception.Message)
}
function Cleanup-Call([string]$Name, [scriptblock]$Operation) {
    try { $null=Invoke-ComCall $Name $Operation }
    catch {
        $null=$script:CurrentEntry.cleanup_errors.Add([ordered]@{stage=$Name;message=$_.Exception.Message;hresults=(Com-ErrorCode $_)})
        # Preserve the original main error; cleanup failures are separate evidence.
        if($script:CurrentEntry.status -ne 'failed') { $script:CurrentEntry.status='cleanup_failed' }
        Flush-Receipt
        Write-Warning ('Cleanup '+$Name+': '+$_.Exception.Message)
    }
}
function RgbValue([string]$Hex) {
    if ($Hex -notmatch '^#[0-9A-Fa-f]{6}$') { throw "Expected #RRGGBB color, got '$Hex'." }
    return [Convert]::ToInt32($Hex.Substring(1,2),16) + 256*[Convert]::ToInt32($Hex.Substring(3,2),16) + 65536*[Convert]::ToInt32($Hex.Substring(5,2),16)
}
function RgbFormula([string]$Hex) {
    $rgb = RgbValue $Hex
    return ('RGB({0},{1},{2})' -f ($rgb -band 255), (($rgb -shr 8) -band 255), (($rgb -shr 16) -band 255))
}
function Check-Color([string]$Color, [bool]$AllowNone=$true) {
    if ($AllowNone -and $Color -eq 'none') { return }
    $null = RgbValue $Color
}
function Set-VCell($Shape, [string]$Cell, [string]$Formula) {
    try { $Shape.CellsU($Cell).FormulaU = $Formula }
    catch { throw "Cannot set $Cell on '$($Shape.NameU)': $($_.Exception.Message)" }
}
function Side-Point($Object, [string]$Side) {
    $x=[double]$Object.x; $y=[double]$Object.y; $w=[double]$Object.w; $h=[double]$Object.h
    switch($Side) {
        top { return @(($x+$w/2),$y) }
        bottom { return @(($x+$w/2),($y+$h)) }
        left { return @($x,($y+$h/2)) }
        right { return @(($x+$w),($y+$h/2)) }
        default { throw "Unknown connector side '$Side'." }
    }
}
function V-GluePoint([string]$Side) {
    switch($Side) {
        top { return @(0.5,1.0) }
        bottom { return @(0.5,0.0) }
        left { return @(0.0,0.5) }
        right { return @(1.0,0.5) }
        default { throw "Unknown connector side '$Side'." }
    }
}
function P-Site($Object, [string]$Side) {
    if ($Object.type -eq 'ellipse') {
        return @{top=1;left=3;bottom=5;right=7}[$Side]
    }
    return @{top=1;left=2;bottom=3;right=4}[$Side]
}
function Set-PStyle($Shape, $Object, [bool]$IsPath) {
    $stroke=[string](Value $Object 'stroke' '#293344')
    if($stroke -eq 'none') { $Shape.Line.Visible=0 } else {
        $Shape.Line.Visible=-1
        $Shape.Line.ForeColor.RGB=RgbValue $stroke
        $Shape.Line.Weight=[single](Num $Object 'stroke_width' 1.0)
        $Shape.Line.DashStyle=1
    }
    $fill=[string](Value $Object 'fill' 'none')
    if($IsPath -or $fill -eq 'none') { $Shape.Fill.Visible=0 } else {
        $Shape.Fill.Visible=-1; $Shape.Fill.Solid()
        $Shape.Fill.ForeColor.RGB=RgbValue $fill
    }
    if($IsPath) {
        $Shape.Line.BeginArrowheadStyle=1
        $Shape.Line.EndArrowheadStyle=1
        if((Value $Object 'head' 'none') -eq 'triangle') { $Shape.Line.EndArrowheadStyle=2 }
        $headSize=[string](Value $Object 'head_size' 'small')
        $Shape.Line.EndArrowheadLength=@{small=1;medium=2;large=3}[[string](Value $Object 'head_length' $headSize)]
        $Shape.Line.EndArrowheadWidth=@{small=1;medium=2;large=3}[[string](Value $Object 'head_width' $headSize)]
    }
    $text=[string](Value $Object 'text' '')
    if(-not $IsPath -and ($text.Length -gt 0 -or $Object.type -eq 'text')) {
        $tf=$Shape.TextFrame2
        $tf.AutoSize=0; $tf.WordWrap=-1; $tf.VerticalAnchor=3
        $margin=Num $Object 'text_margin' 2
        $tf.MarginLeft=$margin; $tf.MarginRight=$margin; $tf.MarginTop=$margin; $tf.MarginBottom=$margin
        $tf.TextRange.Text=$text
        $tf.TextRange.Font.Name=[string](Value $Object 'font' 'Times New Roman')
        $tf.TextRange.Font.Size=[single](Num $Object 'font_size' 16)
        $tf.TextRange.Font.Bold=0
        if([bool](Value $Object 'bold' $false)) { $tf.TextRange.Font.Bold=-1 }
        $tf.TextRange.Font.Fill.ForeColor.RGB=RgbValue ([string](Value $Object 'text_color' '#172033'))
        $tf.TextRange.ParagraphFormat.Alignment=@{left=1;center=2;right=3}[[string](Value $Object 'align' 'center')]
        $tf.TextRange.ParagraphFormat.SpaceBefore=0; $tf.TextRange.ParagraphFormat.SpaceAfter=0
    }
}
function Set-VStyle($Shape, $Object, [bool]$IsPath, $Document) {
    $stroke=[string](Value $Object 'stroke' '#293344')
    if($stroke -eq 'none') { Set-VCell $Shape 'LinePattern' '0' } else {
        Set-VCell $Shape 'LinePattern' '1'
        Set-VCell $Shape 'LineColor' (RgbFormula $stroke)
        Set-VCell $Shape 'LineWeight' ((NumberText (Num $Object 'stroke_width' 1.0))+' pt')
    }
    $fill=[string](Value $Object 'fill' 'none')
    if($IsPath -or $fill -eq 'none') { Set-VCell $Shape 'FillPattern' '0' } else {
        Set-VCell $Shape 'FillPattern' '1'; Set-VCell $Shape 'FillForegnd' (RgbFormula $fill)
    }
    Set-VCell $Shape 'ShdwPattern' '0'
    if($IsPath) {
        Set-VCell $Shape 'BeginArrow' '0'; Set-VCell $Shape 'EndArrow' '0'
        if((Value $Object 'head' 'none') -eq 'triangle') { Set-VCell $Shape 'EndArrow' '4' }
        Set-VCell $Shape 'EndArrowSize' ([string]@{small=1;medium=2;large=3}[[string](Value $Object 'head_size' 'small')])
    }
    $text=[string](Value $Object 'text' '')
    if(-not $IsPath -and ($text.Length -gt 0 -or $Object.type -eq 'text')) {
        $Shape.Text=$text
        $fontName=[string](Value $Object 'font' 'Times New Roman')
        $fontId=$Document.Fonts.Item($fontName).ID
        Set-VCell $Shape 'Char.Font' ([string]$fontId)
        Set-VCell $Shape 'Char.Size' ((NumberText (Num $Object 'font_size' 16))+' pt')
        Set-VCell $Shape 'Char.Color' (RgbFormula ([string](Value $Object 'text_color' '#172033')))
        $bold='0'; if([bool](Value $Object 'bold' $false)) { $bold='1' }
        Set-VCell $Shape 'Char.Style' $bold
        Set-VCell $Shape 'Para.HorzAlign' ([string]@{left=0;center=1;right=2}[[string](Value $Object 'align' 'center')])
        Set-VCell $Shape 'VerticalAlign' '1'
        foreach($cell in @('LeftMargin','RightMargin','TopMargin','BottomMargin')) {
            Set-VCell $Shape $cell ((NumberText (Num $Object 'text_margin' 2))+' pt')
        }
    }
}

# Validate all input before starting an application.
$ScenePath=[IO.Path]::GetFullPath($ScenePath)
$OutputDir=[IO.Path]::GetFullPath($OutputDir)
$scene=Get-Content -LiteralPath $ScenePath -Raw -Encoding UTF8 | ConvertFrom-Json
$canvas=$scene.canvas
$width=Num $canvas 'width' 0; $height=Num $canvas 'height' 0
if($width -lt 72 -or $height -lt 72 -or $width -gt 4032 -or $height -gt 4032) { throw 'Canvas width/height must be 72..4032 points.' }
$name=[string](Value $scene 'name' 'figure')
if($name -notmatch '^[A-Za-z0-9][A-Za-z0-9_.-]*$' -or $name.EndsWith('.')) { throw 'Scene name must be a simple ASCII filename stem.' }
if(@($scene.objects).Count -eq 0) { throw 'Scene contains no objects.' }
$spec=@{}
$nodes=@(); $paths=@()
foreach($obj in $scene.objects) {
    $id=[string]$obj.id
    if($id -notmatch '^[A-Za-z][A-Za-z0-9_.-]*$' -or $spec.ContainsKey($id)) { throw "Invalid or duplicate id '$id'." }
    if($id -ieq $CanvasBackgroundId) { throw "Scene id '$id' is reserved for the native Visio canvas background." }
    $kind=[string]$obj.type
    if($kind -notin @('rect','roundRect','ellipse','text','line','polyline','connector')) { throw "Unsupported type '$kind' on '$id'." }
    $spec[$id]=$obj
    Check-Color ([string](Value $obj 'fill' 'none'))
    Check-Color ([string](Value $obj 'stroke' '#293344'))
    Check-Color ([string](Value $obj 'text_color' '#172033')) $false
    if((Num $obj 'stroke_width' 1) -le 0) { throw "stroke_width must be positive: $id" }
    if((Num $obj 'font_size' 16) -le 0) { throw "font_size must be positive: $id" }
    if((Value $obj 'head' 'none') -notin @('none','triangle')) { throw "Unsupported head: $id" }
    if((Value $obj 'head_size' 'small') -notin @('small','medium','large')) { throw "Unsupported head_size: $id" }
    foreach($key in @('head_width','head_length')) {
        if((Value $obj $key (Value $obj 'head_size' 'small')) -notin @('small','medium','large')) { throw "Unsupported ${key}: $id" }
        if($Target -in @('visio','both') -and $null -ne $obj.PSObject.Properties[$key]) {
            Write-Warning "${id}: $key affects only PowerPoint; Visio uses the combined head_size setting."
        }
    }
    if((Value $obj 'align' 'center') -notin @('left','center','right')) { throw "Unsupported align: $id" }
    if($kind -in @('rect','roundRect','ellipse','text')) {
        foreach($key in @('x','y','w','h')) { if($null -eq $obj.PSObject.Properties[$key]) { throw "Missing '$key' on '$id'." } }
        $x=Num $obj 'x' 0; $y=Num $obj 'y' 0; $w=Num $obj 'w' 0; $h=Num $obj 'h' 0
        if($x -lt 0 -or $y -lt 0 -or $w -le 0 -or $h -le 0 -or $x+$w -gt $width -or $y+$h -gt $height) { throw "Object outside canvas or nonpositive size: $id" }
        $nodes+=,$obj
    } else { $paths+=,$obj }
}
foreach($obj in $paths) {
    if($obj.type -eq 'connector') {
        foreach($key in @('from','to')) {
            $ref=[string]$obj.$key
            if(-not $spec.ContainsKey($ref) -or $spec[$ref].type -notin @('rect','roundRect','ellipse','text')) { throw "Invalid connector endpoint '$ref'." }
        }
        foreach($key in @('from_side','to_side')) { if((Value $obj $key '') -notin @('top','bottom','left','right')) { throw "Invalid connector $key on $($obj.id)." } }
        if((Value $obj 'route' 'straight') -notin @('straight','elbow')) { throw "Invalid connector route on $($obj.id)." }
    } else {
        $points=@($obj.points)
        if($points.Count -lt 2 -or ($obj.type -eq 'line' -and $points.Count -ne 2)) { throw "Invalid point count on $($obj.id)." }
        for($i=0;$i -lt $points.Count;$i++) {
            if(@($points[$i]).Count -ne 2) { throw "Each point must be [x,y]: $($obj.id)" }
            $x=[double]$points[$i][0]; $y=[double]$points[$i][1]
            if([double]::IsNaN($x) -or [double]::IsNaN($y) -or $x -lt 0 -or $y -lt 0 -or $x -gt $width -or $y -gt $height) { throw "Path point outside canvas: $($obj.id)" }
            if($i -gt 0 -and $x -eq [double]$points[$i-1][0] -and $y -eq [double]$points[$i-1][1]) { throw "Duplicate path point creates zero-length segment: $($obj.id)" }
        }
    }
}
$formats=@(); if($Target -in @('pptx','both')) { $formats+='pptx' }; if($Target -in @('visio','both')) { $formats+='vsdx' }
$files=@()
foreach($format in $formats) {
    $files+=(Join-Path $OutputDir ($name+'.'+$format))
    $files+=(Join-Path $OutputDir ($name+'_'+$format+'.png'))
    if($ExportEmf) { $files+=(Join-Path $OutputDir ($name+'_'+$format+'.emf')) }
}
$receiptPath=Join-Path $OutputDir ($name+'_render_receipt.json'); $files+=$receiptPath
foreach($file in $files) { if(Test-Path -LiteralPath $file) { throw "Refusing to overwrite '$file'. Choose a new output directory or scene name." } }
if($ValidateOnly) {
    [pscustomobject]@{status='schema_valid';target=$Target;objects=$spec.Count;planned_files=$files} | ConvertTo-Json -Depth 6
    exit 0
}
$null=New-Item -ItemType Directory -Path $OutputDir -Force
$results=New-Object System.Collections.ArrayList
$script:Receipt=[ordered]@{status='running';author='XiaoN';copyright='NCEPU-Senior';scene=$ScenePath;target=$Target;started_utc=[DateTime]::UtcNow.ToString('o');updated_utc=$null;results=$results;arrowhead_sizing='PowerPoint: head_width/head_length override head_size independently. Visio: head_size only.';visual_review='Required: inspect each PNG at whole-figure scale and connector details.'}
Flush-Receipt

if($Target -in @('pptx','both')) {
    $app=$null; $doc=$null; $reopened=$null
    $file=Join-Path $OutputDir ($name+'.pptx')
    $png=Join-Path $OutputDir ($name+'_pptx.png')
    $script:CurrentEntry=[ordered]@{format='pptx';status='running';path=$file;preview=$png;current_stage='pending';stage_history=(New-Object System.Collections.ArrayList);com_retries=(New-Object System.Collections.ArrayList);cleanup_errors=(New-Object System.Collections.ArrayList);application=[ordered]@{windows_pid=$null;hwnd=$null;identity_warning=$null};error=$null;saved_reopened=$false;native_objects=$null;scene_objects=$spec.Count;canvas_background_objects=0;canvas_background_id=$null;attached_connectors=$null;renderer='Microsoft PowerPoint';visual_review='pending'}
    $null=$results.Add($script:CurrentEntry); Flush-Receipt
    try {
        $app=Invoke-ComCall 'create_application' { New-Object -ComObject PowerPoint.Application }
        Record-AppIdentity $app 'pptx'
        $doc=Invoke-ComCall 'create_document' { $app.Presentations.Add(0) }
        Stage 'configure_canvas'
        $doc.PageSetup.SlideWidth=[single]$width; $doc.PageSetup.SlideHeight=[single]$height
        $slide=$doc.Slides.Add(1,12); $slide.FollowMasterBackground=0
        $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB=RgbValue '#FFFFFF'
        $created=@{}
        foreach($obj in $nodes) {
            Stage ('draw_node:'+([string]$obj.id))
            $kind=[string]$obj.type
            if($kind -eq 'text') { $shape=$slide.Shapes.AddTextbox(1,[single]$obj.x,[single]$obj.y,[single]$obj.w,[single]$obj.h) }
            else {
                $shapeType=@{rect=1;roundRect=5;ellipse=9}[$kind]
                $shape=$slide.Shapes.AddShape($shapeType,[single]$obj.x,[single]$obj.y,[single]$obj.w,[single]$obj.h)
                if($kind -eq 'roundRect') { $shape.Adjustments.Item(1)=[single]0.12 }
            }
            $shape.Name=[string]$obj.id; $created[[string]$obj.id]=$shape
            Set-PStyle $shape $obj $false
        }
        foreach($obj in $paths) {
            Stage ('draw_path:'+([string]$obj.id))
            if($obj.type -eq 'connector') {
                $start=Side-Point $spec[[string]$obj.from] ([string]$obj.from_side)
                $end=Side-Point $spec[[string]$obj.to] ([string]$obj.to_side)
                $route=1; if((Value $obj 'route' 'straight') -eq 'elbow') { $route=2 }
                $shape=$slide.Shapes.AddConnector($route,[single]$start[0],[single]$start[1],[single]$end[0],[single]$end[1])
                $shape.ConnectorFormat.BeginConnect($created[[string]$obj.from],(P-Site $spec[[string]$obj.from] ([string]$obj.from_side)))
                $shape.ConnectorFormat.EndConnect($created[[string]$obj.to],(P-Site $spec[[string]$obj.to] ([string]$obj.to_side)))
            } else {
                $pts=@($obj.points)
                if($obj.type -eq 'line') { $shape=$slide.Shapes.AddLine([single]$pts[0][0],[single]$pts[0][1],[single]$pts[1][0],[single]$pts[1][1]) }
                else {
                    $builder=$slide.Shapes.BuildFreeform(1,[single]$pts[0][0],[single]$pts[0][1])
                    for($i=1;$i -lt $pts.Count;$i++) { $builder.AddNodes(0,1,[single]$pts[$i][0],[single]$pts[$i][1]) }
                    $shape=$builder.ConvertToShape()
                }
            }
            $shape.Name=[string]$obj.id; $created[[string]$obj.id]=$shape
            Set-PStyle $shape $obj $true
        }
        $null=Invoke-ComCall 'save_native' { $doc.SaveAs($file,24) }
        $null=Invoke-ComCall 'close_saved_document' { $doc.Close() }; $doc=$null
        $reopened=Invoke-ComCall 'reopen_saved_document' { $app.Presentations.Open($file,-1,0,0) }
        Stage 'verify_reopened_objects'
        $check=$reopened.Slides.Item(1)
        if($check.Shapes.Count -ne $spec.Count) { throw 'PowerPoint saved shape count does not match scene.' }
        $attached=0
        foreach($obj in $scene.objects) {
            $shape=$check.Shapes.Item([string]$obj.id)
            if($obj.type -eq 'connector') {
                if($shape.ConnectorFormat.BeginConnected -ne -1 -or $shape.ConnectorFormat.EndConnected -ne -1) { throw "Detached connector after reopening: $($obj.id)" }
                $attached++
            }
        }
        $script:CurrentEntry.native_objects=$check.Shapes.Count
        $script:CurrentEntry.attached_connectors=$attached
        $script:CurrentEntry.saved_reopened=$true
        $null=Invoke-ComCall 'export_png' { $check.Export($png,'PNG',[int]($width*2),[int]($height*2)) }
        if($ExportEmf) { $null=Invoke-ComCall 'export_emf' { $check.Export((Join-Path $OutputDir ($name+'_pptx.emf')),'EMF') } }
        $script:CurrentEntry.status='native_saved_and_reopened'
        Stage 'target_verified'
    } catch {
        Failed-Target $_
    } finally {
        if($null -ne $reopened) { Cleanup-Call 'close_reopened_document' { $reopened.Close() } }
        if($null -ne $doc) { Cleanup-Call 'discard_owned_unsaved_document' { $doc.Saved=-1; $doc.Close() } }
        if($null -ne $app) { Cleanup-Call 'release_application_reference' { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) } }
        Stage 'target_finished'
    }
}
if($Target -in @('visio','both')) {
    $app=$null; $doc=$null; $reopened=$null; $guardDocument=$null
    $file=Join-Path $OutputDir ($name+'.vsdx')
    $png=Join-Path $OutputDir ($name+'_vsdx.png')
    $script:CurrentEntry=[ordered]@{format='vsdx';status='running';path=$file;preview=$png;current_stage='pending';stage_history=(New-Object System.Collections.ArrayList);com_retries=(New-Object System.Collections.ArrayList);cleanup_errors=(New-Object System.Collections.ArrayList);application=[ordered]@{windows_pid=$null;hwnd=$null;visio_instance_id=$null;identity_warning=$null;reused_existing_instance=$null;initial_document_count=$null;startup_worker_pid=$null;startup_worker_receipt=$null;startup_worker_state='not_needed'};guard=[ordered]@{created=$false;closed=$false;document_id=$null;title=$null;origin=$null};error=$null;saved_reopened=$false;native_objects=$null;scene_objects=$spec.Count;canvas_background_objects=0;canvas_background_id=$CanvasBackgroundId;attached_connectors=$null;renderer='Microsoft Visio';visual_review='pending'}
    $null=$results.Add($script:CurrentEntry); Flush-Receipt
    try {
        try {
            $app=Invoke-ComCall 'get_active_application' { [Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application') }
            $script:CurrentEntry.application.reused_existing_instance=$true
        } catch {
            # Only MK_E_UNAVAILABLE means no running instance. Other failures are
            # surfaced rather than starting a second instance behind a busy one.
            if('800401E3' -notin (Com-ErrorCode $_)) { throw }
            $script:CurrentEntry.application.reused_existing_instance=$false
            $workerScript=Join-Path $PSScriptRoot 'start_visio_host.ps1'
            $workerReceipt=Join-Path $OutputDir ($name+'_visio_host.json')
            $workerPowerShell=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
            if(-not (Test-Path -LiteralPath $workerScript -PathType Leaf)) { throw "Missing startup worker: $workerScript" }
            if(Test-Path -LiteralPath $workerReceipt) { throw "An existing startup receipt needs inspection before another launch: $workerReceipt" }
            $guardToken='sft-'+[Guid]::NewGuid().ToString('N')
            $script:CurrentEntry.application.startup_worker_receipt=$workerReceipt
            $script:CurrentEntry.application.startup_worker_state='starting'
            Stage 'start_isolated_visio_worker'
            # Arguments are passed directly to powershell.exe, not to another shell.
            # Windows paths cannot contain quotes; all path arguments are quoted.
            $workerArguments=@('-NoProfile','-NonInteractive','-STA','-File',('"'+$workerScript+'"'),'-ReceiptPath',('"'+$workerReceipt+'"'),'-GuardToken',$guardToken)
            $hostWorker=Start-Process -FilePath $workerPowerShell -ArgumentList $workerArguments -WindowStyle Hidden -PassThru
            $script:CurrentEntry.application.startup_worker_pid=$hostWorker.Id
            $script:CurrentEntry.application.startup_worker_state='running'
            Stage 'wait_for_visio_worker'
            if(-not $hostWorker.WaitForExit(20000)) {
                $script:CurrentEntry.application.startup_worker_state='timeout_may_still_be_running'
                Flush-Receipt
                throw "Visio startup worker PID $($hostWorker.Id) did not exit within 20 seconds. It may still be running; inspect its process and receipt before continuing. No process was killed or restarted."
            }
            $hostWorker.Refresh()
            if($hostWorker.ExitCode -ne 0) {
                $script:CurrentEntry.application.startup_worker_state='failed'
                $workerFailure='No worker receipt was produced.'
                if(Test-Path -LiteralPath $workerReceipt) { $workerFailure=Get-Content -LiteralPath $workerReceipt -Raw -Encoding UTF8 }
                throw "Visio startup worker exited $($hostWorker.ExitCode): $workerFailure"
            }
            if(-not (Test-Path -LiteralPath $workerReceipt)) { throw 'Visio startup worker returned success without its receipt.' }
            $hostState=Get-Content -LiteralPath $workerReceipt -Raw -Encoding UTF8 | ConvertFrom-Json
            if($hostState.status -ne 'ready' -or $hostState.guard_token -cne $guardToken) { throw 'Visio startup receipt is not ready or has the wrong guard token.' }
            $script:CurrentEntry.application.startup_worker_state='ready'
            $app=Invoke-ComCall 'reattach_initialized_visio' { [Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application') }
            $attachedHwnd=Invoke-ComCall 'verify_initialized_visio_hwnd' { [long]$app.WindowHandle32 }
            if($attachedHwnd -ne [long]$hostState.hwnd) { throw 'GetActiveObject selected a different Visio HWND. No document in that instance will be changed.' }
            $candidateGuard=Invoke-ComCall 'locate_owned_startup_guard' { $app.Documents.ItemFromID([int]$hostState.guard_document_id) }
            if([string]$candidateGuard.Title -cne [string]$hostState.guard_title -or [string]$hostState.guard_title -cne ('Senior-FigTrans startup guard '+$guardToken)) { throw 'Startup guard ID/title validation failed. The document will not be closed.' }
            $guardDocument=$candidateGuard
            $script:CurrentEntry.guard.created=$true
            $script:CurrentEntry.guard.document_id=[int]$hostState.guard_document_id
            $script:CurrentEntry.guard.title=[string]$hostState.guard_title
            $script:CurrentEntry.guard.origin='isolated_startup_worker'
            Stage 'owned_startup_guard_verified'
        }
        Record-AppIdentity $app 'vsdx'
        $script:CurrentEntry.application.initial_document_count=Invoke-ComCall 'inspect_existing_document_count' { $app.Documents.Count }
        $doc=Invoke-ComCall 'create_document' { $app.Documents.Add('') }
        Stage 'configure_canvas'
        $page=$doc.Pages.Item(1); $page.Name='Figure'
        Set-VCell $page.PageSheet 'PageWidth' ((NumberText $width)+' pt')
        Set-VCell $page.PageSheet 'PageHeight' ((NumberText $height)+' pt')
        Stage 'create_native_canvas_background'
        # First shape is behind every scene object. An opaque native fill, not an
        # invisible box or bitmap, defines the full export extent in Visio.
        $canvasBackground=$page.DrawRectangle(0,0,($width/72),($height/72))
        $canvasBackground.NameU=$CanvasBackgroundId
        Set-VCell $canvasBackground 'FillPattern' '1'
        Set-VCell $canvasBackground 'FillForegnd' 'RGB(255,255,255)'
        Set-VCell $canvasBackground 'FillForegndTrans' '0'
        Set-VCell $canvasBackground 'LinePattern' '0'
        Set-VCell $canvasBackground 'ShdwPattern' '0'
        $script:CurrentEntry.canvas_background_objects=1
        $created=@{}
        foreach($obj in $nodes) {
            Stage ('draw_node:'+([string]$obj.id))
            $x=[double]$obj.x/72; $y=($height-[double]$obj.y-[double]$obj.h)/72
            $right=([double]$obj.x+[double]$obj.w)/72; $top=($height-[double]$obj.y)/72
            if($obj.type -eq 'ellipse') { $shape=$page.DrawOval($x,$y,$right,$top) }
            else { $shape=$page.DrawRectangle($x,$y,$right,$top) }
            $shape.NameU=[string]$obj.id; $created[[string]$obj.id]=$shape
            Set-VStyle $shape $obj $false $doc
            if($obj.type -eq 'roundRect') { Set-VCell $shape 'Rounding' ((NumberText ([Math]::Min([double]$obj.w,[double]$obj.h)*0.12))+' pt') }
        }
        foreach($obj in $paths) {
            Stage ('draw_path:'+([string]$obj.id))
            if($obj.type -eq 'connector') {
                $shape=$page.Drop($app.ConnectorToolDataObject,0,0)
                $shape.NameU=[string]$obj.id
                $route='2'; if((Value $obj 'route' 'straight') -eq 'elbow') { $route='1' }
                Set-VCell $shape 'ShapeRouteStyle' $route
                Set-VCell $shape 'ConLineJumpCode' '1'
                $start=V-GluePoint ([string]$obj.from_side); $end=V-GluePoint ([string]$obj.to_side)
                $shape.CellsU('BeginX').GlueToPos($created[[string]$obj.from],[double]$start[0],[double]$start[1])
                $shape.CellsU('EndX').GlueToPos($created[[string]$obj.to],[double]$end[0],[double]$end[1])
            } else {
                $pts=@($obj.points)
                if($obj.type -eq 'line') { $shape=$page.DrawLine(([double]$pts[0][0]/72),(($height-[double]$pts[0][1])/72),([double]$pts[1][0]/72),(($height-[double]$pts[1][1])/72)) }
                else {
                    [double[]]$xy=New-Object double[] ($pts.Count*2)
                    for($i=0;$i -lt $pts.Count;$i++) { $xy[$i*2]=[double]$pts[$i][0]/72; $xy[$i*2+1]=($height-[double]$pts[$i][1])/72 }
                    $shape=$page.DrawPolyline([ref]$xy,8)
                }
                $shape.NameU=[string]$obj.id
            }
            $created[[string]$obj.id]=$shape; Set-VStyle $shape $obj $true $doc
        }
        $null=Invoke-ComCall 'save_native' { $doc.SaveAs($file) }
        $documentCount=Invoke-ComCall 'check_guard_requirement' { $app.Documents.Count }
        if($documentCount -eq 1) {
            # Keeping one owned document open avoids the observed last-document
            # close -> start page transition before the next Documents.Open call.
            $guardDocument=Invoke-ComCall 'create_owned_guard_document' { $app.Documents.Add('') }
            $script:CurrentEntry.guard.created=$true
            $script:CurrentEntry.guard.document_id=$guardDocument.ID
            $script:CurrentEntry.guard.origin='render_process'
            Stage 'owned_guard_ready'
        }
        $null=Invoke-ComCall 'close_saved_document' { $doc.Close() }; $doc=$null
        # Reopen only the new macro-free VSDX built above, never an arbitrary input file.
        # Reattach after isolated initialization and keep an owned guard where
        # needed; a same-client cold startup stalled here even with a guard.
        # OpenEx alone was not the cause of the reproduced host-state problem.
        # No application security or licensing setting is changed.
        $reopened=Invoke-ComCall 'reopen_saved_document' { $app.Documents.Open($file) }
        Stage 'verify_reopened_objects'
        $check=$reopened.Pages.Item(1)
        if($check.Shapes.Count -ne ($spec.Count+1)) { throw 'Visio saved shape count does not match scene plus native canvas background.' }
        $null=$check.Shapes.ItemU($CanvasBackgroundId)
        $attached=0
        foreach($obj in $scene.objects) {
            $shape=$check.Shapes.ItemU([string]$obj.id)
            if($obj.type -eq 'connector') {
                if($shape.Connects.Count -lt 2) { throw "Detached connector after reopening: $($obj.id)" }
                $attached++
            }
        }
        $script:CurrentEntry.native_objects=$check.Shapes.Count
        $script:CurrentEntry.attached_connectors=$attached
        $script:CurrentEntry.saved_reopened=$true
        $null=Invoke-ComCall 'export_png' { $check.Export($png) }
        if($ExportEmf) { $null=Invoke-ComCall 'export_emf' { $check.Export((Join-Path $OutputDir ($name+'_vsdx.emf'))) } }
        $script:CurrentEntry.status='native_saved_and_reopened'
        Stage 'target_verified'
    } catch {
        Failed-Target $_
    } finally {
        if($null -ne $reopened) { Cleanup-Call 'close_reopened_document' { $reopened.Close() } }
        if($null -ne $doc) { Cleanup-Call 'discard_owned_unsaved_document' { $doc.Saved=$true; $doc.Close() } }
        if($null -ne $guardDocument) {
            Cleanup-Call 'discard_owned_guard_document' {
                $guardDocument.Saved=$true
                $guardDocument.Close()
                $script:CurrentEntry.guard.closed=$true
            }
        }
        if($null -ne $app) { Cleanup-Call 'release_application_reference' { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) } }
        Stage 'target_finished'
    }
}
$missing=New-Object System.Collections.ArrayList
foreach($file in $files | Where-Object { $_ -ne $receiptPath }) {
    if(-not (Test-Path -LiteralPath $file) -or (Get-Item -LiteralPath $file).Length -eq 0) { $null=$missing.Add($file) }
}
$script:Receipt['missing_or_empty_outputs']=@($missing)
$successful=@($results | Where-Object { $_.status -eq 'native_saved_and_reopened' }).Count
if($successful -eq $formats.Count -and $missing.Count -eq 0) { $script:Receipt.status='native_saved_and_reopened' }
elseif($successful -gt 0) { $script:Receipt.status='partial' }
else { $script:Receipt.status='failed' }
Flush-Receipt
$script:Receipt | ConvertTo-Json -Depth 12
if($script:Receipt.status -ne 'native_saved_and_reopened') { exit 1 }
