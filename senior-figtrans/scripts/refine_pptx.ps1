# Copyright (c) 2026 NCEPU-Senior. Author: XiaoN.
<#
.SYNOPSIS
Optional native PowerPoint refinement on a copy of an existing PPTX.
.DESCRIPTION
The JSON spec names top-level native components, receiver groups, closed local
line/cubic arrow paths, rear ports and observed/target formula baselines (pt).
Local arrow coordinates put the tip at [0,0], facing +x; direction rotates that
path right/left/down/up. Existing connectors are retargeted to invisible rear
ports, then the receiver, its editable head and its port are grouped together.
Only translations are covered by the optional move/edit/save/reopen proof.
This is PowerPoint-only. It does not add MathType, recognize formulas, measure
glyph baselines automatically, support Visio, or flatten anything into pictures.
Geometry outside the slide or larger than max_head_extent_pt is rejected, never
silently scaled or clipped. The default maximum head extent is one third of the
smaller slide dimension; an explicit limit may raise it up to that dimension.
All output collisions and the spec are checked before COM startup. SourceDeck
is never opened for editing. Only exact task-owned paths are closed. A unique
owned guard stays open across reopen calls; no Quit or process termination.
Native execution requires Windows PowerShell 5.1 (powershell.exe), not pwsh.
.EXAMPLE
powershell -NoProfile -File refine_pptx.ps1 -SourceDeck native.pptx -SpecPath refinement.json -OutputDir new-output -ValidateOnly
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$SourceDeck,
    [Parameter(Mandatory=$true)][string]$SpecPath,
    [Parameter(Mandatory=$true)][string]$OutputDir,
    [switch]$ValidateOnly
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'

function Get-Value($Object,[string]$Name,$Default) {
    $property=$Object.PSObject.Properties[$Name]
    if($null -eq $property -or $null -eq $property.Value){return $Default}
    return $property.Value
}
function Finite($Number,[string]$Name) {
    $value=[double]$Number
    if([double]::IsNaN($value) -or [double]::IsInfinity($value)){throw "Non-finite $Name"}
    return $value
}
function Pair($Point,[string]$Name) {
    if(@($Point).Count -ne 2){throw "$Name must be [x,y]"}
    return [double[]]@((Finite $Point[0] $Name),(Finite $Point[1] $Name))
}
function Baseline-Shift($Placement) {
    $observed=$Placement.PSObject.Properties['observed_baseline_pt'];$target=$Placement.PSObject.Properties['target_baseline_pt']
    if(($null -eq $observed) -ne ($null -eq $target)){throw 'Formula baseline measurements must provide both observed and target values'}
    if($null -eq $observed){return 0.0}
    return (Finite $target.Value 'target_baseline_pt')-(Finite $observed.Value 'observed_baseline_pt')
}
function Rotate-Point($Point,$Tip,[string]$Direction) {
    $x=[double]$Point[0];$y=[double]$Point[1]
    switch($Direction){
        right {return [double[]]@(($Tip[0]+$x),($Tip[1]+$y))}
        left  {return [double[]]@(($Tip[0]-$x),($Tip[1]-$y))}
        down  {return [double[]]@(($Tip[0]-$y),($Tip[1]+$x))}
        up    {return [double[]]@(($Tip[0]+$y),($Tip[1]-$x))}
        default {throw "Unsupported arrow direction: $Direction"}
    }
}
function Assert-InCanvas($Point,[double]$Width,[double]$Height,[string]$Name) {
    if($Point[0] -lt 0 -or $Point[1] -lt 0 -or $Point[0] -gt $Width -or $Point[1] -gt $Height){throw "Geometry outside slide: $Name"}
}
function Read-ZipXml($Zip,[string]$Part) {
    $entry=$Zip.GetEntry($Part)
    if($null -eq $entry){throw "Missing PPTX part: $Part"}
    $reader=New-Object IO.StreamReader($entry.Open())
    try {return [xml]$reader.ReadToEnd()} finally {$reader.Dispose()}
}
function Xml-Bounds($Element,$Namespaces) {
    $xfrm=$Element.SelectSingleNode('p:spPr/a:xfrm',$Namespaces)
    if($null -eq $xfrm){throw 'A named source component has no native transform'}
    $off=$xfrm.SelectSingleNode('a:off',$Namespaces);$ext=$xfrm.SelectSingleNode('a:ext',$Namespaces)
    if($null -eq $off -or $null -eq $ext){throw 'A named source component has an incomplete transform'}
    return [double[]]@(([double]$off.x/12700),([double]$off.y/12700),([double]$ext.cx/12700),([double]$ext.cy/12700))
}
function Bounds($Shape) {
    return [ordered]@{id=[long]$Shape.Id;left=[double]$Shape.Left;top=[double]$Shape.Top;
        width=[double]$Shape.Width;height=[double]$Shape.Height;rotation=[double]$Shape.Rotation}
}
function Find-Shape($Shapes,[string]$Name) {
    for($i=1;$i -le $Shapes.Count;$i++){
        $shape=$Shapes.Item($i)
        if([string]$shape.Name -ceq $Name){return $shape}
        if([int]$shape.Type -eq 6){$found=Find-Shape $shape.GroupItems $Name;if($null -ne $found){return $found}}
    }
    return $null
}
function Endpoint($Shape,[bool]$End) {
    $w=[double]$Shape.Width;$h=[double]$Shape.Height
    $x=-$w/2;$y=-$h/2
    if($End){$x=$w/2;$y=$h/2}
    if([int]$Shape.HorizontalFlip -eq -1){$x=-$x}
    if([int]$Shape.VerticalFlip -eq -1){$y=-$y}
    $angle=[double]$Shape.Rotation*[Math]::PI/180
    return [double[]]@(($Shape.Left+$w/2+$x*[Math]::Cos($angle)-$y*[Math]::Sin($angle)),
        ($Shape.Top+$h/2+$x*[Math]::Sin($angle)+$y*[Math]::Cos($angle)))
}
function Flush {
    $script:receipt.updated_utc=[DateTime]::UtcNow.ToString('o')
    $script:receipt | ConvertTo-Json -Depth 24 | Set-Content -LiteralPath $receiptPath -Encoding UTF8
}
function Stage([string]$Name){$script:receipt.stage=$Name;Flush}
function Get-HostApp {
    # A running but unregistered host is not permission to start another process.
    try {return [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')}
    catch {
        if(@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count){throw 'A running PowerPoint host is not registered in ROT. Preserve it; restore its task-owned session before retrying.'}
        return New-Object -ComObject PowerPoint.Application
    }
}
function Open-Owned([string]$Path,[int]$ReadOnly) {
    if($Path -notin $ownedPaths){throw 'Refusing to open an unowned output path'}
    $hostApp=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    $foundGuard=$false
    for($i=1;$i -le $hostApp.Presentations.Count;$i++){if([string]$hostApp.Presentations.Item($i).FullName -ieq $guardPath){$foundGuard=$true;break}}
    if(!$foundGuard){throw 'The exact task-owned guard is absent from the active PowerPoint host'}
    return $hostApp.Presentations.Open($Path,$ReadOnly,0,0)
}
function Close-Owned($Document,[string]$Path,[bool]$Saved) {
    if($Path -notin $ownedPaths -or [string]$Document.FullName -ine $Path){throw 'Document ownership check failed; document left open'}
    if($Saved -and (!(Test-Path -LiteralPath $Path -PathType Leaf) -or !(Get-Item -LiteralPath $Path).Length)){throw 'Saved output is missing; document left open'}
    $Document.Saved=-1
    $Document.Close()
}
function Export-Owned($Document,[string]$Path) {
    $scale=3
    $Document.Slides.Item($slideNumber).Export($Path,'PNG',[int]($width*$scale),[int]($height*$scale))
}
function Snapshot($Slide) {
    $links=@(foreach($entry in $sourceConnectors){
        $shape=Find-Shape $Slide.Shapes $entry.name;$connection=$shape.ConnectorFormat
        if($connection.BeginConnected -ne -1 -or $connection.EndConnected -ne -1){throw "Detached connector: $($entry.name)"}
        $endName=$entry.end_name
        if($arrowByName.ContainsKey($entry.name)){$endName='sft_port_'+$entry.name}
        if([string]$connection.BeginConnectedShape.Name -cne $entry.begin_name -or [string]$connection.EndConnectedShape.Name -cne $endName){throw "Wrong connector target: $($entry.name)"}
        if([long]$connection.BeginConnectedShape.Id -ne $entry.begin_id -or [int]$connection.BeginConnectionSite -ne $entry.begin_site){throw "Source begin ID/site changed: $($entry.name)"}
        if($arrowByName.ContainsKey($entry.name)){
            $expected=$generatedPorts[$entry.name]
            if([long]$connection.EndConnectedShape.Id -ne $expected.id -or [int]$connection.EndConnectionSite -ne $expected.site){throw "Generated rear port ID/site changed: $($entry.name)"}
        } elseif([long]$connection.EndConnectedShape.Id -ne $entry.end_id -or [int]$connection.EndConnectionSite -ne $entry.end_site){throw "Source end ID/site changed: $($entry.name)"}
        [ordered]@{name=$entry.name;id=[long]$shape.Id;begin_id=[long]$connection.BeginConnectedShape.Id;
            begin_name=[string]$connection.BeginConnectedShape.Name;begin_site=[int]$connection.BeginConnectionSite;
            end_id=[long]$connection.EndConnectedShape.Id;end_name=[string]$connection.EndConnectedShape.Name;
            end_site=[int]$connection.EndConnectionSite;begin_pt=(Endpoint $shape $false);end_pt=(Endpoint $shape $true)}
    })
    $groupRows=@(foreach($groupSpec in $groups){
        $group=Find-Shape $Slide.Shapes ([string]$groupSpec.id)
        $members=@(for($i=1;$i -le $group.GroupItems.Count;$i++){ $member=$group.GroupItems.Item($i);[ordered]@{name=[string]$member.Name;id=[long]$member.Id} })
        if($members.Count -ne $expectedMembers[[string]$groupSpec.id].Count){throw "Receiver group lost native members: $($groupSpec.id)"}
        foreach($member in $expectedMembers[[string]$groupSpec.id]){if($member -cnotin @($members | ForEach-Object name)){throw "Missing editable group member: $member"}}
        [ordered]@{name=[string]$group.Name;bounds=(Bounds $group);members=$members}
    })
    $text=$null
    if($null -ne $editText){$shape=Find-Shape $Slide.Shapes ([string]$editText.id);$text=[ordered]@{name=[string]$shape.Name;id=[long]$shape.Id;value=([string]$shape.TextFrame.TextRange.Text).Replace("`r","`n")}}
    return [ordered]@{connectors=$links;groups=$groupRows;text=$text}
}
function Assert-Edit($State,$Before) {
    $groupBefore=@($Before.groups | Where-Object name -CEQ ([string]$edit.move))[0]
    $groupAfter=@($State.groups | Where-Object name -CEQ ([string]$edit.move))[0]
    if($groupBefore.bounds.id -ne $groupAfter.bounds.id -or [Math]::Abs($groupAfter.bounds.left-$groupBefore.bounds.left-$move[0]) -gt 0.025 -or [Math]::Abs($groupAfter.bounds.top-$groupBefore.bounds.top-$move[1]) -gt 0.025){throw 'Receiver group did not preserve its ID and requested translation'}
    $movingIds=@($groupBefore.members | ForEach-Object id)+@($groupBefore.bounds.id)
    for($i=0;$i -lt $Before.connectors.Count;$i++){
        $prior=$Before.connectors[$i];$now=$State.connectors[$i]
        foreach($key in @('id','name','begin_id','end_id','begin_name','end_name','begin_site','end_site')){if($now[$key] -cne $prior[$key]){throw "Changed $key on $($prior.name)"}}
        foreach($end in @('begin','end')){
            $expected=@(0,0);if($prior[$end+'_id'] -in $movingIds){$expected=$move}
            for($axis=0;$axis -lt 2;$axis++){if([Math]::Abs($now[$end+'_pt'][$axis]-$prior[$end+'_pt'][$axis]-$expected[$axis]) -gt 0.025){throw "Endpoint did not follow receiver or stationary target: $($prior.name)/$end"}}
        }
    }
    foreach($group in $Before.groups){
        $updated=@($State.groups | Where-Object name -CEQ $group.name)[0]
        if(($group.members | ConvertTo-Json -Compress) -cne ($updated.members | ConvertTo-Json -Compress)){throw "Receiver group membership changed: $($group.name)"}
    }
    if($null -ne $editText -and ($State.text.id -ne $Before.text.id -or $State.text.value -cne ([string]$editText.value).Replace("`r","`n"))){throw 'Edited native text or its object ID was not preserved'}
}

# Validate source native names, all numerical geometry, topology and output paths
# before obtaining a COM application or modifying any document.
$SourceDeck=[IO.Path]::GetFullPath($SourceDeck);$SpecPath=[IO.Path]::GetFullPath($SpecPath);$OutputDir=[IO.Path]::GetFullPath($OutputDir)
if([IO.Path]::GetExtension($SourceDeck) -ine '.pptx'){throw 'SourceDeck must be a PPTX'}
$config=Get-Content -LiteralPath $SpecPath -Raw -Encoding UTF8 | ConvertFrom-Json
$name=[string](Get-Value $config 'name' 'refined')
if($name -notmatch '^[A-Za-z0-9][A-Za-z0-9_.-]*$' -or $name.EndsWith('.')){throw 'name must be a simple ASCII filename stem'}
$slideNumber=[int](Get-Value $config 'slide' 1)
$groups=@(Get-Value $config 'receiver_groups' @());$arrows=@(Get-Value $config 'arrowheads' @());$placements=@(Get-Value $config 'formula_placements' @())
$edit=Get-Value $config 'edit_check' $null;$editText=$null;$move=@(0,0)
$native=@{};$byId=@{};$topLevel=@{};$sourceConnectors=@()
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip=[IO.Compression.ZipFile]::OpenRead($SourceDeck)
try {
    $presentation=Read-ZipXml $zip 'ppt/presentation.xml';$rels=Read-ZipXml $zip 'ppt/_rels/presentation.xml.rels'
    $width=[double]$presentation.presentation.sldSz.cx/12700;$height=[double]$presentation.presentation.sldSz.cy/12700
    $slideIds=@($presentation.presentation.sldIdLst.sldId)
    if($slideNumber -lt 1 -or $slideNumber -gt $slideIds.Count){throw 'Requested slide is absent'}
    $relationshipId=$slideIds[$slideNumber-1].GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $relationship=@($rels.Relationships.Relationship | Where-Object Id -CEQ $relationshipId)[0]
    $part=([Uri]::new([Uri]'http://package/ppt/presentation.xml',[string]$relationship.Target)).AbsolutePath.TrimStart('/')
    $xml=Read-ZipXml $zip $part
    $ns=New-Object Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('p','http://schemas.openxmlformats.org/presentationml/2006/main');$ns.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
    $tree=$xml.SelectSingleNode('/p:sld/p:cSld/p:spTree',$ns)
    foreach($element in $tree.SelectNodes('.//p:sp|.//p:cxnSp|.//p:grpSp|.//p:pic|.//p:graphicFrame',$ns)){
        $identity=$element.SelectSingleNode('p:nvSpPr/p:cNvPr|p:nvCxnSpPr/p:cNvPr|p:nvGrpSpPr/p:cNvPr|p:nvPicPr/p:cNvPr|p:nvGraphicFramePr/p:cNvPr',$ns)
        $shapeName=[string]$identity.name;$shapeId=[string]$identity.id
        if($native.ContainsKey($shapeName) -or $byId.ContainsKey($shapeId)){throw 'Source has duplicate native shape names or IDs'}
        $native[$shapeName]=$element;$byId[$shapeId]=$shapeName
        if($element.ParentNode -eq $tree){$topLevel[$shapeName]=$element}
    }
    foreach($entry in $topLevel.GetEnumerator()){
        if($entry.Value.LocalName -ne 'cxnSp'){continue}
        $begin=@($entry.Value.SelectNodes('p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn',$ns));$end=@($entry.Value.SelectNodes('p:nvCxnSpPr/p:cNvCxnSpPr/a:endCxn',$ns))
        if($begin.Count -eq 0 -and $end.Count -eq 0){continue}
        if($begin.Count -ne 1 -or $end.Count -ne 1 -or !$byId.ContainsKey([string]$begin[0].id) -or !$byId.ContainsKey([string]$end[0].id)){throw "Invalid or one-ended source connector: $($entry.Key)"}
        $sourceConnectors+=@([ordered]@{name=$entry.Key;begin_name=$byId[[string]$begin[0].id];end_name=$byId[[string]$end[0].id];
            begin_id=[long]$begin[0].id;end_id=[long]$end[0].id;begin_site=([int]$begin[0].idx+1);end_site=([int]$end[0].idx+1)})
    }
    $groupByName=@{};$groupOwner=@{};$arrowByName=@{};$groupBounds=@{}
    foreach($group in $groups){
        $id=[string]$group.id
        if($id -notmatch '^[A-Za-z][A-Za-z0-9_.-]*$' -or $native.ContainsKey($id) -or $groupByName.ContainsKey($id)){throw "Invalid, duplicate or existing receiver group: $id"}
        if(@($group.members).Count -lt 1){throw "Receiver group has no members: $id"}
        $bounds=@($width,$height,0.0,0.0)
        foreach($member in $group.members){
            $member=[string]$member
            if(!$topLevel.ContainsKey($member) -or $topLevel[$member].LocalName -notin @('sp','cxnSp') -or $groupOwner.ContainsKey($member)){throw "Receiver members must be distinct top-level native shapes: $member"}
            if([string]$topLevel[$member].SelectSingleNode('.//p:cNvPr',$ns).name -cne $member){throw "Receiver member name has incorrect case: $member"}
            $memberTransform=$topLevel[$member].SelectSingleNode('p:spPr/a:xfrm',$ns)
            if(($memberTransform.HasAttribute('rot') -and [int]$memberTransform.rot -ne 0) -or $memberTransform.GetAttribute('flipH') -in @('1','true') -or $memberTransform.GetAttribute('flipV') -in @('1','true')){throw "Initial receiver members must be unrotated and unflipped: $member"}
            if($member -in @($sourceConnectors | ForEach-Object name)){throw 'Attached connectors stay outside receiver groups'}
            $groupOwner[$member]=$id;$b=Xml-Bounds $topLevel[$member] $ns
            $bounds[0]=[Math]::Min($bounds[0],$b[0]);$bounds[1]=[Math]::Min($bounds[1],$b[1]);$bounds[2]=[Math]::Max($bounds[2],($b[0]+$b[2]));$bounds[3]=[Math]::Max($bounds[3],($b[1]+$b[3]))
        }
        $groupByName[$id]=$group;$groupBounds[$id]=$bounds
    }
    $placed=@{}
    foreach($placement in $placements){
        $id=[string]$placement.id
        if($placed.ContainsKey($id)){throw "Duplicate formula placement: $id"};$placed[$id]=$true
        if(!$topLevel.ContainsKey($id) -or $topLevel[$id].LocalName -ne 'sp' -or $null -eq $topLevel[$id].SelectSingleNode('p:txBody',$ns)){throw "Formula component must be native top-level text: $id"}
        if([string]::IsNullOrWhiteSpace([string]$placement.evidence)){throw "Formula placement needs observed evidence: $id"}
        $shift=Baseline-Shift $placement
        $dx=Finite (Get-Value $placement 'delta_x_pt' 0) 'delta_x_pt';$b=Xml-Bounds $topLevel[$id] $ns
        Assert-InCanvas @(($b[0]+$dx),($b[1]+$shift)) $width $height $id
        Assert-InCanvas @(($b[0]+$b[2]+$dx),($b[1]+$b[3]+$shift)) $width $height $id
        if($groupOwner.ContainsKey($id)){$gb=$groupBounds[$groupOwner[$id]];$gb[0]=[Math]::Min($gb[0],($b[0]+$dx));$gb[1]=[Math]::Min($gb[1],($b[1]+$shift));$gb[2]=[Math]::Max($gb[2],($b[0]+$b[2]+$dx));$gb[3]=[Math]::Max($gb[3],($b[1]+$b[3]+$shift))}
    }
    $maxExtent=Finite (Get-Value $config 'max_head_extent_pt' ([Math]::Min($width,$height)/3)) 'max_head_extent_pt'
    if($maxExtent -le 0 -or $maxExtent -gt [Math]::Min($width,$height)){throw 'max_head_extent_pt must be positive and no larger than the smaller slide dimension'}
    foreach($arrow in $arrows){
        $id=[string]$arrow.connector;$receiver=[string]$arrow.receiver_group
        $link=@($sourceConnectors | Where-Object name -CEQ $id)
        if($link.Count -ne 1 -or $arrowByName.ContainsKey($id) -or !$groupByName.ContainsKey($receiver)){throw "Invalid arrow connector or receiver: $id"}
        if($link[0].end_name -notin $groupByName[$receiver].members){throw "Receiver group does not own the connector destination: $id"}
        foreach($generated in @('sft_head_'+$id,'sft_port_'+$id)){if($native.ContainsKey($generated) -or $groupByName.ContainsKey($generated)){throw "Reserved generated name exists: $generated"}}
        $direction=[string]$arrow.direction;$tip=Pair $arrow.tip 'tip';$rear=Pair $arrow.rear_offset 'rear_offset'
        if($direction -notin @('right','left','down','up')){throw "Unsupported arrow direction: $id"}
        $points=@();$start=Pair $arrow.path.start 'path.start';$points+=,$start;$last=$start;$hasTip=$false
        $segments=@($arrow.path.segments)
        if($segments.Count -lt 2 -or $segments.Count -gt 64){throw 'A closed arrow path needs 2..64 segments'}
        foreach($segment in $segments){
            if($segment.type -notin @('line','cubic')){throw "Unsupported path segment on $id"}
            if($segment.type -eq 'cubic'){$points+=,(Pair $segment.control1 'control1');$points+=,(Pair $segment.control2 'control2')}
            $last=Pair $segment.to 'to';$points+=,$last
            if([Math]::Abs($last[0])+[Math]::Abs($last[1]) -lt 0.00001){$hasTip=$true}
        }
        if(!$hasTip -or [Math]::Abs($last[0]-$start[0])+[Math]::Abs($last[1]-$start[1]) -gt 0.00001){throw "Arrow path must include tip [0,0] and close at its start: $id"}
        $xs=@($points | ForEach-Object {$_[0]});$ys=@($points | ForEach-Object {$_[1]})
        $minX=($xs | Measure-Object -Minimum).Minimum;$maxX=($xs | Measure-Object -Maximum).Maximum;$minY=($ys | Measure-Object -Minimum).Minimum;$maxY=($ys | Measure-Object -Maximum).Maximum
        if($maxX -gt 0.00001 -or $maxX-$minX -le 0 -or $maxY-$minY -le 0 -or $maxX-$minX -gt $maxExtent -or $maxY-$minY -gt $maxExtent){throw "Excessive or invalid custom head geometry: $id"}
        if($rear[0] -ge 0 -or $rear[0] -lt $minX -or $rear[1] -lt $minY -or $rear[1] -gt $maxY){throw "Rear port must lie inside the head bounds behind its tip: $id"}
        $portSize=Finite (Get-Value $arrow 'port_size' 0.25) 'port_size'
        if($portSize -le 0 -or $portSize -gt [Math]::Min(($maxX-$minX),($maxY-$minY))){throw "Invalid port_size: $id"}
        if([string](Get-Value $arrow 'fill' '#000000') -notmatch '^#[0-9A-Fa-f]{6}$'){throw 'Arrow fill must be #RRGGBB'}
        $gb=$groupBounds[$receiver]
        foreach($point in $points){$p=Rotate-Point $point $tip $direction;Assert-InCanvas $p $width $height $id;$gb[0]=[Math]::Min($gb[0],$p[0]);$gb[1]=[Math]::Min($gb[1],$p[1]);$gb[2]=[Math]::Max($gb[2],$p[0]);$gb[3]=[Math]::Max($gb[3],$p[1])}
        $arrowByName[$id]=$arrow
    }
    foreach($group in $groups){
        $headCount=@($arrows | Where-Object receiver_group -CEQ ([string]$group.id)).Count
        if(@($group.members).Count+2*$headCount -lt 2){throw 'A native receiver group requires at least two final members'}
    }
    if($null -ne $edit){
        if(!$groupByName.ContainsKey([string]$edit.move)){throw 'edit_check.move must name a receiver group'}
        $move=Pair $edit.delta 'edit_check.delta';if([Math]::Abs($move[0])+[Math]::Abs($move[1]) -lt 0.001){throw 'edit_check must specify a nonzero translation'}
        $gb=$groupBounds[[string]$edit.move];Assert-InCanvas @(($gb[0]+$move[0]),($gb[1]+$move[1])) $width $height 'edit_check';Assert-InCanvas @(($gb[2]+$move[0]),($gb[3]+$move[1])) $width $height 'edit_check'
        $editText=Get-Value $edit 'text' $null
        if($null -ne $editText -and ([string]$editText.id -notin $groupByName[[string]$edit.move].members -or $null -eq $native[[string]$editText.id].SelectSingleNode('p:txBody',$ns))){throw 'edit_check.text must name native text within the moved receiver group'}
    }
} finally {$zip.Dispose()}
$file=Join-Path $OutputDir ($name+'.pptx');$png=Join-Path $OutputDir ($name+'_pptx.png');$edited=Join-Path $OutputDir ($name+'_edited.pptx');$editedPng=Join-Path $OutputDir ($name+'_edited_pptx.png');$receiptPath=Join-Path $OutputDir ($name+'_refine_receipt.json')
$planned=@($file,$png,$receiptPath);if($null -ne $edit){$planned+=@($edited,$editedPng)}
foreach($path in $planned){if(Test-Path -LiteralPath $path){throw "Refusing to overwrite: $path"}}
if($SourceDeck -in $planned){throw 'Source and output paths must differ'}
if($ValidateOnly){[ordered]@{status='schema_valid';target='pptx';receiver_groups=$groups.Count;custom_heads=$arrows.Count;formula_placements=$placements.Count;planned_files=$planned} | ConvertTo-Json -Depth 6;exit 0}
if($PSVersionTable.PSEdition -ne 'Desktop'){throw 'Native execution requires Windows PowerShell 5.1 (powershell.exe)'}
$null=New-Item -ItemType Directory -Path $OutputDir -Force
$guardPath=Join-Path $OutputDir ('.sft-guard-'+[guid]::NewGuid().ToString('N')+'.pptx')
$ownedPaths=@($file,$edited,$guardPath);$app=$null;$guard=$null;$doc=$null;$currentPath=$null
$script:receipt=[ordered]@{status='running';stage='initialize';source=$SourceDeck;spec=$SpecPath;target='pptx';updated_utc=$null;renderer='Microsoft PowerPoint';
    source_opened_for_editing=$false;source_sha256=(Get-FileHash -LiteralPath $SourceDeck -Algorithm SHA256).Hash;output=$file;preview=$png;edited_output=$null;edited_preview=$null;host=[ordered]@{version=$null;hwnd=$null;initial_presentations=$null};guard=[ordered]@{path=$guardPath;created=$false;closed=$false};
    formula_placements=$placements;arrowheads=@();before=$null;after_edit=$null;after_reopen=$null;error=$null;cleanup_errors=@();visual_review='required'}
Flush
try {
    Stage 'acquire_application';$app=Get-HostApp
    $script:receipt.host.version=[string]$app.Version;$script:receipt.host.initial_presentations=[int]$app.Presentations.Count
    try {$script:receipt.host.hwnd=[long]$app.HWND} catch {}
    Stage 'create_owned_guard';$guard=$app.Presentations.Add(0);$null=$guard.Slides.Add(1,12);$guard.SaveAs($guardPath,24);$script:receipt.guard.created=$true
    Copy-Item -LiteralPath $SourceDeck -Destination $file
    Stage 'open_output_copy';$doc=Open-Owned $file 0;$currentPath=$file;$slide=$doc.Slides.Item($slideNumber)
    foreach($placement in $placements){
        Stage ('place_formula:'+([string]$placement.id));$shape=Find-Shape $slide.Shapes ([string]$placement.id)
        $shape.IncrementLeft([single](Get-Value $placement 'delta_x_pt' 0));$shape.IncrementTop([single](Baseline-Shift $placement))
    }
    $expectedMembers=@{};$generatedPorts=@{};foreach($group in $groups){$expectedMembers[[string]$group.id]=@($group.members)}
    foreach($arrow in $arrows){
        $id=[string]$arrow.connector;Stage ('custom_arrow:'+ $id)
        $tip=Pair $arrow.tip 'tip';$direction=[string]$arrow.direction;$start=Rotate-Point $arrow.path.start $tip $direction
        $builder=$slide.Shapes.BuildFreeform(1,[single]$start[0],[single]$start[1])
        foreach($segment in $arrow.path.segments){
            $to=Rotate-Point $segment.to $tip $direction
            if($segment.type -eq 'line'){$builder.AddNodes(0,1,[single]$to[0],[single]$to[1])}
            else {$c1=Rotate-Point $segment.control1 $tip $direction;$c2=Rotate-Point $segment.control2 $tip $direction;$builder.AddNodes(1,1,[single]$c1[0],[single]$c1[1],[single]$c2[0],[single]$c2[1],[single]$to[0],[single]$to[1])}
        }
        $head=$builder.ConvertToShape();$head.Name='sft_head_'+$id;$head.Fill.Solid();$head.Line.Visible=0
        $color=[string](Get-Value $arrow 'fill' '#000000');$head.Fill.ForeColor.RGB=([Convert]::ToInt32($color.Substring(1,2),16)+256*[Convert]::ToInt32($color.Substring(3,2),16)+65536*[Convert]::ToInt32($color.Substring(5,2),16))
        $rear=Rotate-Point $arrow.rear_offset $tip $direction;$size=[single](Get-Value $arrow 'port_size' 0.25)
        $px=$rear[0];$py=$rear[1]
        switch($direction){right {$py-=$size/2;$site=2};left {$px-=$size;$py-=$size/2;$site=4};down {$px-=$size/2;$site=1};up {$px-=$size/2;$py-=$size;$site=3}}
        $port=$slide.Shapes.AddShape(1,[single]$px,[single]$py,$size,$size);$port.Name='sft_port_'+$id;$port.Fill.Visible=0;$port.Line.Visible=0
        $line=Find-Shape $slide.Shapes $id;$line.Line.EndArrowheadStyle=1;$line.ConnectorFormat.EndDisconnect();$line.ConnectorFormat.EndConnect($port,$site)
        $expectedMembers[[string]$arrow.receiver_group]+=@([string]$head.Name,[string]$port.Name)
        $generatedPorts[$id]=[ordered]@{id=[long]$port.Id;site=$site}
        $script:receipt.arrowheads+=@([ordered]@{connector=$id;receiver_group=[string]$arrow.receiver_group;head=(Bounds $head);port=(Bounds $port);tip_pt=$tip;rear_pt=$rear;port_site=$site})
    }
    foreach($group in $groups){Stage ('group_receiver:'+([string]$group.id));$shape=$slide.Shapes.Range([object[]]$expectedMembers[[string]$group.id]).Group();$shape.Name=[string]$group.id}
    Stage 'save_refined';$doc.Save();Close-Owned $doc $file $true;$doc=$null
    Stage 'reopen_refined';$doc=Open-Owned $file -1;$currentPath=$file;$script:receipt.before=Snapshot ($doc.Slides.Item($slideNumber));Export-Owned $doc $png;Close-Owned $doc $file $true;$doc=$null
    if($null -ne $edit){
        Copy-Item -LiteralPath $file -Destination $edited;Stage 'open_edit_copy';$doc=Open-Owned $edited 0;$currentPath=$edited;$slide=$doc.Slides.Item($slideNumber)
        if($null -ne $editText){$shape=Find-Shape $slide.Shapes ([string]$editText.id);$shape.TextFrame.TextRange.Text=[string]$editText.value}
        $group=Find-Shape $slide.Shapes ([string]$edit.move);$group.IncrementLeft([single]$move[0]);$group.IncrementTop([single]$move[1])
        $script:receipt.after_edit=Snapshot $slide;Assert-Edit $script:receipt.after_edit $script:receipt.before
        Stage 'save_edited';$doc.Save();Close-Owned $doc $edited $true;$doc=$null
        Stage 'reopen_edited';$doc=Open-Owned $edited -1;$currentPath=$edited;$script:receipt.after_reopen=Snapshot ($doc.Slides.Item($slideNumber));Assert-Edit $script:receipt.after_reopen $script:receipt.before;Export-Owned $doc $editedPng
        $script:receipt.edited_output=$edited;$script:receipt.edited_preview=$editedPng
    }
    $script:receipt.status='native_saved_and_reopened';Stage 'verified'
} catch {$script:receipt.status='failed';$script:receipt.error=[ordered]@{stage=$script:receipt.stage;message=$_.Exception.Message;stack=$_.ScriptStackTrace};throw}
finally {
    if($null -ne $doc){try {Close-Owned $doc $currentPath $false} catch {$script:receipt.cleanup_errors+=@($_.Exception.Message)}}
    if($null -ne $guard){try {Close-Owned $guard $guardPath $true;$script:receipt.guard.closed=$true} catch {$script:receipt.cleanup_errors+=@($_.Exception.Message)}}
    if($null -ne $app){try {$null=[Runtime.InteropServices.Marshal]::ReleaseComObject($app)} catch {$script:receipt.cleanup_errors+=@($_.Exception.Message)}}
    if($script:receipt.cleanup_errors.Count -gt 0 -and $script:receipt.status -ne 'failed'){$script:receipt.status='cleanup_failed'}
    Flush
}
if($script:receipt.status -ne 'native_saved_and_reopened'){exit 1}
