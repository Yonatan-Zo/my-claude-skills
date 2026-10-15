# Export every slide of a deck to PNG with PowerPoint and report animation steps per slide.
# Usage: powershell -File scripts/render.ps1 -Deck output\green_roof_talk_v3.pptx
param([string]$Deck = "output\green_roof_talk_v3.pptx")
$root = Split-Path -Parent $PSScriptRoot
$deckPath = Join-Path $root $Deck
$out = Join-Path $env:TEMP "green_roof_talk_build\render"
New-Item -ItemType Directory -Force $out | Out-Null
Get-ChildItem $out -Filter "slide-*.png" | Remove-Item
$pp = New-Object -ComObject PowerPoint.Application
try {
    $pres = $pp.Presentations.Open($deckPath, $true, $false, $false)
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $sl = $pres.Slides.Item($i)
        $sl.Export((Join-Path $out ("slide-{0:D2}.png" -f $i)), "PNG", 1600, 900)
        "slide {0}: {1} animation effects" -f $i, $sl.TimeLine.MainSequence.Count
    }
    $pres.Close()
} finally {
    $pp.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
}
