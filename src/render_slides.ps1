param(
    [string]$Source = "C:\Users\aleen\OneDrive\Desktop\OBI_PSO\AO_PSO_Presentation_v2.pptx",
    [string]$OutDir = "C:\Users\aleen\OneDrive\Desktop\OBI_PSO\src\_slidepng"
)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoCTrue
try {
    $pres = $ppt.Presentations.Open($Source, $true, $true, $false)
    foreach ($i in 1..$pres.Slides.Count) {
        $slide = $pres.Slides.Item($i)
        $out = Join-Path $OutDir ("slide_{0:D2}.png" -f $i)
        $slide.Export($out, "PNG", 1600, 900)
    }
    $pres.Close()
} finally {
    $ppt.Quit()
}
Write-Output "rendered to $OutDir"
