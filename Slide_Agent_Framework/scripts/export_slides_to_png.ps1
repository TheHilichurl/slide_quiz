param([string]$pptxPath, [string]$outPrefix)
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open((Resolve-Path $pptxPath).Path, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $outFile = [System.IO.Path]::GetFullPath("$outPrefix$i.png")
    $pres.Slides.Item($i).Export($outFile, 'PNG', 1920, 1080)
    Write-Output "Exported slide $i -> $outFile"
}
$pres.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
