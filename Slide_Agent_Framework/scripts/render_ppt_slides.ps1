$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $ppt.Presentations.Open('C:\Users\Giang\Desktop\slideA3A4\exports\slide_A4_hoan_thien_Editable.pptx', [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$pres.Slides.Item(1).Export('C:\Users\Giang\Desktop\slideA3A4\Slide_Agent_Framework\exports\real_ppt_slide_1.png', 'PNG', 1920, 1080)
$pres.Slides.Item(2).Export('C:\Users\Giang\Desktop\slideA3A4\Slide_Agent_Framework\exports\real_ppt_slide_2.png', 'PNG', 1920, 1080)
$pres.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Output "SUCCESS: Exported real PPTX slides to PNG!"
