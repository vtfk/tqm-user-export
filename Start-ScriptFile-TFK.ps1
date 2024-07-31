#$scriptName = "TQM"
#$folder = "\\VT-A-SRV-BAS01\c$\Windows\Logs\AutoRun\$scriptName"
$script = Get-ChildItem -Path "$PSScriptRoot\*" -Include "Start-UserExport-TFK.ps1" | Select -First 1 | Select -ExpandProperty FullName

#$items = Get-ChildItem $folder
<#
if($items.Count -eq 0)
{
    Write-Host "Nothing todo!" -ForegroundColor Green -BackgroundColor Magenta
    return
}
#>
# delete autorun items (Får nå se hvor hardt den under tryner da....) lar den ligge kommentert ut inntil videre
#$items | Remove-Item -Force

# run $scriptName script
Write-Host "Running $scriptName script!" -ForegroundColor Green -BackgroundColor Magenta
Start-Process powershell -ArgumentList "-NoLogo -ExecutionPolicy Bypass -File $script" -Wait
