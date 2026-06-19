#requires -Version 5.1
<# Created by Dewald Pretorius #>
param([string]$OutputPath)
if(-not $OutputPath){$OutputPath="$([Environment]::GetFolderPath('Desktop'))\Word_Print_Layout_Reports"};New-Item $OutputPath -ItemType Directory -Force|Out-Null
$printers=Get-Printer -ErrorAction SilentlyContinue|Select-Object Name,DriverName,PortName,PrinterStatus;$fonts=Get-ChildItem "$env:WINDIR\Fonts" -File -ErrorAction SilentlyContinue|Select-Object -First 200 Name
@('WORD PRINTING AND LAYOUT TROUBLESHOOTER','Created by Dewald Pretorius',"Generated: $(Get-Date)",'Review margins, page size, section breaks, printer driver, default printer, font substitution, and Print to PDF comparison.',($printers|Format-Table -AutoSize|Out-String -Width 220),($fonts|Format-Table -AutoSize|Out-String -Width 220))|Set-Content (Join-Path $OutputPath 'Report.txt') -Encoding UTF8