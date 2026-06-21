#requires -Version 5.1
<# Created by Dewald Pretorius. #>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
  [ValidateSet('Diagnose','ResetOfficeCache','StartPrintSpooler')][string]$Action='Diagnose',
  [string]$OutputPath=(Join-Path ([Environment]::GetFolderPath('Desktop')) 'Word_Printing_Repair')
)
$ErrorActionPreference='Stop'
$cachePath="$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache"
New-Item -ItemType Directory -Path $OutputPath -Force|Out-Null
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss';$logPath=Join-Path $OutputPath "Repair_$stamp.log"
function Log([string]$Message){$line='{0:u} {1}' -f (Get-Date),$Message;Write-Host $line;Add-Content -LiteralPath $logPath -Value $line}
[ordered]@{Action=$Action;WordRunning=[bool](Get-Process WINWORD -ErrorAction SilentlyContinue);Spooler=(Get-Service Spooler -ErrorAction SilentlyContinue|Select-Object Name,Status,StartType);Printers=@(Get-Printer -ErrorAction SilentlyContinue|Select-Object Name,DriverName,PortName,PrinterStatus);CacheExists=(Test-Path -LiteralPath $cachePath)}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $OutputPath "PreRepair_$stamp.json") -Encoding UTF8
if($Action -eq 'Diagnose'){Log '[COMPLETE] Snapshot saved.';exit 0}
try{
  if($Action -eq 'ResetOfficeCache' -and $PSCmdlet.ShouldProcess($cachePath,'Back up and reset Office cache')){
    if(Get-Process WINWORD -ErrorAction SilentlyContinue){throw 'Close Word before resetting the cache.'}
    if(Test-Path -LiteralPath $cachePath){$backup="$cachePath.backup-$stamp";Move-Item -LiteralPath $cachePath -Destination $backup -Force;New-Item -ItemType Directory -Path $cachePath -Force|Out-Null;Log "[BACKUP] $backup"}
  }
  elseif($Action -eq 'StartPrintSpooler' -and $PSCmdlet.ShouldProcess('Print Spooler','Start if stopped')){
    $service=Get-Service Spooler
    if($service.Status -eq 'Stopped'){Start-Service Spooler}
    Start-Sleep -Seconds 2
    if((Get-Service Spooler).Status -ne 'Running'){throw 'Print Spooler is not running.'}
  }
}catch{Log "[FAILED] $($_.Exception.Message)";exit 5}
Log '[COMPLETE] Repair completed.'
exit 0
