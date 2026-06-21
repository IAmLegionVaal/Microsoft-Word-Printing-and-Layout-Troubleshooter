# Microsoft Word Printing and Layout Troubleshooter

Created by **Dewald Pretorius**.

The repository includes the original print-layout, margin, pagination, font, section, and printer diagnostics plus a guarded `Repair.ps1` helper with `Diagnose`, `ResetOfficeCache`, and `StartPrintSpooler` actions.

```powershell
.\Repair.ps1 -Action Diagnose
.\Repair.ps1 -Action ResetOfficeCache -WhatIf
.\Repair.ps1 -Action StartPrintSpooler -Confirm
```

Close Word before cache repair. Existing Office cache data is preserved as a timestamped backup. The spooler action starts the service only when stopped and verifies that it is running. Source-reviewed for Windows PowerShell 5.1; not runtime-tested against every printer driver or Word build.
