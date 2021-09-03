# NOTES:
# The VisualCron API uses WCF which isn't available in .NET Core so
# the PowerShell code must be run under Windows PowerShell and
# not PowerShell Core (pwsh).
#

Add-Type -Path 'C:\Program Files (x86)\VisualCron\VisualCronAPI.dll'
Add-Type -Path 'C:\Program Files (x86)\VisualCron\VisualCron.dll'
New-Object VisualCronAPI.Connection -Property @{Address = '10.110.0.135'; UseADLogon = 'True' } `
| ForEach-Object { (New-Object VisualCronAPI.Client).Connect($_) } `
| ForEach-Object { $_.Jobs.GetAll() }