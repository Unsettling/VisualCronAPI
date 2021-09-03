# NOTES:
# The VisualCron API uses WCF which isn't available in .NET Core so
# the PowerShell code must be run under Windows PowerShell and
# not PowerShell Core (pwsh).
#

using module .\VisualCronAPI.psd1
. .\Remove-NullEmptyProperties.ps1

function Show-VcJobs {
  $jobs = Get-VcConnection -UseADLogon | Connect-VcServer | Get-VcJobs
  $jobs | Remove-NullEmptyProperties | ConvertTo-Json -Depth 1
  $jobs.Tasks | Remove-NullEmptyProperties | ConvertTo-Json -Depth 1
}
