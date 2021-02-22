<#
.SYNOPSIS
  Visual Cron API access.

.DESCRIPTION
  Use the Visual Cron API to create documentation for the Jobs and Tasks.

.PARAMETER Application
  If an application name is provided then that application's Jobs will be exported.

.INPUTS
  None.

.OUTPUTS
  None.

.NOTES
  Version:        1.0
  Author:         Richard Bogle
  Creation Date:  2020-02-01
  Purpose/Change: Initial script development

.EXAMPLE
  PS> .\Export-VisualCronJobs.ps1

.EXAMPLE
  Export the Jobs for Electra.
  PS> .\Export-VisualCronJobs.ps1 'Electra'

.LINK
  https://www.visualcron.com/doc/HTML/powershell.html
#>

[CmdletBinding()]

PARAM (
  [string]$application
)

Function Get-VcAPIPath {
  $programFilesPath = if (${Env:PROCESSOR_ARCHITECTURE} -eq 'x86') { ${Env:ProgramFiles} } else { ${Env:ProgramFiles(x86)} }
  Join-Path $programFilesPath VisualCron\VisualCronAPI.dll
}

Function Get-VcServer {
  [CmdletBinding()]
  param ([string]$ComputerName, 
    [int]$Port,
    [System.Management.Automation.PSCredential]$Credential)

  $apiPath = Get-VCAPIPath
  if (!(Test-Path $apiPath)) { Throw "VisualCron does not appear to be installed. API library not found at `"$apiPath`"." }
  [Reflection.Assembly]::LoadFrom($apiPath) | Out-Null
  $conn = New-Object VisualCronAPI.Connection
  $conn.Address = if ([String]::IsNullOrEmpty($ComputerName)) { ${Env:COMPUTERNAME} } else { $ComputerName }
  if (!($credential -eq $null)) {
    $conn.UseADLogon = $true
    $netcred = $credential.GetNetworkCredential()
    $conn.UserName = $netcred.UserName
    $conn.Password = $netcred.Password
  }
  $client = New-Object VisualCronAPI.Client
  $client.Connect($conn)
}

Function Get-VcJob {
  [CmdletBinding()]
  param ([string]$ComputerName, 
    [int]$Port,
    [System.Management.Automation.PSCredential]$Credential,
    [switch]$Active)

  $ps = New-Object Collections.Hashtable($psBoundParameters)
  $ps.Remove('Active') | Out-Null
  $server = Get-VCServer @ps
  $server.Jobs.GetAll() `
  | Where-Object { !($Active) -or $_.Stats.Active } `
  | Add-Member ScriptMethod Start { $server.Jobs.Run($this, $false, $false, $false, $null) }.GetNewClosure() -PassThru
}
 
Function Get-VcTask {
  [CmdletBinding()]
  param ([string]$ComputerName, 
    [int]$Port,
    [System.Management.Automation.PSCredential]$Credential,
    [switch]$Active)

  Get-VCJob @psBoundParameters `
  | ForEach-Object { 
    $job = $_
    $_.Tasks `
    | Where-Object { !($Active) -or $_.Stats.Active } `
    | Sort-Object Order `
    | ForEach-Object { Add-Member NoteProperty   Job     -InputObject $_ $job                -PassThru } `
    | ForEach-Object { Add-Member ScriptProperty JobName -InputObject $_ { $this.Job.Name } -PassThru } `
    | ForEach-Object { Add-Member ScriptProperty Group   -InputObject $_ { $this.Job.Group } -PassThru }
  }
}

Function Get-VcTaskExecute {
  [CmdletBinding()]
  param ([string]$ComputerName, 
    [int]$Port,
    [System.Management.Automation.PSCredential]$Credential,
    [switch]$Active)

  Get-VCTask @psBoundParameters `
  | Where-Object { !($null -eq $_.Execute) } 
}

Function Get-VcTaskCommandLine {
  [CmdletBinding()]
  param ([string]$ComputerName, 
    [int]$Port,
    [System.Management.Automation.PSCredential]$Credential,
    [switch]$Active)

  Get-VCTaskExecute @psBoundParameters `
  | ForEach-Object { Add-Member ScriptProperty CmdLine   -InputObject $_ { $this.Execute.CmdLine } -PassThru } `
  | ForEach-Object { Add-Member ScriptProperty Arguments -InputObject $_ { $this.Execute.Arguments } -PassThru }
}