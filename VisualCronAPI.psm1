# NOTES:
# The VisualCron API uses WCF which isn't available in .NET Core so
# the PowerShell code must be run under Windows PowerShell and
# not PowerShell Core (pwsh).
#

<#
.SYNOPSIS
  Fetch a VisualCron API Connection object.

.DESCRIPTION
  Create the VisualCron API Connection object with the supplied params.

.PARAMETER Address
  The VisualCron server's address.

.PARAMETER Port
  The port number to connect on.

.PARAMETER ConnectionType
  The connection type is either 'Remote' or 'Local'.

.PARAMETER UserName
  User to connect as.

.PARAMETER Password
  The password for the user account.

.PARAMETER SecurePassword
  The password for the user account as a SecureString.

.PARAMETER Credential
  The PSCredential for the service account.

.PARAMETER Local
  A switch for when you only want to connect to your local server anonymously.

.PARAMETER UseADLogon
  A switch for when you need to use Active Directory logon.
  UserName and Password will be ignored.
  Do not specify at the same time as the Local flag as it isn't supported.

.INPUTS
  None.

.OUTPUTS
  VisualCronAPI.Connection.

.EXAMPLE
  PS> Get-VcConnection '10.110.0.135' 16444 'Remote' -UserName 'svc-user' -Password 'ar2@#$3'
  The complete set of positional parameters for when you have the plain-text password.

.EXAMPLE
  PS> Get-VcConnection -Address 'ServerName' -UserName 'svc-user' -Password 'ar2@#$3'
  As the previous example but now rely on the default values.

.EXAMPLE
  PS> Get-VcConnection -Address $Address -UserName $UserName -SecurePassword $SecurePassword
  Call with pre-existing variables and the SecureString password.

.EXAMPLE
  PS> Get-VcConnection '10.110.0.135' -Credential (Get-Credential)
  The PSCredential for the service account, entered into a dialog box.

.EXAMPLE
  PS> Get-VcConnection -Local
  A switch for when you only want to connect to your local server anonymously.

.LINK
  https://www.visualcron.com/doc/HTML/powershell.html
#>
function Get-VcConnection {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'Password', Justification = "It's not me, it's you.")]
  param (
    [Parameter(Position = 0)]
    [string] $Address = $Env:COMPUTERNAME, 

    [Parameter(Position = 1)]
    [int] $Port = 16444,

    [Parameter(Position = 2)]
    [string] $ConnectionType = 'Remote',

    [Parameter(Mandatory, Position = 3, ParameterSetName = 'Password')]
    [Parameter(Mandatory, Position = 3, ParameterSetName = 'SecurePassword')]
    [string] $UserName,

    [Parameter(Mandatory, ParameterSetName = 'Password')]
    [string] $Password,

    [Parameter(Mandatory, ParameterSetName = 'SecurePassword')]
    [SecureString] $SecurePassword,

    [Parameter(Mandatory, ParameterSetName = 'Credential')]
    [System.Management.Automation.PSCredential] $Credential,
    
    [Parameter(ParameterSetName = 'Local')]
    [switch] $Local,
    
    [Parameter(ParameterSetName = 'UseADLogon')]
    [switch] $UseADLogon)

  if ($SecurePassword) {
    $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
  }
  if ($Credential) {
    $Password = $Credential.GetNetworkCredential().Password
    $UserName = $Credential.UserName
  }

  $props = @{
    Address        = $Address
    Port           = $Port
    ConnectionType = $ConnectionType
  }
  if ($Local -eq $false -and $UseADLogon -eq $false) {
    $props["UserName"] = $UserName
    $props["PassWord"] = $Password
  }
  if ($UseADLogon -eq $true) {
    $props["UseADLogon"] = $true
  }
  New-Object VisualCronAPI.Connection -Property $props
}

<#
.SYNOPSIS
  Connect to a VisualCron API Server.

.DESCRIPTION
  Create the VisualCron API Server object with the supplied Connection object.

.INPUTS
  VisualCronAPI.Connection.

.OUTPUTS
  VisualCronAPI.Server.

.EXAMPLE
  PS> $server = $connection | Connect-VcServer

.LINK
  https://www.visualcron.com/doc/HTML/powershell.html
#>
function Connect-VcServer {
  param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [VisualCronAPI.Connection] $Connection)

  (New-Object VisualCronAPI.Client).Connect($Connection, $true)
}

<#
.SYNOPSIS
  Fetch all the Jobs from a VisualCron API Server.

.DESCRIPTION
  Fetch all the Jobs from the supplied Server.

.INPUTS
  VisualCronAPI.Server.

.OUTPUTS
  VisualCronAPI.Jobs list.

.EXAMPLE
  PS> $jobs = $server | Get-VcJob

.LINK
  https://www.visualcron.com/doc/HTML/powershell.html
#>
function Get-VcJobs {
  param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [VisualCronAPI.Server] $server, 
    [switch] $active)

  $server.Jobs.GetAll() `
  | Where-Object { !($active) -or $_.Stats.Active }
}

<#
.SYNOPSIS
  Fetch all the Tasks from a VisualCron API Server.

.DESCRIPTION
  Fetch all the Tasks from the supplied Server.
  The Tasks are decorated with the Job Name and Job Group.

.INPUTS
  VisualCronAPI.Server.

.OUTPUTS
  VisualCronAPI.Tasks list.

.EXAMPLE
  PS> $tasks = $server | Get-VcTasks

.LINK
  https://www.visualcron.com/doc/HTML/powershell.html
#>
function Get-VcTasks {
  param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [VisualCronAPI.Server] $server, 
    [switch] $active)

  Get-VCJobs @psBoundParameters `
  | ForEach-Object { 
    $job = $_
    $_.Tasks `
    | Where-Object { !($active) -or $_.Stats.Active } `
    | Sort-Object Order `
    | ForEach-Object { Add-Member NoteProperty   Job     -InputObject $_ $job -PassThru } `
    | ForEach-Object { Add-Member ScriptProperty JobName -InputObject $_ { $this.Job.Name } -PassThru } `
    | ForEach-Object { Add-Member ScriptProperty Group   -InputObject $_ { $this.Job.Group } -PassThru }
  }
}

function Get-UserVariable {
  param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [VisualCronAPI.Server] $server, 
    [string] $variableName)

  $server.Variables.GetGenericVariable('{uservar(' + $variableName + ')}')
}

function Get-ServerVariable {
  param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [VisualCronAPI.Server] $server, 
    [string] $variableName)

  $server.Variables.GetGenericVariable('{server(' + $variableName + ')}')
}

function Get-Variable {
  param (
    [Parameter(Mandatory, ValueFromPipeline)]
    [VisualCronAPI.Server] $server, 
    [string] $variableName)

  $server.Variables.GetGenericVariable('{' + $variableName + '}')
}