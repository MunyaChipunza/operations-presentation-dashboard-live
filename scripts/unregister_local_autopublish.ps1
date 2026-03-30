[CmdletBinding()]
param(
    [string]$TaskName = "Operations PPT Dashboard Auto Publish"
)

$ErrorActionPreference = "Stop"
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
