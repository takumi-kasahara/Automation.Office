using namespace System.Diagnostics.CodeAnalysis

function Start-DialogSuppressor {
  [CmdletBinding()]
  [OutputType([System.Management.Automation.Job])]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function starts a background job to suppress dialogs and does not require user confirmation')]
  param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]
    $TargetExe
  )
  $csPath = $PSScriptRoot | Join-Path -ChildPath 'DialogSuppressor.cs'
  $stopFilePath = $env:TEMP | Join-Path -ChildPath ("NUIDialogSuppressor.$([guid]::NewGuid().ToString('N')).stop")
  $job = Start-Job -ScriptBlock {
    param()
    Add-Type -LiteralPath $using:csPath
    [NUIDialogSuppressor]::Run($using:stopFilePath, $using:TargetExe)
    [VBProjectDialogSuppressor]::Run($using:stopFilePath, $using:TargetExe)
  }
  $job | Add-Member -MemberType NoteProperty -Name StopFilePath -Value $stopFilePath
  return $job
}
function Stop-DialogSuppressor {
  [CmdletBinding()]
  [OutputType([void])]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function stops a background job and performs cleanup, which is an internal state change')]
  param (
    [Parameter(Mandatory)]
    [System.Management.Automation.Job]
    $Job
  )
  $stopFilePath = $Job.StopFilePath
  try {
    New-Item -Path $stopFilePath -ItemType File -Force -WhatIf:$false -Confirm:$false | Out-Null
    Receive-Job -Job $Job -Wait -AutoRemoveJob
  } finally {
    if (Test-Path -LiteralPath $stopFilePath) {
      Remove-Item -LiteralPath $stopFilePath -Force -WhatIf:$false -Confirm:$false
    }
  }
}
