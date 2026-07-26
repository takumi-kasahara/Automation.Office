using namespace System.Diagnostics.CodeAnalysis
using namespace System.IO
using namespace System.Management.Automation

Set-StrictMode -Version Latest

function New-ErrorRecord {
  [CmdletBinding()]
  [OutputType([System.Management.Automation.ErrorRecord])]
  [SuppressMessage('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Function creates error records but does not change state')]
  param (
    [Parameter(Mandatory)]
    [ValidateSet(
      'FileIsReadOnly'
      , 'ItemAlreadyExists'
      , 'ItemNotFound'
    )]
    [string]
    $ErrorId,
    [string]
    $TargetObject
  )
  switch -Exact -CaseSensitive ($ErrorId) {
    'FileIsReadOnly' {
      return [ErrorRecord]::new(
        [IOException]::new("$($TargetObject.FullName) is read-only.")
        , 'FileIsReadOnly'
        , [ErrorCategory]::PermissionDenied
        , $TargetObject
      )
    }
    'ItemAlreadyExists' {
      return [ErrorRecord]::new(
        [IOException]::new("$TargetObject already exists. Use -Force to overwrite the file.")
        , 'ItemAlreadyExists'
        , [ErrorCategory]::ResourceExists
        , $TargetObject
      )
    }
    'ItemNotFound' {
      return [ErrorRecord]::new(
        [ItemNotFoundException]::new("$TargetObject not found.")
        , 'ItemNotFound'
        , [ErrorCategory]::ObjectNotFound
        , $TargetObject
      )
    }
    default {
      throw [ArgumentException]::new("Invalid:$ErrorId", 'ErrorId')
    }
  }
}
function Start-NUIDialogSuppressor {
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
  }
  $job | Add-Member -MemberType NoteProperty -Name StopFilePath -Value $stopFilePath
  return $job
}
function Stop-NUIDialogSuppressor {
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
