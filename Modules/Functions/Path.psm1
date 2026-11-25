using namespace Microsoft.PowerShell.Commands

Set-StrictMode -Version Latest

function Wait-Path {
  <#
  .SYNOPSIS
    Waits until a path exists.

  .DESCRIPTION
    Repeats Test-Path until it returns $true or the number of attempts reaches Count.
    Sleeps between attempts with Start-Sleep. Files created by external programs may not exist immediately after the command returns.

  .PARAMETER Path
    Specifies the path to test. Wildcards are permitted. Same as Test-Path -Path.

  .PARAMETER LiteralPath
    Specifies the path to test without wildcard interpretation. Same as Test-Path -LiteralPath.

  .PARAMETER Filter
    Same as Test-Path -Filter.

  .PARAMETER Include
    Same as Test-Path -Include.

  .PARAMETER Exclude
    Same as Test-Path -Exclude.

  .PARAMETER PathType
    Same as Test-Path -PathType.

  .PARAMETER NewerThan
    Same as Test-Path -NewerThan.

  .PARAMETER OlderThan
    Same as Test-Path -OlderThan.

  .PARAMETER Credential
    Same as Test-Path -Credential.

  .PARAMETER Seconds
    Specifies the wait time between attempts in seconds. Same as Start-Sleep -Seconds.

  .PARAMETER Milliseconds
    Specifies the wait time between attempts in milliseconds. Same as Start-Sleep -Milliseconds. The default is 100.

  .PARAMETER Count
    Specifies the maximum number of attempts. The default is 10.

  .EXAMPLE
    Wait-Path -LiteralPath 'C:\temp\file.txt'

    Tests the path up to 10 times with 100 milliseconds between attempts.

  .EXAMPLE
    Wait-Path -Path 'C:\temp\*.txt' -Seconds 1 -Count 5

    Tests the path up to 5 times with 1 second between attempts.

  .OUTPUTS
    System.Boolean
  #>
  [CmdletBinding(DefaultParameterSetName = 'PathMilliseconds')]
  [OutputType([bool])]
  param (
    [Parameter(Mandatory, ParameterSetName = 'PathSeconds', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'PathMilliseconds', Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [SupportsWildcards()]
    [string[]]
    $Path,
    [Alias('PSPath', 'LP')]
    [Parameter(Mandatory, ParameterSetName = 'LiteralPathSeconds', ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'LiteralPathMilliseconds', ValueFromPipelineByPropertyName)]
    [string[]]
    $LiteralPath,
    [string]
    $Filter,
    [SupportsWildcards()]
    [string[]]
    $Include,
    [SupportsWildcards()]
    [string[]]
    $Exclude,
    [Alias('Type')]
    [TestPathType]
    $PathType = [TestPathType]::Any,
    [datetime]
    $NewerThan,
    [datetime]
    $OlderThan,
    [pscredential]
    [System.Management.Automation.Credential()]
    $Credential,
    [Parameter(Mandatory, ParameterSetName = 'PathSeconds')]
    [Parameter(Mandatory, ParameterSetName = 'LiteralPathSeconds')]
    [ValidateRange(0, 2147483)]
    [int]
    $Seconds,
    [Parameter(ParameterSetName = 'PathMilliseconds')]
    [Parameter(ParameterSetName = 'LiteralPathMilliseconds')]
    [ValidateRange(0, [int]::MaxValue)]
    [int]
    $Milliseconds = 100,
    [ValidateRange(1, [int]::MaxValue)]
    [int]
    $Count = 10
  )
  process {
    $testPathParameters = @{}
    foreach ($key in $PSBoundParameters.Keys) {
      if ($key -notin @('Seconds', 'Milliseconds', 'Count')) {
        $testPathParameters[$key] = $PSBoundParameters[$key]
      }
    }
    $sleepParameters = if ($PSBoundParameters.ContainsKey('Seconds')) {
      @{ Seconds = $Seconds }
    } else {
      @{ Milliseconds = $Milliseconds }
    }
    $activity = 'Waiting for path'
    $target = if ($PSCmdlet.ParameterSetName -like 'LiteralPath*') {
      $LiteralPath -join ', '
    } else {
      $Path -join ', '
    }
    try {
      foreach ($attempt in 1..$Count) {
        Write-Progress -Activity $activity -Status $target -PercentComplete ([int](($attempt - 1) * 100 / $Count))
        if (Test-Path @testPathParameters) {
          return $true
        }
        if ($attempt -lt $Count) {
          Start-Sleep @sleepParameters
        }
      }
      return $false
    } finally {
      Write-Progress -Activity $activity -Completed
    }
  }
}
