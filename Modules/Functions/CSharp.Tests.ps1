using namespace System.IO

[CmdletBinding()]
param ()
Set-StrictMode -Version Latest

Describe 'Add-Type' {
  It 'compiles DialogSuppressor' {
    Add-Type -LiteralPath ($PSScriptRoot | Join-Path -ChildPath 'DialogSuppressor.cs')

    'NUIDialogSuppressor' -as [type] | Should-NotBeNull
    'VBProjectDialogSuppressor' -as [type] | Should-NotBeNull
  }
  It 'compiles PathCompatibility' {
    Add-Type -LiteralPath ($PSScriptRoot | Join-Path -ChildPath 'PathCompatibility.cs')

    'PathCompatibility' -as [type] | Should-NotBeNull
  }
}
