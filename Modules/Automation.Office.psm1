Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -Filter '*.cs' |
ForEach-Object { Add-Type -LiteralPath $_.FullName }
Get-ChildItem -LiteralPath ($PSScriptRoot | Join-Path -ChildPath 'Sources') -Recurse -Filter '*.ps1' |
Where-Object { $_.BaseName -notlike '*.Tests' } |
ForEach-Object { . $_.FullName }

# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/register-argumentcompleter?view=powershell-5.1
Register-ArgumentCompleter -CommandName @(
  'Get-ExcelDocumentProperty'
  'Get-ExcelDocumentPropertyValue'
  'Set-ExcelDocumentProperty'
  'Get-WordDocumentProperty'
  'Get-WordDocumentPropertyValue'
  'Set-WordDocumentProperty'
  'Get-PowerPointDocumentProperty'
  'Get-PowerPointDocumentPropertyValue'
  'Set-PowerPointDocumentProperty'
) -ParameterName 'Name' -ScriptBlock {
  param (
    [string]
    $CommandName,
    [string]
    $ParameterName,
    [string]
    $WordToComplete,
    [System.Management.Automation.Language.CommandAst]
    $CommandAst,
    [System.Collections.ICollection]
    $FakeBoundParameters
  )
  $null = @($CommandName, $ParameterName, $CommandAst)
  if ($FakeBoundParameters.ContainsKey('Custom') -and $FakeBoundParameters['Custom']) {
    return
  }
  $propertyNames = @(
    'Title'
    'Subject'
    'Author'
    'Keywords'
    'Comments'
    'Template'
    'Last Author'
    'Revision Number'
    'Application Name'
    'Last Print Date'
    'Creation Date'
    'Last Save Time'
    'Total Editing Time'
    'Number of Pages'
    'Number of Words'
    'Number of Characters'
    'Security'
    'Category'
    'Format'
    'Manager'
    'Company'
    'Number of Bytes'
    'Number of Lines'
    'Number of Paragraphs'
    'Number of Slides'
    'Number of Notes'
    'Number of Hidden Slides'
    'Number of Multimedia Clips'
    'Hyperlink Base'
    'Number of Characters (with spaces)'
    'Content Type'
    'Content Status'
    'Language'
    'Document Version'
  )
  $prefix = ([string]$WordToComplete).TrimStart([char[]]@('"', "'"))
  $propertyNames |
  Where-Object { $_.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) } |
  ForEach-Object {
    $completionText = if ($_ -match '\s') { "'$_'" } else { $_ }
    [System.Management.Automation.CompletionResult]::new($completionText, $_, 'ParameterValue', $_)
  }
}
