@{
  RootModule           = 'Automation.Office.psm1'
  ModuleVersion        = '1.0.0'
  CompatiblePSEditions = @('Desktop')
  FunctionsToExport    = @(
    #region Access
    'New-AccessFile'
    'Open-AccessFile'
    'Get-AccessAppProperty'
    'Set-AccessAppProperty'
    'Get-AccessFileProperty'
    'Set-AccessFileProperty'
    'Get-AccessTable'
    'Get-AccessTableColumn'
    'Get-AccessView'
    'Get-AccessViewColumn'
    'Invoke-AccessQuery'
    'Export-AccessDatabase'
    'Import-AccessDatabase'
    'Export-AccessVBProject'
    'Import-AccessVBProject'
    #endregion
    #region Excel
    'New-ExcelFile'
    'Open-ExcelFile'
    'Get-ExcelAppProperty'
    'Set-ExcelAppProperty'
    'Get-ExcelFileProperty'
    'Set-ExcelFileProperty'
    'Get-ExcelTable'
    'Get-ExcelTableColumn'
    'Invoke-ExcelQuery'
    'Get-ExcelDocumentProperty'
    'Get-ExcelPropertyValue'
    'Set-ExcelDocumentProperty'
    'Remove-ExcelDocumentProperty'
    'Export-ExcelVBProject'
    'Import-ExcelVBProject'
    #endregion
    #region Word
    'New-WordFile'
    'Open-WordFile'
    'Get-WordAppProperty'
    'Set-WordAppProperty'
    'Get-WordFileProperty'
    'Set-WordFileProperty'
    'Get-WordDocumentProperty'
    'Get-WordPropertyValue'
    'Set-WordDocumentProperty'
    'Remove-WordDocumentProperty'
    'Export-WordVBProject'
    'Import-WordVBProject'
    #endregion
    #region PowerPoint
    'New-PowerPointFile'
    'Open-PowerPointFile'
    'Get-PowerPointAppProperty'
    'Set-PowerPointAppProperty'
    'Get-PowerPointFileProperty'
    'Set-PowerPointFileProperty'
    'Get-PowerPointSpeakerNote'
    'Get-PowerPointDocumentProperty'
    'Get-PowerPointPropertyValue'
    'Set-PowerPointDocumentProperty'
    'Remove-PowerPointDocumentProperty'
    'Export-PowerPointVBProject'
    'Import-PowerPointVBProject'
    #endregion
    #region OneNote
    'Get-OneNoteHierarchy'
    'New-OneNoteNotebook'
    'New-OneNoteSection'
    'New-OneNotePage'
    'Export-OneNoteHierarchy'
    'Export-OneNotePageContent'
    'Export-OneNotePageAsDocument'
    'Export-OneNoteBinaryObject'
    'Import-OneNoteHierarchy'
    'Import-OneNotePageContent'
    #endregion
  )
  CmdletsToExport      = @()
  VariablesToExport    = @()
  AliasesToExport      = @()
}
