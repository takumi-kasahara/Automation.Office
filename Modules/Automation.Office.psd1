@{
  RootModule           = 'Automation.Office.psm1'
  ModuleVersion        = '1.0.0'
  CompatiblePSEditions = @('Desktop')
  FunctionsToExport    = @(
    #region Access
    'Get-AccessVersionInfo'
    'New-AccessFile'
    'Open-AccessFile'
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
    'Get-ExcelVersionInfo'
    'New-ExcelFile'
    'Open-ExcelFile'
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
    'Get-WordVersionInfo'
    'New-WordFile'
    'Open-WordFile'
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
    'Get-PowerPointVersionInfo'
    'New-PowerPointFile'
    'Open-PowerPointFile'
    'Get-PowerPointFileProperty'
    'Set-PowerPointFileProperty'
    'Get-PowerPointDocumentProperty'
    'Get-PowerPointPropertyValue'
    'Set-PowerPointDocumentProperty'
    'Remove-PowerPointDocumentProperty'
    'Export-PowerPointVBProject'
    'Import-PowerPointVBProject'
    'Get-PowerPointSpeakerNote'
    'Export-PowerPointSpeakerNote'
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
