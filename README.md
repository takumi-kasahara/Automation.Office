# Automation.Office

A PowerShell module for automating Microsoft Office applications (Access / Excel / PowerPoint / Word / OneNote).

## Requirements

- Windows 11
- PowerShell 5
  - [Pester v5](https://pester.dev/docs/introduction/installation)
  - [PSScriptAnalyzer](https://learn.microsoft.com/en-us/powershell/utility-modules/psscriptanalyzer/overview?view=ps-modules)
- Microsoft 365 Apps
  - [Rubberduck v2.x](https://rubberduckvba.ca/)

## Installation

Run from the repository root:

``` powershell
./Install.bat
```

If you want to run tests:

``` powershell
Invoke-Pester -Path .\Modules\Tests
```
