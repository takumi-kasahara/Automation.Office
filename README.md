# Automation.Office

A PowerShell module for automating Microsoft Office applications (Access / Excel / PowerPoint / Word / OneNote).

## Requirements

- Windows 11
- PowerShell 5
  - [Pester v6](https://pester.dev/docs/introduction/installation)
  - [PlatyPS](https://learn.microsoft.com/en-us/powershell/utility-modules/platyps/overview?view=ps-modules)
  - [PSScriptAnalyzer](https://learn.microsoft.com/en-us/powershell/utility-modules/psscriptanalyzer/overview?view=ps-modules)
- Microsoft 365 Apps
  - [Rubberduck v2.x](https://rubberduckvba.ca/)

## Installation

Run from the repository root:

```powershell
.\Install.cmd
```

If you want to install Pester v6, run:

```powershell
Install-PSResource -Name Pester
```

If you want to install PlatyPS, run:

```powershell
Install-PSResource -Name Microsoft.PowerShell.PlatyPS
```

If you want to install PSScriptAnalyzer, run:

```powershell
Install-PSResource -Name PSScriptAnalyzer
```
