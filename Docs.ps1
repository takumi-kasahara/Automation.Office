<#
.SYNOPSIS
  Generates Markdown command documentation for the repository PowerShell scripts.

.DESCRIPTION
  Generates or refreshes Markdown help files using PlatyPS (New-MarkdownCommandHelp) for all *.ps1 scripts found in the repository root and in selected skill script folders.
  Then applies Prettier formatting and markdownlint-cli2 auto-fixes to the generated Markdown files.
  The script changes its working directory to the script root before processing.

.NOTES
  Requires the platyPS (New-MarkdownCommandHelp) command to be available.
  Requires Node.js tooling for Prettier and markdownlint-cli2.
  Regenerated files are written in-place and may overwrite existing help files.
#>
[CmdletBinding()]
param ()
# Set-StrictMode -Version Latest
Set-Location -LiteralPath $PSScriptRoot

$docs = '.\Docs'
Import-Module -Name .\Modules\Automation.Office.psd1 -Force
New-MarkdownCommandHelp -ModuleInfo (Get-Module -Name 'Automation.Office') -OutputFolder $docs -Force

Get-ChildItem -Filter '*.ps1' |
ForEach-Object { Get-Command -Name $_.FullName } |
ForEach-Object { New-MarkdownCommandHelp -CommandInfo $_ -OutputFolder $docs -Force }

npx.cmd prettier --write "$docs\**\*.md"
npx.cmd markdownlint-cli2 --fix "$docs\**\*.md"

$scripts = '.\Scripts'
Get-ChildItem -Path $scripts -Filter '*.ps1' -File |
ForEach-Object { Get-Command -Name $_.FullName } |
ForEach-Object { New-MarkdownCommandHelp -CommandInfo $_ -OutputFolder $scripts -Force }

npx.cmd prettier --write "$scripts\**\*.md"
npx.cmd markdownlint-cli2 --fix "$scripts\**\*.md"
