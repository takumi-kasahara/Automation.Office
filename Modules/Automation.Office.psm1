Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -Filter '*.cs' |
ForEach-Object { Add-Type -LiteralPath $_.FullName }
Get-ChildItem -LiteralPath ($PSScriptRoot | Join-Path -ChildPath 'Sources') -Recurse -Filter '*.psm1' |
ForEach-Object { Import-Module -Name $_.FullName }
