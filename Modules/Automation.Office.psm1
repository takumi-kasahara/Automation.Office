Get-ChildItem -LiteralPath $PSScriptRoot -Recurse -Filter '*.cs' |
ForEach-Object { Add-Type -LiteralPath $_.FullName }
Get-ChildItem -LiteralPath ($PSScriptRoot | Join-Path -ChildPath 'Sources') -Recurse -Filter '*.ps1' |
Where-Object { $_.BaseName -notlike '*.Tests' } |
ForEach-Object { . $_.FullName }
