using namespace System.Globalization

Set-StrictMode -Version Latest

function ConvertTo-UnescapedJson {
  <#
  .SYNOPSIS
    Converts an object to JSON without unnecessary Unicode escapes.

  .DESCRIPTION
    Uses the Windows PowerShell 5.1 ConvertTo-Json parameters and pipeline behavior.
    Decodes Unicode escapes while preserving escaped backslashes, quotation marks,
    and control characters so that the result remains valid JSON.

  .PARAMETER InputObject
    The object to convert. Accepts pipeline input and null.

  .PARAMETER Depth
    The serialization depth passed to ConvertTo-Json.
    The default is `2`.

  .PARAMETER Compress
    Omits formatting whitespace from the JSON.

  .EXAMPLE
    '&<>' | ConvertTo-UnescapedJson -Compress

    Returns the JSON string '"&<>"'.

  .OUTPUTS
    System.String

  .NOTES
    Only Unicode escapes are decoded; other JSON escape sequences are preserved.
  #>
  [CmdletBinding(RemotingCapability = 'None')]
  [OutputType([string])]
  param (
    [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
    [AllowNull()]
    [object]
    $InputObject,
    [ValidateRange(1, 100)]
    [int]
    $Depth = 2,
    [switch]
    $Compress
  )
  process {
    ConvertTo-Json @PSBoundParameters |
    ForEach-Object {
      [Regex]::Replace(
        $_,
        '\\(?:u(?<HHHH>[a-fA-F0-9]{4})|.)',
        {
          param ($match)
          if ($match.Groups['HHHH'].Success) {
            $value = [int]::Parse($match.Groups['HHHH'].Value, [NumberStyles]::HexNumber)
            if ($value -in 0x26, 0x27, 0x3C, 0x3E) {
              return ([char]$value).ToString()
            }
          }
          return $match.Value
        }
      )
    }
  }
}
