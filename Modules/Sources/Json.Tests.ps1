using module .\Json.psm1

[CmdletBinding()]
param ()
Set-StrictMode -Version Latest

Describe 'ConvertTo-UnescapedJson' {
  BeforeAll {
    $script:text = -join @(
      [char]0x21
      [char]0x22
      [char]0x23
      [char]0x24
      [char]0x25
      [char]0x26
      [char]0x27
      [char]0x28
      [char]0x29
      [char]0x2A
      [char]0x2B
      [char]0x2C
      [char]0x2D
      [char]0x2E
      [char]0x2F
      [char]0x3A
      [char]0x3B
      [char]0x3C
      [char]0x3D
      [char]0x3E
      [char]0x3F
      [char]0x40
      [char]0x5B
      [char]0x5C
      [char]0x5D
      [char]0x5E
      [char]0x5F
      [char]0x60
      [char]0x7B
      [char]0x7C
      [char]0x7D
      [char]0x7E
    )
  }
  It 'should create JSON string with special characters' {
    $text | ConvertTo-Json | ConvertFrom-Json | Should-BeString $text
    $text | ConvertTo-Json | Should-BeString '"!\"#$%\u0026\u0027()*+,-./:;\u003c=\u003e?@[\\]^_`{|}~"'
  }
  It 'should decode JSON string with special characters' {
    $text | ConvertTo-UnescapedJson | ConvertFrom-Json | Should-BeString $text
    $text | ConvertTo-UnescapedJson | Should-BeString '"!\"#$%&''()*+,-./:;<=>?@[\\]^_`{|}~"'
  }
}
