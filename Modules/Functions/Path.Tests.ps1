using module .\Path.psm1

[CmdletBinding()]
[System.Diagnostics.CodeAnalysis.SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Test scripts often use variables for setup and verification that may not be assigned in a way that satisfies this rule')]
param ()
Set-StrictMode -Version Latest

InModuleScope -ModuleName 'Path' {
  Describe 'Wait-Path' {
    BeforeAll {
      Mock -CommandName Start-Sleep -MockWith {}
    }
    Context 'ParameterSetName' {
      It 'returns true by Path when the path exists' {
        Mock -CommandName Test-Path -MockWith { $true } -ParameterFilter { $Path -eq 'C:\*.txt' }

        Wait-Path -Path 'C:\*.txt' | Should-BeTrue
        Should-Invoke -CommandName Test-Path -Times 1 -Exactly
      }
      It 'returns true by LiteralPath when the path exists' {
        Mock -CommandName Test-Path -MockWith { $true } -ParameterFilter { $LiteralPath -eq 'C:\[a].txt' }

        Wait-Path -LiteralPath 'C:\[a].txt' | Should-BeTrue
        Should-Invoke -CommandName Test-Path -Times 1 -Exactly
      }
      It 'sleeps by Seconds' {
        Mock -CommandName Test-Path -MockWith { $false }

        Wait-Path -LiteralPath 'C:\a.txt' -Seconds 2 -Count 2 | Should-BeFalse
        Should-Invoke -CommandName Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 2 }
      }
      It 'sleeps by Milliseconds' {
        Mock -CommandName Test-Path -MockWith { $false }

        Wait-Path -Path 'C:\a.txt' -Milliseconds 50 -Count 2 | Should-BeFalse
        Should-Invoke -CommandName Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 50 }
      }
    }
    Context 'Output' {
      It 'returns true when the path appears on a later attempt' {
        $script:attempts = 0
        Mock -CommandName Test-Path -MockWith { $script:attempts++; $script:attempts -ge 3 }

        $result = Wait-Path -LiteralPath 'C:\a.txt'
        $result | Should-BeTrue
        $result | Should-HaveType ([bool])
        Should-Invoke -CommandName Test-Path -Times 3 -Exactly
        Should-Invoke -CommandName Start-Sleep -Times 2 -Exactly
      }
      It 'returns false when the path never appears' {
        Mock -CommandName Test-Path -MockWith { $false }

        Wait-Path -LiteralPath 'C:\a.txt' | Should-BeFalse
      }
    }
    Context 'Other parameters' {
      It 'tries 10 times and sleeps 100 milliseconds by default' {
        Mock -CommandName Test-Path -MockWith { $false }

        Wait-Path -LiteralPath 'C:\a.txt' | Should-BeFalse
        Should-Invoke -CommandName Test-Path -Times 10 -Exactly
        Should-Invoke -CommandName Start-Sleep -Times 9 -Exactly -ParameterFilter { $Milliseconds -eq 100 }
      }
      It 'tries the number of times specified by Count' {
        Mock -CommandName Test-Path -MockWith { $false }

        Wait-Path -LiteralPath 'C:\a.txt' -Count 3 | Should-BeFalse
        Should-Invoke -CommandName Test-Path -Times 3 -Exactly
        Should-Invoke -CommandName Start-Sleep -Times 2 -Exactly
      }
      It 'passes Test-Path parameters through' {
        Mock -CommandName Test-Path -MockWith { $true } -ParameterFilter {
          $PathType -eq 'Leaf' -and $Filter -eq '*.txt' -and $Include -eq 'a*' -and $Exclude -eq 'b*'
        }

        Wait-Path -Path 'C:\*' -PathType Leaf -Filter '*.txt' -Include 'a*' -Exclude 'b*' | Should-BeTrue
      }
      It 'does not pass Count or sleep parameters to Test-Path' {
        Mock -CommandName Test-Path -MockWith { $true }

        Wait-Path -LiteralPath 'C:\a.txt' -Seconds 1 -Count 2 | Should-BeTrue
        Should-Invoke -CommandName Test-Path -Times 1 -Exactly -ParameterFilter {
          -not $PSBoundParameters.ContainsKey('Count') -and -not $PSBoundParameters.ContainsKey('Seconds')
        }
      }
    }
    Context 'Edge cases' {
      It 'does not sleep when Count is 1' {
        Mock -CommandName Test-Path -MockWith { $false }

        Wait-Path -LiteralPath 'C:\a.txt' -Count 1 | Should-BeFalse
        Should-NotInvoke -CommandName Start-Sleep
      }
      It 'uses a real path that already exists' {
        Wait-Path -LiteralPath $env:TEMP -PathType Container | Should-BeTrue
      }
    }
  }
}
