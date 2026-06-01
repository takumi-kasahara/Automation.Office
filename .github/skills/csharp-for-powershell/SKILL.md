---
name: csharp-for-powershell
description: 'This skill creates C# files that can be loaded by PowerShell 5.1 Add-Type for this repository.'
argument-hint: 'Enter the desired functionality and type overview for the C# file.'
user-invocable: true
---

# C# for PowerShell

- Generate a C# source file that can be loaded by PowerShell 5.1 `Add-Type`.
- The goal is for `Add-Type -Path '<file>.cs'` to succeed in compilation and loading.

## Guidelines

1. Constraint priority:
   - First: .NET Framework API compatibility.
   - Second: PowerShell 5.1 `Add-Type` compilation compatibility.
   - Third: minimal namespace usage.
2. If the request depends on unsupported APIs or features, explain the limitation and suggest a compatible alternative.
3. Generate C# source containing `public class` or `public static class`.
4. Include only namespaces that are actually used by the generated code.
   - Start with `using System;` when needed.
   - Common examples: `System`, `System.IO`, `System.Collections.Generic`.
