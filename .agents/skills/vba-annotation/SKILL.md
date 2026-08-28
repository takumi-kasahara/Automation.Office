---
name: vba-annotation
description: "Use Rubberduck and VB_Attribute Annotations to implement and maintain VBA modules with consistent documentation and inspection quality."
argument-hint: "Provide: (1) target VBA component path, and (2) what to add or change (module/procedures/annotations)."
user-invocable: true
---

# VBA Annotations for Rubberduck

## Guidelines

Use these steps in order:

1. Confirm target scope: module-level updates, procedure-level updates, or both.
2. Open the target component under `Apps/*/Components/` and identify module type (`.bas`, `.cls`, `.frm`, `.vba`).
3. Add or update Rubberduck VB_Attribute Annotations as structured documentation comments.
4. Ensure annotations stay close to their target (module header or directly above the procedure/member).
5. Run Rubberduck Parse and Code Inspections, and resolve annotation-related findings.
6. Validate repository workflow by running compile/test commands relevant to the changed components.

- Prefer VB_Attribute Annotations over ad-hoc comments when documenting intent, usage, and behavior.
- Keep annotation syntax exactly as defined by Rubberduck docs.
- When multiple annotations appear in the same scope, keep their appearance order in dictionary order by annotation name.
- When adding or renaming procedures, update related annotations in the same change.
- For class modules, ensure class-level metadata and member-level annotations are both reviewed.
- Do not add annotations that contradict runtime behavior; update implementation first, then documentation annotations.

## Annotation Checklist

- Module-level annotation exists and reflects the module responsibility.
- Public procedures/functions include intent-focused annotation text.
- Complex logic blocks include concise rationale in annotation comments.
- Any suppression/ignore annotation is justified and minimal.
- Annotation wording matches parameter names and return behavior.

## Annotation Placement Guide

Use this guide to decide where each annotation can be placed and what hidden attribute it controls.

### Scope and Limitations

- Supported scope: standard modules (`.bas`), class modules (`.cls`), and document modules (`.vba`).
- Repository rule: do not omit document modules. Apply annotations to document modules when they are in scope for the requested change.
- Class-only annotations remain class-only (`@PredeclaredId`, `@Exposed`, `@DefaultMember`, `@Enumerator`).
- Place annotations immediately above the target declaration.
- After editing annotations, run Rubberduck Parse and Code Inspections, then synchronize attributes.

### Module-Level Annotations

- `@ModuleDescription("...")`: place near module header; controls `VB_Description`; applies to Standard, Class, Document.
- `@PredeclaredId`: place near module header; controls `VB_PredeclaredId = True`; applies to Class.
- `@Exposed`: place near module header; controls `VB_Exposed = True`; applies to Class.

Example:

```vb
'@ModuleDescription("Description of the module's purpose")
```

### Variable-Level Annotation

Use `@VariableDescription("...")` on the line immediately above a module field declaration.

| Annotation                    | Where to place                      | Target attribute    | Applies to    |
| ----------------------------- | ----------------------------------- | ------------------- | ------------- |
| `@VariableDescription("...")` | Immediately above field declaration | `VB_VarDescription` | Module fields |

Example:

```vb
'@VariableDescription("It's a thing")
Private thing As Something
```

### Member-Level Annotations

For procedure/property members, place annotations immediately above the member declaration. For `Property` procedures, the attribute is typically attached to the `Get` member.

- `@Description("...")`: place immediately above procedure/property; controls `MemberName.VB_Description`; applies to Standard, Class, Document.
- `@ExcelHotkey("D")`: place immediately above macro procedure; controls `VB_ProcData.VB_Invoke_Func`; applies to Standard, Document (Excel macros).
- `@DefaultMember`: place immediately above one class member only; controls `VB_UserMemId = 0`; applies to Class.
- `@Enumerator`: place immediately above `NewEnum` member; controls `VB_UserMemId = -4`; applies to Class.

Examples:

```vb
'@Description("Does something")
Public Sub DoSomething()
End Sub
```

```vb
'@ExcelHotkey("D")
Public Sub DoSomething()
End Sub
```

```vb
'@DefaultMember
Public Property Get Item(ByVal index As Long) As Variant
End Property

'@Enumerator
Public Property Get NewEnum() As IUnknown
End Property
```

### Value Rules and Notes

- `@ExcelHotkey` expects a 1-character string.
- Lowercase hotkey value maps to `Ctrl + <key>`.
- Uppercase hotkey value maps to `Ctrl + Shift + <key>`.
- Only one class member can be the `@DefaultMember`.
- For collection classes, `@Enumerator` is typically used on a `NewEnum` member returning `IUnknown`.

## Completion Checks

1. Rubberduck parsing succeeds with no syntax errors.
2. Rubberduck inspections show no new critical findings from this change.
3. Changed components still compile through repository build workflow.
4. Annotation comments and code behavior are consistent after final review.

## References

- [VB_Attribute Annotations](https://github.com/rubberduck-vba/Rubberduck/wiki/VB_Attribute-Annotations)
