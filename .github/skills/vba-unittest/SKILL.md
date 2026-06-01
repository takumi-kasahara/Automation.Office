---
name: vba-unittest
description: 'Create or update Rubberduck VBA unit tests, including test module setup, lifecycle hooks, and VBA Moq-based doubles.'
argument-hint: 'Provide: (1) target VBA component path, (2) behavior to test, and (3) whether to add new tests or update existing tests.'
user-invocable: true
---

# Rubberduck VBA Unit Testing

## Guidelines

Use these steps in order:

1. Confirm test scope: new tests, extending existing tests, or refactoring unstable tests.
2. Identify target component under `Apps/*/Components/` and list behaviors to verify.
3. Locate or create a dedicated test module for that component, and use `./assets/TestModule1.bas` as the default template when creating a new module.
4. Add or update Rubberduck test annotations and test lifecycle procedures.
5. Implement test cases with Arrange, Act, Assert structure.
6. Add VBA Moq-based test doubles when dependencies or side effects make direct testing difficult.
7. Run Rubberduck Parse and Unit Tests, then fix failures before finishing.

- Keep one test module focused on one production component or one coherent feature area.
- For newly created test modules, start from `./assets/TestModule1.bas` and adapt names/categories/assertions to the target behavior.
- Name test procedures so the expected behavior is obvious from the procedure name.
- Prefer deterministic tests; avoid filesystem, clock, UI, or COM state dependencies unless explicitly mocked/faked.
- When tests expose design issues, improve production code for testability (dependency seams, narrower responsibilities) before adding complex test setup.
- Do not mute failing tests; either fix them or document why they are intentionally pending.

## Test Module Checklist

- Module has Rubberduck test markers (for example, `@TestModule` and a test folder annotation).
- Optional lifecycle procedures are added only when needed (`@ModuleInitialize`, `@ModuleCleanup`, `@TestInitialize`, `@TestCleanup`).
- Each test has a single behavioral expectation.
- Assertions verify behavior, not implementation details.
- Error-path tests are explicit (for example, expected runtime error behavior).

## What You Can Test

Use this guide to choose test targets and test style.

### Pure Logic Tests

- Purpose: verify calculations, string/date transforms, and branching rules.
- Input style: direct literals and edge values.
- Typical assertions: return value, output argument, or state after call.
- Notes: no mocks needed when behavior has no external dependency.

### Stateful Behavior Tests

- Purpose: verify state transitions in class modules (for example, initialize, mutate, reset).
- Input style: call sequence that mirrors realistic usage.
- Typical assertions: property values and invariants before/after each step.
- Notes: reset mutable state in `@TestCleanup` to avoid cross-test leakage.

### Error Contract Tests

- Purpose: verify expected failures for invalid inputs and unsupported operations.
- Input style: invalid value, missing prerequisite, or illegal call order.
- Typical assertions: raised error number and stable message contract.
- Notes: keep one expected failure reason per test to simplify diagnosis.

### Interaction Tests with VBA Moq

- Purpose: verify collaboration behavior when code calls workbook/worksheet APIs, dialogs, filesystem wrappers, or COM collaborators.
- Input style: arrange doubles for collaborator interfaces/classes and configure expected calls.
- Typical assertions: interaction count/order, passed arguments, and fallback behavior.
- Notes: verify interactions only when interactions are part of the requirement.

### Side-Effect Boundary Tests

- Purpose: verify commands that produce side effects (file write, worksheet edit, export/import) without touching real external resources.
- Input style: wrap side effects behind seams and substitute doubles.
- Typical assertions: requested action parameters and resulting observable state.
- Notes: do not depend on machine-local paths, locale, or current time unless injected.

### Regression and Flaky-Test Prevention

- Purpose: lock in fixes for previously reported defects and unstable behavior.
- Input style: minimal reproducer inputs from the original bug scenario.
- Typical assertions: bug no longer reproduces and related behavior stays unchanged.
- Notes: remove hidden shared state and time/order coupling when stabilizing tests.

## Test Type Matrix

| Test type            | Best for                         | Double needed | Primary assertion          |
| -------------------- | -------------------------------- | ------------- | -------------------------- |
| Pure logic           | deterministic business rules     | No            | return value/state         |
| Stateful behavior    | class lifecycle and invariants   | Sometimes     | state transition           |
| Error contract       | invalid input and guard clauses  | No            | error number/message       |
| Interaction (Moq)    | collaborator calls and protocol  | Yes           | call/argument verification |
| Side-effect boundary | IO and COM boundaries            | Yes           | requested side effect      |
| Regression           | bug reproduction and fix lock-in | Depends       | non-recurrence             |

## Branching Logic

Use this decision flow when writing tests:

1. If behavior is pure and has no external dependency:
   - Write direct AAA tests without mocks.
2. If behavior depends on external state (file system, worksheet/workbook state, dialogs, COM objects):
   - Introduce a seam and use VBA Moq test doubles.
3. If behavior can raise expected errors:
   - Add explicit failure-path tests and verify error number/message contract.
4. If existing tests are flaky:
   - Remove shared mutable state and isolate setup/teardown per test.

## VBA Moq Guidance

- Use mocks/stubs/spies to isolate collaborators and side effects.
- Mock only the collaboration points required for the behavior under test.
- Verify interactions only when interaction itself is part of the requirement.
- Keep mock setup small; if setup is large, split production behavior into smaller units.
- Prefer readable test data builders/helpers over repeated inline setup.

## Completion Checks

1. Rubberduck parsing succeeds with no syntax errors.
2. All related unit tests pass locally with stable results across reruns.
3. No new critical Rubberduck inspection findings are introduced.
4. Tests clearly express expected behavior and are maintainable by another contributor.

## References

- [Unit Testing](https://github.com/rubberduck-vba/Rubberduck/wiki/Unit-Testing)
- [VBA Moq Mocking Framework](https://github.com/rubberduck-vba/Rubberduck/wiki/VBA-Moq-Mocking-Framework)
