# Iterative Insanity: Managing Calculation Order in Excel Circular References

## Problem

Record the date and time when a cell changes. Requirements:

- No VBA, no macros
- Formulas only
- Date must not update when the file is reopened

## Solution

Two columns:

**Intermediate (M):**

```excel
=IF(F3="","",IF(OR(M3="",F3<>M3),F3,M3))
```

**Target (G):**

```excel
=IF(F3="","",IF(OR(G3="",F3<>M3),NOW(),G3))
```

Excel settings:

- File → Options → Formulas → Enable iterative calculation
- Maximum Iterations: `1`
- Maximum Change: `0.001`

## Findings

Two rules affect correctness, discovered in practice.

### 1. Calculation order on a single sheet

With iterative calculation enabled, Excel evaluates cells in physical order: left to right, top to bottom. Logical dependencies are ignored.

- If M is to the left of G, M is already updated by the time G is evaluated. The condition `F3<>M3` is false, and the date is not recorded.
- **M must be to the right of G.**

### 2. Calculation order across sheets

When M is moved to a separate sheet, the order is determined by the alphabetical order of sheet names.

- For G to be evaluated after M, the sheet containing G must come **before** the sheet containing M in alphabetical order.

## General Principle

Control calculation order through the dependency graph:

- G depends on M
- M does not depend on G

→ Excel evaluates them in the correct order automatically.

## Conclusion

Excel's behavior in iterative calculation is deterministic and documented (order on a sheet, order across sheets). Accounting for these rules allows building predictable constructions with circular references.

## References

- [Stack Overflow: Calculation order when iterative calculation is enabled](https://stackoverflow.com/questions/18084718/calculation-order-when-iterative-calculation-circular-reference-is-enabled#answer-18086484)