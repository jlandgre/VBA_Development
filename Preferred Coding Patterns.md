ExcelSteps PivotTable class contains preferred coding patterns

1. Initialize optional-object attributes in `.Init` so downstream code can assume objects exist.
```vb
Set .dictParams = New Dictionary
If Not dictParams Is Nothing Then Set .dictParams = dictParams
```

2. Validate and normalize early, then keep operational methods simple.
- Run validation methods before configuration/build methods.
- Normalize mixed inputs (array or string) into one internal format.
- Fail early on invalid spec, unknown names, or overlap.
```vb
If Not InitPivotTable(...) Then GoTo ErrorExit
If Not CreatePivotCacheAndTable(...) Then GoTo ErrorExit
If Not ValidateFieldSpecs(...) Then GoTo ErrorExit
If Not ValidateAnalytes(...) Then GoTo ErrorExit
If Not ConfigurePivotFields(...) Then GoTo ErrorExit
```

3. Prefer one canonical internal format for repeated downstream use.
- Store normalized row/column field specs as CSV class attributes.
- Reuse one helper (for example `ValidateFieldNamesExist`) from multiple validators.

4. Use `For Each` for collection/array iteration when index math is not needed.
```vb
i = 1
For Each fieldName In Split(CStr(fieldsCsv), ",")
	With pvt.pvtTable.PivotFields(fieldName)
		.Orientation = fieldOrientation
		.Position = i
	End With
	i = i + 1
Next fieldName
```

5. For fixed-shape arrays, validate shape once and then index directly.
```vb
If errs.IsFail(LBound(analyteDef) <> 0, 3) Then GoTo ErrorExit
If errs.IsFail(UBound(analyteDef) <> 1, 4) Then GoTo ErrorExit
fieldName = CStr(analyteDef(0))
xFunc = CLng(analyteDef(1))
```

6. Keep error-handling scope explicit when using `On Error Resume Next`.
- Use it only around the risky line(s).
- Immediately test `Err.Number` with `errs.IsFail`.
- Always restore error mode after success (`On Error GoTo 0`, then `On Error GoTo ErrorExit` if handled).

7. Remove stale flexibility and dead code after refactors.
- If inputs are now constrained by validator(s), delete conditional branches for old input shapes.
- Remove unused variables and obsolete helper methods.

8. Keep method order consistent with execution order.
- Place validation methods near the top and before configuration/build helpers to improve readability for humans and AI.

9. Use distinct `iCodeLocal` values per failure reason within each function.
- Makes tests and diagnostics deterministic (`errs.Locn`, `errs.iCodeLocal`).

10. Prefer fewer internal variables; use locals only when they improve clarity.
- Avoid introducing temporary variables when a value can be used directly and read clearly.
- Add locals when they materially improve readability (for example very long names, repeated expressions, or complex range arguments).
- Respect `ByRef`/`ByVal` intent in signatures; direct use of class attributes inside methods is preferred unless a local variable is clearly better.
