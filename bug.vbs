Late Binding: VBScript's late binding can lead to runtime errors if an object or method doesn't exist.  This is especially problematic when dealing with COM objects or external libraries where version inconsistencies can occur.
```vbscript
Dim objExcel
Set objExcel = CreateObject("Excel.Application")
' ...use objExcel...
Set objExcel = Nothing
```
If the Excel application isn't installed, or the version is incompatible, the `CreateObject` call will fail, potentially causing a runtime error.

Type Mismatches:  VBScript is weakly typed, so type mismatches can be subtle and difficult to debug.  Improper data handling (e.g., trying to perform arithmetic on a string) leads to unexpected results.
```vbscript
Dim x, y
x = "10"
y = 5
Dim z
z = x + y  ' Type mismatch error
```

Implicit Type Coercion: VBScript's automatic type conversions can hide errors. The behavior might not be what's expected, leading to logical bugs.
```vbscript
Dim a, b
a = "10"
b = "20"
Dim c
c = a + b 'String concatenation rather than addition
```

Unhandled Exceptions:  VBScript's error handling is basic. Missing error handling can cause your script to crash unexpectedly. Using `On Error Resume Next` is sometimes used, but this can mask serious problems.
```vbscript
' ...some code that might fail...
On Error Resume Next
' ...more code...
' Error is silently ignored
```

Incorrect Object References:  Failure to properly set and release object references (`Set obj = Nothing`) leads to memory leaks and resource exhaustion. This is especially important when working with COM objects.