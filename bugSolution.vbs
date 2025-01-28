Improved error handling, explicit type checking, and proper object management are crucial for robust VBScript. Here's how to address the issues:

Late Binding:
```vbscript
On Error GoTo ErrorHandler
Dim objExcel
Set objExcel = CreateObject("Excel.Application")
' ...use objExcel...
Set objExcel = Nothing
Exit Sub
ErrorHandler:
MsgBox "Error: " & Err.Description
Exit Sub
```
Explicit type checking (where applicable):
```vbscript
Dim x As Integer, y As Integer
x = 10
y = 5
Dim z As Integer
z = x + y
```
Avoid implicit type coercion:
```vbscript
Dim a As Integer, b As Integer
a = CInt("10")
b = CInt("20")
Dim c As Integer
c = a + b
```
Proper exception handling:
```vbscript
On Error GoTo ErrorHandler
' ...Code that might cause an error...
Exit Sub
ErrorHandler:
' Handle the error appropriately...
MsgBox "Error occurred: " & Err.Description
Err.Clear
Resume Next 'or Exit Sub
```
Properly release COM objects:
```vbscript
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
' ...Use objFSO...
Set objFSO = Nothing
```
Always use `On Error GoTo` for better error handling than `On Error Resume Next`.