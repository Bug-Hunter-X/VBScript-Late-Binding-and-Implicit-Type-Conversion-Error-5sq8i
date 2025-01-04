Option Explicit

Dim num1 As Integer
Dim num2 As Integer
Dim sum As Integer

'Explicitly declare variables and assign values
num1 = 10
num2 = "20"

'Error Handling
On Error Resume Next

'Type checking before operation
if IsNumeric(num2) then
 num2 = CInt(num2) 'convert to integer if numeric
 sum = num1 + num2
 WScript.Echo "Sum: " & sum
else
 WScript.Echo "Error: num2 is not a number" 
end if

On Error GoTo 0