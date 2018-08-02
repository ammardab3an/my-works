Attribute VB_Name = "Module1"

Sub Macro4()


Dim a As Integer
Dim b As Integer
a = 41
b = 41

Dim aa As String
Dim bb As String
Dim cc As String
Dim dd As String

For i = 1 To 30

a = a + 9
b = b + 1
aa = "C" & a
bb = "D" & a
cc = "I" & b
dd = "J" & b
    
    Range(aa, bb).Select
    Selection.Copy
    
    Range(cc, dd).Select
    ActiveSheet.Paste
Next


End Sub

