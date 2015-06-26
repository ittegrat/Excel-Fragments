Attribute VB_Name = "HyperFunctions"
Option Explicit

Public Function HyperAddr(wsName) As String
  HyperAddr = _
    "[" & ThisWorkbook.Name & "]" & _
    "'" & wsName & "'" & _
    "!" & "A1"
End Function

' TO DO: rng.Address(External:=TRUE) works ?
Public Function Addr2Text(rng As Range) As String
  Addr2Text = _
    "[" & ThisWorkbook.Name & "]" & _
    "'" & rng.Parent.Name & "'" & _
    "!" & rng.Address
End Function
