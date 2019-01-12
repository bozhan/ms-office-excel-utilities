Attribute VB_Name = "Init"
'---------------------------------------------------------------------------------------
' Procedure : initWorkbook
' Date      : 09.04.2014
' Descr.    : Initialize Status, Progress and UI elements properties
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub initWorkbook()
On Error GoTo initWorkbook_Error

'check if macros are enabled
'set tool name
'set table field definitions source and dest sheet names

initWorkbook_Exit:
On Error Resume Next
 
Exit Sub

initWorkbook_Error:
  MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure initWorkbook of module modInit" & vbLf & _
    INFO_ERR_MSG
Resume initWorkbook_Exit
End Sub


