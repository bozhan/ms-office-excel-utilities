VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : EventHandler
' Author    : Ivanov, Bozhan
' Purpose   : EVENTS & EVENT HANDLING
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : HandleEvent
' Descr.    : Manages fired events
'---------------------------------------------------------------------------------------
Public Sub HandleEvent(ByVal caller As String)
  Dim exec As Executor
  
  Select Case caller
    Case Setting(ImportSourceButtonName)
      Set exec = New Executor
      exec.ImportSources
      Set exec = Nothing
    Case Setting(ExecuteButtonName)
      Set exec = New Executor
      exec.GenerateDiscrepancyReport
      Set exec = Nothing
    Case Setting(HelpButtonName)
      HelpEvent
  End Select
End Sub

Private Sub HelpEvent()
  MsgBox InfoErrMsg(), _
    vbInformation, _
    "Support contact..."
End Sub
