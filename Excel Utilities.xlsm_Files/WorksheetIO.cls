VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : WorksheetIO
' Author    : Ivanov, Bozhan
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Private xlApp As Excel.Application
Private xlWb As Excel.Workbook
Private xlWs As Excel.Worksheet
Private headColumnCol As New HeadColumnDict
Private xlRng As Range
Private hiddenCols As Variant, hiddenRows As Variant
Private hiddenComments As Collection
Private condHeadColumns As New HeadColumnDict

Property Get Application() As Excel.Application
  If xlApp Is Nothing _
  And Not Me.Workbook Is Nothing Then
    Set xlApp = Me.Workbook.Application
  End If
  
  Set Application = xlApp
End Property

Property Let Application(ByRef app As Excel.Application)
  Set xlApp = app
End Property

Property Get Workbook() As Excel.Workbook
  If xlWb Is Nothing _
  And Not Me.Worksheet Is Nothing Then
    Set xlWb = Me.Worksheet.Parent
  End If
  
  Set Workbook = xlWb
End Property

Property Let Workbook(ByRef wb As Excel.Workbook)
  Set xlWb = wb
End Property

Property Get Worksheet() As Excel.Worksheet
  If xlWs Is Nothing _
  And Not Me.UsedRange Is Nothing Then
    Set xlWs = xlRng.Worksheet
  End If
  
  Set Worksheet = xlWs
End Property

Property Let Worksheet(ByRef xlWorksheet As Excel.Worksheet)
  Set xlWs = xlWorksheet
End Property

Property Get UsedRange() As Excel.Range
  If xlRng Is Nothing _
  And Not xlWs Is Nothing Then
    Set xlRng = xlWs.UsedRange
  End If
  
  Set UsedRange = xlRng
End Property

Property Let UsedRange(ByRef rng As Excel.Range)
  Set xlRng = rng
End Property

Property Get hcc() As HeadColumnDict
  Set hcc = headColumnCol
End Property

Property Let hcc(ByRef HeadColumnColObject As HeadColumnDict)
  Set headColumnCol = HeadColumnColObject
End Property

Property Get ConditionHCC() As HeadColumnDict
  Set ConditionHCC = condHeadColumns
End Property

Property Let ConditionHCC(ByRef HeadColumnColObject As HeadColumnDict)
  Set condHeadColumns = HeadColumnColObject
End Property

Public Sub RevealHiddenData()
  hiddenCols = Util.Excel.GetHiddenColumnsIndeces(Me.Worksheet, Me.UsedRange)
  hiddenRows = Util.Excel.GetHiddenRowsIndeces(Me.Worksheet, Me.UsedRange)
  Set hiddenComments = Util.Excel.GetHiddenComments(Me.Worksheet)
  Util.Excel.ShowColumnsWithIndeces Me.Worksheet, hiddenCols
  Util.Excel.ShowRowsWithIndeces Me.Worksheet, hiddenRows
End Sub

Public Sub HideRevealedData()
  Util.Excel.HideColumnsWithIndeces Me.Worksheet, hiddenCols
  Util.Excel.HideRowsWithIndeces Me.Worksheet, hiddenRows
  Util.Excel.HideComments hiddenComments
End Sub

Public Sub TurnOffScreenUpdatingAndAlerts()
  Util.Excel.TurnOffScreenUpdatingAndAlerts Me.Application
End Sub

Public Sub TurnOnScreenUpdatingAndAlerts()
  Util.Excel.TurnOnScreenUpdatingAndAlerts Me.Application
End Sub

Public Sub OpenWorkbookFromPath( _
  ByVal path As String, _
  Optional ByVal Password As Variant, _
  Optional ByVal writeResPassword As Variant _
)
  Set xlWb = Util.Excel.GetWorkbookFromPath( _
    Me.Application, _
    path, _
    Password, _
    writeResPassword)
End Sub

Public Sub CloseAllWorbooks(Optional Save As Boolean = False)
  Util.Excel.CloseWorbooksInApplication Me.Application, Save
End Sub

Public Sub CloseWorbook(Optional Save As Boolean = False)
  If Not Me.Workbook Is Nothing Then
    If Not Util.Excel.IsWorkbookOpen(Me.Application, Me.Workbook.name) Then
      Me.Workbook.Close Save
    End If
  End If
End Sub

