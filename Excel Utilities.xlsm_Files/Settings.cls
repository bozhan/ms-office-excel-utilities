VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Settings
' Author    : Bozhan Ivanov -
' Purpose   : Automatically generated class for Tool Template Excel by Bozhan Ivanov
' Comment   : Default error handling should be avoided here because it references the
'   settings class.
'---------------------------------------------------------------------------------------
Option Explicit
Private Const SETTINGS_LIST_NAME = "GlobalSettings"
Private Const COLUMN_NAME_ID = "ID"
Private Const COLUMN_NAME_VALUE = "Value"
Private Const COLUMN_NAME_DESCRIPTION = "Description"


Public Enum SettingRowId
  ToolSupportContactName = 1
  ToolSupportContactEmail = 2
  dateFormat = 3
  DefaultRowHeight = 4
  FiscalYearStartMonth = 5
  ReportSettingsSheetName = 6
  InputToolSheetName = 7
  ImportSourceButtonName = 8
  ExecuteButtonName = 9
  HelpButtonName = 10
  ReportFieldMappingTableName = 11
  ReportFieldConditionsTableName = 12
  ReportFieldTypesTableName = 13
  SourcesTableName = 14
End Enum


Private wb As Workbook
Private lo As ListObject
Private WS As Worksheet


Private Sub Class_Initialize()
  checkNamesExistance
End Sub


Private Sub Class_Terminate()
  Set wb = Nothing
  Set lo = Nothing
  Set WS = Nothing
End Sub


Public Property Get Workbook() As Workbook
  If wb Is Nothing Then Set wb = ThisWorkbook
  Set Workbook = wb
End Property


Public Property Let Workbook(ByRef wbk As Workbook)
  Set wb = wbk
End Property


Public Property Get ListObject() As ListObject
  If lo Is Nothing Then
    Set lo = Util.Excel.GetListObjectByName(Me.Workbook, SETTINGS_LIST_NAME)
  End If
  Set ListObject = lo
End Property


Public Property Let ListObject(ByRef lobj As ListObject)
  Set lo = lobj
End Property


Public Property Get Worksheet() As Worksheet
  If WS Is Nothing Then
    Set WS = Me.ListObject.Parent
  End If
  Set Worksheet = WS
End Property


Public Property Let Worksheet(ByRef wsh As Worksheet)
  Set WS = wsh
End Property


Public Property Get Setting(ByVal id As SettingRowId) As Variant
  'checkNamesExistance 'not really necessary unless we expect someone to change names of settings before execution
  Setting = Me.Worksheet.Cells(getIdRow(id), Me.ListObject.ListColumns(COLUMN_NAME_VALUE).DataBodyRange.Column).value
End Property


Private Function getIdRow(ByVal id As SettingRowId) As Long
Dim idRange As Range
  Set idRange = Util.Excel.FindFirstInRange(SearchRange:=Me.ListObject.ListColumns(COLUMN_NAME_ID).DataBodyRange, FindWhat:=id)
  If idRange Is Nothing Then
    err.raise Number:=ErrCode.NamedArgumentNotFound, _
      Description:="No entry with the value " & _
        """" & id & """" & _
        " was found in range " & _
        """" & COLUMN_NAME_ID & " - " & Me.ListObject.ListColumns(COLUMN_NAME_ID).DataBodyRange.Address & """" & _
        " in table " & _
        """" & Me.ListObject.name & """" & _
        " in worksheet " & _
        """" & Me.Worksheet.name & """" & vbLf & _
        "Please check your settings definitions and setting IDs!"
  Else
    getIdRow = idRange.row
  End If
End Function


Private Sub checkNamesExistance()
  If Not Util.Excel.IsListObjectName(Me.Workbook, SETTINGS_LIST_NAME) Then
    err.raise Number:=ErrCode.NamedArgumentNotFound, _
        Description:="No table with the name """ & _
        SETTINGS_LIST_NAME & """ was found!" & vbLf & _
        "Please check Settings table name."
  End If
  Util.Excel.GetListObjectColumnIndex Me.ListObject, COLUMN_NAME_ID
  Util.Excel.GetListObjectColumnIndex Me.ListObject, COLUMN_NAME_VALUE
End Sub

