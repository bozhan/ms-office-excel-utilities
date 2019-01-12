VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : ClassGenerator
' Author    : Ivanov, Bozhan
' Date      : 27.03.2018
' Purpose   : Generates class files based on properties in provided workbook and/or
'   provided VBA project in a workbook
'---------------------------------------------------------------------------------------
Const div = "  "
Const lf = vbCrLf
'---------------------------------------------------------------------------------------
' Procedure : generateSettingsClassFile
' Purpose   : generates settings class module based on the settings defined in a settings
'   worksheet and a settings listtable object
'   Returns the path of the class file on successful creation
'---------------------------------------------------------------------------------------
Public Function GenerateSettingsClassFile( _
  ByRef wb As Workbook, _
  ByVal ClassName As String _
) As String
On Error GoTo generateSettingsClassFile_Error
  
  GenerateSettingsClassFile = ""
  
  Dim FilePath As String
  FilePath = GetNewClassFilePath(wb, ClassName)
    
  Dim stream As TextStream
  Set stream = GetNewStream(FilePath)
  WriteToStream stream, GetSettingsClassContent(wb, ClassName)
  stream.Close
  GenerateSettingsClassFile = FilePath
  
generateSettingsClassFile_Exit:
On Error Resume Next
Exit Function

generateSettingsClassFile_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure generateSettingsClassFile of UtilClassModule:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume generateSettingsClassFile_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : generateUtilityClassFile
' Purpose   : generates a Utility class module based on the Utility classes in the curent
'   project. Utility classes are defined by the utils name prefix
'---------------------------------------------------------------------------------------
Public Function GenerateUtilityClassFile( _
  ByRef wb As Workbook, _
  ByVal ClassName As String, _
  Optional ByVal prefix As String = "util", _
  Optional ByVal exclude As String = "" _
) As String
On Error GoTo generateUtilityClassFile_Error
  
  GenerateUtilityClassFile = ""
  
  Dim FilePath As String
  FilePath = GetNewClassFilePath(wb, ClassName)
    
  Dim stream As TextStream
  Set stream = GetNewStream(FilePath)
  
  If Len(exclude) = 0 Then exclude = ClassName
  WriteToStream stream, GetUtilityClassContent(wb, ClassName, prefix, exclude)
  stream.Close
  GenerateUtilityClassFile = FilePath
  
generateUtilityClassFile_Exit:
On Error Resume Next
Exit Function

generateUtilityClassFile_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure generateUtilityClassFile of UtilClassModule:Class Module" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume generateUtilityClassFile_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : getNewClassFilePath
' Author    : Ivanov, Bozhan
' Purpose   : returns a new class file path based on workbook location
'   and class name provided
'---------------------------------------------------------------------------------------
Private Function GetNewClassFilePath( _
  ByRef wb As Workbook, _
  ByVal ClassName As String _
) As String
Dim fileName As String
On Error GoTo getNewClassFilePath_Error
  
  fileName = ClassName & ".cls" '& dt.getDateStamp & ".cls"
  GetNewClassFilePath = Util.File.GetUNCPath(wb.path) & "\" & fileName
  
getNewClassFilePath_Exit:
Exit Function

getNewClassFilePath_Error:
Debug.Print "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure getNewClassFilePath of Class Module ClassGenerator" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
GetNewClassFilePath = ""
Resume getNewClassFilePath_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : getNewStream
' Author    : Ivanov, Bozhan
' Purpose   : Returns a new stream object based on the provided path for the new class
'---------------------------------------------------------------------------------------
Private Function GetNewStream( _
  ByVal FilePath As String _
) As TextStream
  Set GetNewStream = Util.File.CreateTextFile(FilePath, True, False)
End Function

'---------------------------------------------------------------------------------------
' Procedure : writeToStream
' Author    : Ivanov, Bozhan
' Purpose   : writes attribute values, lines are defined as array entries
'---------------------------------------------------------------------------------------
Private Sub WriteToStream( _
  ByRef stream As TextStream, _
  ByVal lines As Variant _
)
Dim ln As New ArrayList
  
  ln.Append lines
  Dim l As Long
  
  For l = ln.lb To ln.ub
    stream.WriteLine ln.value(l)
  Next l
  stream.WriteBlankLines 1
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getSettingsClassContent
' Author    : Ivanov, Bozhan
' Purpose   : Returns an array of lines of content specific for Settings Class
'---------------------------------------------------------------------------------------
Private Function GetSettingsClassContent( _
  ByRef wb As Workbook, _
  ByVal ClassName As String _
) As Variant
Dim settingsListName As String, descriptionColName As String, valueColName As String, idColName As String
settingsListName = "GlobalSettings"
descriptionColName = "Description"
valueColName = "Value"
idColName = "ID"

Dim lines As New ArrayList
  lines.Append "VERSION 1.0 CLASS"
  lines.Append "BEGIN"
  lines.Append div & "MultiUse = -1  'True"
  lines.Append "END"
  lines.Append "Attribute VB_Name = """ & ClassName & """"
  lines.Append "Attribute VB_GlobalNameSpace = False"
  lines.Append "Attribute VB_Creatable = False"
  lines.Append "Attribute VB_PredeclaredId = False"
  lines.Append "Attribute VB_Exposed = False"
  lines.Append "'---------------------------------------------------------------------------------------"
  lines.Append "' Module    : " & ClassName
  lines.Append "' Author    : " & DEFAULT_DEVELOPER_NAME & " - " & "(email: " & DEFAULT_DEVELOPER_CONTACT & ")"
  lines.Append "' Purpose   : Automatically generated class for Tool Template Excel by Bozhan Ivanov"
  lines.Append "' Comment   : Default error handling should be avoided here because it references the"
  lines.Append "'   settings class."
  lines.Append "'---------------------------------------------------------------------------------------"
  lines.Append "Option Explicit"
  lines.Append "Private Const SETTINGS_LIST_NAME = """ & settingsListName & """"
  lines.Append "Private Const COLUMN_NAME_ID = """ & idColName & """"
  lines.Append "Private Const COLUMN_NAME_VALUE = """ & valueColName & """"
  lines.Append "Private Const COLUMN_NAME_DESCRIPTION = """ & descriptionColName & """"
  lines.Append lf
  lines.Append GetEnumStructureForSettingRowId(wb, settingsListName, idColName, descriptionColName)
  lines.Append lf
  lines.Append "Private wb As Workbook"
  lines.Append "Private lo As ListObject"
  lines.Append "Private ws As Worksheet"
  lines.Append lf
  lines.Append "Private Sub Class_Initialize()"
  lines.Append div & "checkNamesExistance"
  lines.Append "End Sub"
  lines.Append lf
  lines.Append "Private Sub Class_Terminate()"
  lines.Append div & "Set wb = Nothing"
  lines.Append div & "Set lo = Nothing"
  lines.Append div & "Set ws = Nothing"
  lines.Append "End Sub"
  lines.Append lf
  lines.Append "Public Property Get Workbook() As Workbook"
  lines.Append div & "If wb Is Nothing Then Set wb = ThisWorkbook"
  lines.Append div & "Set Workbook = wb"
  lines.Append "End Property"
  lines.Append lf
  lines.Append "Public Property Let Workbook(ByRef wbk As Workbook)"
  lines.Append div & "Set wb = wbk"
  lines.Append "End Property"
  lines.Append lf
  lines.Append "Public Property Get ListObject() As ListObject"
  lines.Append div & "If lo Is Nothing Then"
  lines.Append div & div & "Set lo = Util.Excel.GetListObjectByName(Me.Workbook, SETTINGS_LIST_NAME)"
  lines.Append div & "End If"
  lines.Append div & "Set ListObject = lo"
  lines.Append "End Property"
  lines.Append lf
  lines.Append "Public Property Let ListObject(ByRef lobj As ListObject)"
  lines.Append div & "Set lo = lobj"
  lines.Append "End Property"
  lines.Append lf
  lines.Append "Public Property Get Worksheet() As Worksheet"
  lines.Append div & "If ws Is Nothing Then"
  lines.Append div & div & "Set ws = Me.ListObject.Parent"
  lines.Append div & "End If"
  lines.Append div & "Set Worksheet = ws"
  lines.Append "End Property"
  lines.Append lf
  lines.Append "Public Property Let Worksheet(ByRef wsh As Worksheet)"
  lines.Append div & "Set ws = wsh"
  lines.Append "End Property"
  lines.Append lf
  lines.Append "Public Property Get Setting(ByVal id As SettingRowId) As Variant"
  lines.Append div & "'checkNamesExistance 'not really necessary unless we expect someone to change names of settings before execution"
  lines.Append div & "Setting = Me.Worksheet.Cells(getIdRow(id), Me.ListObject.ListColumns(COLUMN_NAME_VALUE).DataBodyRange.Column).Value"
  lines.Append "End Property"
  lines.Append lf
  lines.Append "Private Function getIdRow(ByVal id As SettingRowId) As Long"
  lines.Append "Dim idRange As Range"
  lines.Append div & "Set idRange = Util.Excel.FindFirstInRange(searchRange:=Me.ListObject.ListColumns(COLUMN_NAME_ID).DataBodyRange, FindWhat:=id)"
  lines.Append div & "If idRange Is Nothing Then"
  lines.Append div & "  err.raise Number:=ErrCode.NamedArgumentNotFound, _"
  lines.Append div & "    Description:=""No entry with the value "" & _"
  lines.Append div & "      """""""" & id & """""""" & _"
  lines.Append div & "      "" was found in range "" & _"
  lines.Append div & "      """""""" & COLUMN_NAME_ID & "" - "" & Me.ListObject.ListColumns(COLUMN_NAME_ID).DataBodyRange.Address & """""""" & _"
  lines.Append div & "      "" in table "" & _"
  lines.Append div & "      """""""" & Me.ListObject.name & """""""" & _"
  lines.Append div & "      "" in worksheet "" & _"
  lines.Append div & "      """""""" & Me.Worksheet.name & """""""" & vbLf & _"
  lines.Append div & "      ""Please check your settings definitions and setting IDs!"""
  lines.Append div & "Else"
  lines.Append div & "  getIdRow = idRange.row"
  lines.Append div & "End If"
  lines.Append "End Function"
  lines.Append lf
  lines.Append "Private Sub checkNamesExistance()"
  lines.Append div & "If Not Util.Excel.IsListObjectName(Me.Workbook, SETTINGS_LIST_NAME) Then"
  lines.Append div & "  err.raise Number:=ErrCode.NamedArgumentNotFound, _"
  lines.Append div & "      Description:=""No table with the name """""" & _"
  lines.Append div & "      SETTINGS_LIST_NAME & """""" was found!"" & vbLf & _"
  lines.Append div & "      ""Please check Settings table name."""
  lines.Append div & "End If"
  lines.Append div & "Util.Excel.GetListObjectColumnIndex Me.ListObject, COLUMN_NAME_ID"
  lines.Append div & "Util.Excel.GetListObjectColumnIndex Me.ListObject, COLUMN_NAME_VALUE"
  lines.Append "End Sub"
  
  GetSettingsClassContent = lines.Values
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetEnumStructureForSettingRowId
' Author    : Ivanov, Bozhan
' Purpose   : Generates Enum structure for the setting table defined in workbook
'---------------------------------------------------------------------------------------
Private Function GetEnumStructureForSettingRowId( _
  ByRef wb As Workbook, _
  ByVal settingsListName As String, _
  ByVal idColName As String, _
  ByVal descriptionColName As String _
) As Variant
Dim lo As ListObject
Dim lr As ListRow
Dim cid As Long
Dim lines As New ArrayList
Dim rng As Range

  lines.Append "Public Enum SettingRowId"

  Set lo = Util.Excel.GetListObjectByName(wb, settingsListName)
  cid = Util.Excel.GetListObjectColumnIndex(lo, idColName)
  
  For Each lr In lo.ListRows
    Set rng = lo.DataBodyRange(lr.Index, Util.Excel.GetListObjectColumnIndex(lo, descriptionColName))
    
    If Len(trim(rng.value)) > 0 _
    And Len(trim(lo.DataBodyRange(lr.Index, cid).value)) > 0 Then
      lines.Append "  " & GenerateVariableNameFromCellContent(rng) & " = " & lo.DataBodyRange(lr.Index, cid).value
    End If
  Next lr
  
  lines.Append "End Enum"
  GetEnumStructureForSettingRowId = lines.Values
End Function

Private Function GenerateVariableNameFromCellContent( _
  ByRef rng As Range _
) As String
  Dim s As Variant
  
  s = Util.Strings.concat(Util.RegEx.match(Util.RegExPatterns.VariableNamesPattern, rng.value))
  
  Dim idx As Long
  idx = Util.Strings.GetFirstIndexOfCharType(s, Letter)
  
  If idx = 1 Then
    GenerateVariableNameFromCellContent = s
  ElseIf idx > 1 Then
    GenerateVariableNameFromCellContent = Util.Strings.GetRightOfIndex(s, Util.Strings.GetFirstIndexOfCharType(s, Letter) - 1)
  Else
    GenerateVariableNameFromCellContent = "tmp_" & Util.Strings.getRandomString_Alphabetic(10)
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetUtilityClassContent
' Author    : Ivanov, Bozhan
' Purpose   : Returns an array of lines of content specific for Utility Class
'---------------------------------------------------------------------------------------
Private Function GetUtilityClassContent( _
  ByRef wb As Workbook, _
  Optional ClassName As String = "Utility", _
  Optional ByVal prefix As String = "util", _
  Optional ByVal exclude As String = "Utility" _
) As Variant
Dim lines As New ArrayList
  lines.Append "VERSION 1.0 CLASS"
  lines.Append "BEGIN"
  lines.Append div & "MultiUse = -1  'True"
  lines.Append "END"
  lines.Append "Attribute VB_Name = """ & ClassName & """"
  lines.Append "Attribute VB_GlobalNameSpace = False"
  lines.Append "Attribute VB_Creatable = False"
  lines.Append "Attribute VB_PredeclaredId = False"
  lines.Append "Attribute VB_Exposed = False"
  lines.Append "'---------------------------------------------------------------------------------------"
  lines.Append "' Module    : " & ClassName
  lines.Append "' Author    : " & DEFAULT_DEVELOPER_NAME & " - " & "(email: " & DEFAULT_DEVELOPER_CONTACT & ")"
  lines.Append "' Purpose   : Automatically generated class for Tool Template Excel by Bozhan Ivanov"
  lines.Append "'---------------------------------------------------------------------------------------"
  lines.Append "Option Explicit"
  
  Dim classes As Variant, i As Integer
  classes = GetUtilityClassesNames(wb, prefix, exclude)
  
  For i = LBound(classes) To UBound(classes)
    lines.Append "Dim " & "util" & CStr(i) & " As " & classes(i)
  Next i
  
  lines.Append lf
  lines.Append "Private Sub Class_Initialize()"
    For i = LBound(classes) To UBound(classes)
      lines.Append div & "Set " & prefix & CStr(i) & " = New " & classes(i)
    Next i
  lines.Append "End Sub"
  
  lines.Append lf
  lines.Append "Private Sub Class_Terminate()"
    For i = LBound(classes) To UBound(classes)
      lines.Append div & "Set " & "util" & CStr(i) & " = Nothing"
    Next i
  lines.Append "End Sub"
  
  lines.Append lf
  For i = LBound(classes) To UBound(classes)
    Dim name As String
    
    name = Right$(classes(i), Len(classes(i)) - InStr(1, classes(i), prefix, vbTextCompare) - Len(prefix) + 1)
    lines.Append "Property Get " & name & "()" & " As " & classes(i)
    lines.Append div & "Set " & name & " = " & "util" & CStr(i)
    lines.Append "End Property"
  Next i
  
  lines.Append lf
  lines.Append "Public Sub Refresh()"
  lines.Append div & "Call Class_Initialize"
  lines.Append "End Sub"
  
  GetUtilityClassContent = lines.Values
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetUtilityClassesNames
' Author    : Ivanov, Bozhan
' Purpose   : returns the names of classes that match prefix and are not part of excluded
'   Excluded can be multiple values in CSV form
'---------------------------------------------------------------------------------------
Public Function GetUtilityClassesNames( _
  ByRef wb As Workbook, _
  Optional ByVal prefix As String = "util", _
  Optional exclude As String = "Utility" _
) As Variant
Dim comp As Variant
Dim excluded As New ArrayList
Dim result As New ArrayList
  
  excluded.Append Split(exclude, ",")
  
  For Each comp In wb.VBProject.VBComponents
    If comp.Type = ComponentTypeEnum.ClassModule Then
      If StrComp(Left(comp.name, Len(prefix)), prefix, vbTextCompare) = 0 _
      And Not excluded.Contains(comp.name) Then
        result.Append comp.name
      End If
    End If
  Next comp

  GetUtilityClassesNames = result.Values
End Function

