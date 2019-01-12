VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module      : UtilModules
' Author      : Ivanov, Bozhan
' Description : Modules, Classes and Forms maangement functions
'---------------------------------------------------------------------------------------
Option Explicit

Public Enum ComponentTypeEnum
  ActiveXDesigner = 11
  ClassModule = 2
  Document = 100
  MSForm = 3
  StdModule = 1
End Enum

Property Get AllowedModuleExtensions() As Variant
  AllowedModuleExtensions = Array("cls", "frm", "bas")
End Property

'---------------------------------------------------------------------------------------
' Procedure   : ExportModules
' Description : Exports all current project modules (cls, frm, bas) as files
'   to the desired folder.
'   If the folder does not exist it will be cerated.
'   By default the modules will be exported to folder path 'ThisWorkbook.FullName & "_Files"'
'---------------------------------------------------------------------------------------
Public Sub ExportModules( _
  ByRef wb As Workbook, _
  Optional ByVal dirPath As String = "", _
  Optional ByVal overwrite As Boolean = True _
)
  If wb.VBProject.Protection = 1 Then
    err.raise Number:=ErrCode.CannotPerformRequestedOperation, _
      Description:="The VBA in this workbook is protected, not possible to perform export!"
  End If
  
  Dim f As folder
  If Len(dirPath) = 0 Then 'assign default dir path value
    dirPath = GetDefaultProjectFilesFolderPath(wb)
  ElseIf Not Util.File.FolderExists(dirPath) Then
    Call Util.File.CreateFolder(dirPath)
  End If
  
  Set f = Util.File.GetFolder(dirPath)
  
  Dim cmp As Variant, cmpPath As String
  For Each cmp In wb.VBProject.VBComponents
    Select Case cmp.Type
      Case ComponentTypeEnum.ClassModule
        cmpPath = f.path & "\" & cmp.name & ".cls"
      Case ComponentTypeEnum.MSForm
        cmpPath = f.path & "\" & cmp.name & ".frm"
      Case ComponentTypeEnum.StdModule
        cmpPath = f.path & "\" & cmp.name & ".bas"
      Case ComponentTypeEnum.Document
        GoTo next_component
      Case ComponentTypeEnum.ActiveXDesigner
        GoTo next_component
    End Select
    
    If Util.File.FileExists(cmpPath) Then
      If overwrite Then
        Kill cmpPath
      Else
        GoTo next_component
      End If
    End If
    
    cmp.Export cmpPath
    'wkbSource.VBProject.VBComponents.Remove cmpComponent
next_component:
  Next cmp
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : ImportModules
' Description : Imports all modules (cls, bas) found in a given folder
'   (non recursive search)
'---------------------------------------------------------------------------------------
Public Sub ImportModules( _
  ByRef wb As Workbook, _
  Optional ByVal dirPath As String = "" _
)
  If wb.VBProject.Protection = 1 Then
    err.raise Number:=ErrCode.CannotPerformRequestedOperation, _
      Description:="The VBA in this workbook is protected, not possible to perform export!"
  End If
  
  If Len(dirPath) = 0 Then 'assign default dir path value
    dirPath = GetDefaultProjectFilesFolderPath(wb)
  ElseIf Not Util.File.FolderExists(dirPath) Then
    err.raise Number:=ErrCode.CannotPerformRequestedOperation, _
      Description:="Folder not found: " & dirPath
  End If
  
  Dim f As folder, fl As File, i As Integer, ext As String
  
  Set f = Util.File.GetFolder(dirPath)
  For Each fl In f.Files
    ext = Util.File.GetFileExtension(fl.name)
    For i = LBound(AllowedModuleExtensions) To UBound(AllowedModuleExtensions)
      If StrComp(ext, AllowedModuleExtensions(i), vbTextCompare) = 0 Then
        wb.VBProject.VBComponents.Import fl.path
      End If
    Next i
  Next fl
End Sub

Function GetDefaultProjectFilesFolderPath(ByRef wb As Workbook) As String
  GetDefaultProjectFilesFolderPath = Util.File.CreateFolder(wb.FullName & "_Files").path
End Function

Public Sub DeleteComponentWithName( _
  ByRef wb As Workbook, _
  ByVal name As String _
)
  If ExistsComponentWithName(wb, name) Then
    wb.VBProject.VBComponents.Remove wb.VBProject.VBComponents(name)
  End If
End Sub

Public Function ExistsComponentWithName( _
  ByRef wb As Workbook, _
  ByVal name As String _
) As Boolean
  ExistsComponentWithName = Not (GetComponentWithName(wb, name) Is Nothing)
End Function

Public Function GetComponentWithName( _
  ByRef wb As Workbook, _
  ByVal name As String _
) As Object
  Set GetComponentWithName = Nothing
  Dim cmp As Object
  For Each cmp In wb.VBProject.VBComponents
    If StrComp(name, cmp.name, vbTextCompare) = 0 Then
      Set GetComponentWithName = cmp
      Exit Function
    End If
  Next cmp
End Function

Public Sub DeleteComponents( _
  ByRef wb As Workbook _
)
  Dim cmp As Object
  For Each cmp In wb.VBProject.VBComponents
    If StrComp("UtilModules", cmp.name, vbTextCompare) <> 0 Then
      If cmp.Type = ComponentTypeEnum.ClassModule _
      Or cmp.Type = ComponentTypeEnum.MSForm _
      Or cmp.Type = ComponentTypeEnum.StdModule Then
        wb.VBProject.VBComponents.Remove cmp
      End If
    End If
  Next cmp
End Sub




