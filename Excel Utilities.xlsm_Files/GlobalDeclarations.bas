Attribute VB_Name = "GlobalDeclarations"
'---------------------------------------------------------------------------------------
' Module    : GlobalVariables
' Author    : Ivanov, Bozhan
' Purpose   : Initialization of global variables, constants and procedures
'---------------------------------------------------------------------------------------
Option Explicit

Public LAST_PATH As String
Public Const DEFAULT_DEVELOPER_NAME As String = "<name>"
Public Const DEFAULT_DEVELOPER_CONTACT As String = "<email>"

' Default error message
Public Const INFO_ERR_MSG As String = "For additional support, pleace contact " & _
  DEFAULT_DEVELOPER_NAME & _
  " at " & _
  DEFAULT_DEVELOPER_CONTACT

'====================================================================================================

Public Type FrameClasses
  Access As String
  Excel As String
  FrontPage As String
  Outlook As String
  PowerPoint_95 As String
  PowerPoint_97 As String
  PowerPoint_2000 As String
  PowerPoint_XP As String
  PowerPoint_2010 As String
  Project As String
  Word As String
  UserForm_97 As String
  UserForm_2000 As String
  VBE As String
End Type

Public Type AlertsAndUpdatingStatus
  ScreenUpdating As Boolean
  DisplayAlerts As Boolean
  EnableEvents As Boolean
End Type

Public Type OLEObjectSetting
  classType As Variant '"Forms.CommandButton.1", "Forms.ComboBox.1", ... (you must specify either ClassType or FileName). A string that contains the programmatic identifier for the object to be created. If ClassType is specified, FileName and Link are ignored.
  Height  As Variant '(you must specify either ClassType or FileName). A string that specifies the file to be used to create the OLE object.
  FileLink  As Variant 'True to have the new OLE object based on FileName be linked to that file. If the object isn't linked, the object is created as a copy of the file. The default value is False.
  DisplayAsIcon As Variant 'True to display the new OLE object either as an icon or as its regular picture. If this argument is True, IconFileName and IconIndex can be used to specify an icon.
  IconFileName   As Variant 'A string that specifies the file that contains the icon to be displayed. This argument is used only if DisplayAsIcon is True. If this argument isn't specified or the file contains no icons, the default icon for the OLE class is used.
  IconIndex As Variant 'The number of the icon in the icon file. This is used only if DisplayAsIcon is True and IconFileName refers to a valid file that contains icons. If an icon with the given index number doesn't exist in the file specified by IconFileName, the first icon in the file is used.
  IconLabel  As Variant 'A string that specifies a label to display beneath the icon. This is used only if DisplayAsIcon is True. If this argument is omitted or is an empty string (""), no caption is displayed.
  Left As Variant 'The initial coordinates of the new object, in points, relative to the upper-left corner of cell A1 on a worksheet, or to the upper-left corner of a chart.
  Width  As Variant 'The initial size of the new object, in points.
  Top  As Variant 'The initial coordinates of the new object in points, relative to the upper-left corner of cell A1 on a worksheet, or to the upper-left corner of a chart.
  LinkedCell  As Variant
  Enabled  As Variant
  Visible  As Variant
  ListFillRange As Variant
End Type

'====================================================================================================

Private utils As Utility
Private sett As Settings
Private errh As ErrorHandler
Private cgen As ClassGenerator
Public tmp As Collection

Public Property Get Temp() As Collection
  If tmp Is Nothing Then Set tmp = New Collection
  Set Temp = tmp
End Property

Public Property Get Util() As Utility
  If utils Is Nothing Then Set utils = New Utility
  Set Util = utils
End Property

Public Property Get Setting(ByVal rid As SettingRowId) As Variant
  If sett Is Nothing Then Set sett = New Settings
  Setting = sett.Setting(rid)
End Property

Public Property Get ErrHandler() As ErrorHandler
  If errh Is Nothing Then Set errh = New ErrorHandler
  Set ErrHandler = errh
End Property

Public Property Get ClassGen() As ClassGenerator
  If cgen Is Nothing Then Set cgen = New ClassGenerator
  Set ClassGen = cgen
End Property

'====================================================================================================

Public Function InfoErrMsg() As String
On Error GoTo InfoErrMsg_Error
Dim errMsg As String
  
  errMsg = vbNullString
  errMsg = "For additional support, pleace contact " & _
    Setting(ToolSupportContactName) & _
    " at " & _
    Setting(ToolSupportContactEmail)
  
InfoErrMsg_Exit:
On Error Resume Next
  InfoErrMsg = errMsg
Exit Function

InfoErrMsg_Error:
  errMsg = INFO_ERR_MSG
  Debug.Print err.Number, err.Description
Resume InfoErrMsg_Exit
End Function

Public Sub ButtonEventHandler()
Dim em As EventHandler
  Set em = New EventHandler
  em.HandleEvent Application.caller
  Set em = Nothing
End Sub

Public Sub GenerateSettignsClass()
Dim cg As New ClassGenerator
  Debug.Print cg.GenerateSettingsClassFile(ThisWorkbook, "Settings")
  Set cg = Nothing
End Sub

'Replacement for the Nz() function from MS Access library
Public Function Nz(ByVal value As Variant, Optional ByVal ValueIfNull = "")
    Nz = IIf(IsNull(value), ValueIfNull, value)
End Function

'aqcuire the ribbon pointer on ribbon load
Private Sub pmUI_onLoad(ribbon As IRibbonUI)
  Temp.Add ObjPtr(ribbon), "pmUIRibbonPointer"
End Sub

