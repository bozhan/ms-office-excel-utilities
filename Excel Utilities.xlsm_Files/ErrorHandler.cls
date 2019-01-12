VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ErrCode
  NoError = 0
  InvalidProcedureCall = 5
  Overflow = 6
  OutOfMemory = 7
  SubscriptOutOfRange = 9
  ArrayIsFixedOrLocked = 10
  DivisionByZero = 11
  TypeMismatch = 13
  OutOfStringSpace = 14
  CannotPerformRequestedOperation = 17
  OutOfStackSspace = 28
  FileNotFound = 53
  BadFileMode = 54
  FileAlreadyOpen = 55
  DeviceIOError = 57
  FileAlreadyExists = 58
  PermissionDenied = 70
  PathFileAccessError = 75
  PathNotFound = 76
  ObjectVariableNotSet = 91
  InvalidFileFormat = 321
  CannotAccessRegistry = 335
  ObjectRequired = 424
  ObjectDoesnotSupportProperty = 438
  NamedArgumentNotFound = 448
  ArgumentNotOptional = 449
  ErrorSavingFile = 31036
  ErrorLoadingFile = 31037
End Enum

Private Type SavedError
  Description As String
  Number As Long
  source As String
  HelpFile As String
  HelpContext As Long
  LastDllError As Long
End Type

Dim er As SavedError

Public Sub Save(ByRef e As ErrObject)
  er.Number = e.Number
  er.Description = e.Description
  er.source = e.source
  er.HelpContext = e.HelpContext
  er.HelpFile = e.HelpFile
  er.LastDllError = e.LastDllError
End Sub

Public Sub RaiseLastSaved()
  err.raise er.Number, er.source, er.Description, er.HelpFile, er.HelpContext
End Sub

Property Get Number() As Long
  Number = er.Number
End Property

Property Get Description() As String
  Description = er.Description
End Property

Property Get source() As String
  source = er.source
End Property

Private Sub Class_Initialize()
  er.Number = 0
  er.HelpContext = 0
  er.LastDllError = 0
  er.Description = ""
  er.source = ""
  er.HelpFile = ""
End Sub

'---------------------------------------------------------------------------------------
' Procedure : logErrorToFile
' Purpose   : Logs the occured error with a timestamp to an external log file
'---------------------------------------------------------------------------------------
'
Public Sub logErrorToFile(err As ErrObject, errSource As String)
Dim logFilePath As String
Dim logString As String
Dim logDir As String
On Error GoTo logErrorToFile_Error

    logDir = "C:\Temp"
    
    If Not Util.File.FolderExists(logDir) Then _
      err.raise Number:=ErrCode.PathNotFound, _
        Description:="log path was not found or invalid"
    
    logFilePath = logDir & "\" & Util.DateTime.GetDateStamp() & "_PIdb_ErrorLog.txt"
      
    'Open and wrtie title for log
    Open logFilePath For Append As #1
    logString = Now() & ":" & _
            Util.Windows.UserName & VBA.vbTab & _
            "Error " & err.Number & " (" & err.Description & ")" & " in " & _
            errSource
    Print #1, logString
    Close #1
    
logErrorToFile_Exit:
On Error Resume Next
  Exit Sub

logErrorToFile_Error:
  'logErrorToDb err, "logErrorToFile of Module modGlobal"
  Resume logErrorToFile_Exit
End Sub


