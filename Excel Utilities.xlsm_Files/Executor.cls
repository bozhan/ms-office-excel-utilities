VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Executor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pb As ProgressBarController
'TODO consider everything up to getsettings into a headColumn to go to
'a seperate class where we can define properties that return the headers for settings tables

Property Get SourcesFieldNames() As Variant
  SourcesFieldNames = Array("Worksheet Name", "Workbook Path")
End Property

Property Get MappingFieldNames() As Variant
  MappingFieldNames = Array("ID", "Worksheet Name", "Column Headers", "Field Type", "New Column Header")
End Property

Property Get ConditionsFieldNames() As Variant
  ConditionsFieldNames = Array("Worksheet Name", "Column Headers", "Condition", "Value")
End Property

Property Get FieldTypesFieldNames() As Variant
  FieldTypesFieldNames = Array("Type")
End Property

Public Sub ImportSources()
On Error GoTo ImportSources_Error
  
  CheckSourcesTableConsistency
  
  Dim wbPath As String
  wbPath = Util.File.GetFilePathViaDialogWindow( _
    InitialFilename:=Util.File.LastPath, _
    title:="Choose Head Count Source File...", _
    FilterDesc:="Excel", _
    FilterExt:=Util.File.DefaultExcelExt)

  If wbPath = "" Then Exit Sub
  Util.File.LastPath = wbPath
  
  CheckWorkbookExistance wbPath
  
  Dim app As Excel.Application, wb As Excel.Workbook
  Set app = CreateObject("Excel.Application")
  app.Visible = False
  Set wb = Util.Excel.GetWorkbookFromPath(app, wbPath)
  
  Dim selectedSheetNames As Variant
  selectedSheetNames = Util.Form.GenerateListChoiceForm( _
    Util.Excel.GetWorksheetNames(wb), _
    "Select worksheet(s) from list...", _
    "Select worksheet(s) from """ & wb.name & """ to import as source", _
    "", _
    fmMultiSelectExtended)
  
  Dim sources As ListObject
  Set sources = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(SourcesTableName))
  
  'TODO check for name conflicts before inserting source
  
  Dim i As Integer
  For i = LBound(selectedSheetNames) To UBound(selectedSheetNames)
    If sources.ListRows(1).Range.Insert(XlInsertShiftDirection.xlShiftDown, False) Then
      sources.DataBodyRange(1, sources.ListColumns("Worksheet Name").Index).value = selectedSheetNames(i)
      sources.DataBodyRange(1, sources.ListColumns("Workbook Path").Index).value = Util.File.GetUNCPath(wb.FullName)
    Else
      err.raise Number:=ErrCode.PermissionDenied, _
        Description:="Could not enter new row into table """ & sources.name & """!"
    End If
  Next i
  
  If ExistsDuplicateSourceWorksheet() Then
    MsgBox "Duplicate sources found in table """ & Setting(SourcesTableName) & """" & vbLf & _
      "Worksheet names in Sources must be unique!", _
      vbInformation
  End If
  
ImportSources_Exit:
  If Not app Is Nothing Then
    Util.Excel.CloseWorbooksInApplication app
    app.Quit
  End If
Exit Sub

ImportSources_Error:
MsgBox "Error " & err.Number & " (" & err.Description & ") " & _
  "in procedure ImportSources of Class Module Executor" & vbLf & _
  InfoErrMsg(), vbExclamation, "Error occurred..."
Resume ImportSources_Exit
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GenerateDiscrepancyReport
' Author    : Ivanov, Bozhan
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub GenerateDiscrepancyReport()
On Error GoTo GenerateDiscrepancyReport_Error
  Set pb = New ProgressBarController
  pb.Caption = "Generating Head Count Discrepancy Report..."
  pb.Show
  
  CheckReportSettings
  
  Dim app As Excel.Application
  Set app = CreateObject("Excel.Application")
  app.Visible = False
  Util.Excel.TurnOffScreenUpdatingAndAlerts app

  Dim sources As Collection, src As WorksheetIO
  Set sources = GetSources(app)
  
  Dim conditions As Collection
  Set conditions = GetConditions()
  
  Dim prims As New ArrayList
  prims.Values = GetUniquePrimaryKeys(GetKeyHeadColumnsFromSources(sources))

  Dim dst As WorksheetIO
  Set dst = GetDestination(app)
  
  Dim keyIndeces As Collection
  Set keyIndeces = GetSheetKeyIndeces()
  
  Dim progressLimit As Single, dstamp As Date, avg As Single, refRate As Single
  'set 90% limit of the left progress for the overall progress on steps
  progressLimit = pb.Progress + pb.ProgressLeft * 0.9
  pb.TextProgress = "Generating Report"
  pb.IsCancelable = True
  refRate = 0.03
  
  Dim l As Long, max As Long, rowOffset As Long
  rowOffset = 1
  max = prims.ub
  
  For l = prims.lb To max
    Util.Timer.StartCounter
    
    'nz(trim()) will deal with null and empty values
    If Len(Nz(trim(prims(l)))) > 0 Then
      ImportSearchResults _
        sources, _
        keyIndeces, _
        dst, _
        prims(l), _
        conditions, _
        rowOffset
    End If
    
    'Updating ETR
    avg = avg + (Util.Timer.ElapsedTime(MagnitudeOrder.one, 3) * (max - l))
    If l Mod Int(max * refRate) = 0 Then
      dstamp = TimeSerial(0, 0, avg / (Int(max * refRate)))
      avg = 0
      pb.TextProgress = "Generating Report..." & " | ETR " & Format$(dstamp, "hh:nn:ss")
    End If
    
    'Detecting Cancel command
    If pb.IsCanceled Then
      pb.Caption = "Report canceled..."
      pb.TextResult = "Report generating process canceled!"
      pb.AlertIcon.Visible = True
      GoTo GenerateDiscrepancyReport_Exit
    End If
    
    pb.Progress = (progressLimit * (l / max))
  Next l
  
  pb.IncreseByPercent
  pb.TextProgress = "Formatting Report"
  FormatReportSheet dst
  
  pb.Progress = pb.ProgressMax
  pb.TextProgress = ""
  pb.TextResult = "Report successfully generated in workbook:" & vbLf & vbLf & _
    dst.Workbook.FullName
  pb.SuccessIcon.Visible = True
  
GenerateDiscrepancyReport_Exit:
  If Not sources Is Nothing Then
    For Each src In sources
      If Util.Excel.IsWorkbookOpen(app, src.Workbook) Then
        src.HideRevealedData
        src.Workbook.Close False
      End If
    Next src
  End If
  
  If Not dst Is Nothing Then
    If Util.Excel.IsWorkbookOpen(app, dst.Workbook) Then
      dst.Workbook.Close True
    End If
  End If
  
  If Not app Is Nothing Then
    Util.Excel.TurnOnScreenUpdatingAndAlerts app
    app.Quit
  End If
  pb.ShowResultControls
Exit Sub

GenerateDiscrepancyReport_Error:
  Debug.Print err.Number, err.Description
  pb.TextResult = "Error " & err.Number & _
    " in procedure GenerateDiscrepancyReport of Class Module Executor" & vbLf & vbLf & _
    err.Description
  pb.Caption = "Error occurred ..."
  pb.AlertIcon.Visible = True
Resume GenerateDiscrepancyReport_Exit
End Sub

Private Sub CheckReportSettings()
  Dim progressLimit As Single, steps As Integer
  progressLimit = pb.Progress + pb.ProgressLeft * 0.05 'set 5% limit for the overall progress on steps
  steps = 6
  
  pb.TextProgress = "Checking Report Settings"
  
  pb.Progress = (progressLimit * (1 / steps))
  CheckSourcesTableConsistency
  
  pb.Progress = (progressLimit * (2 / steps))
  CheckFieldMappingTableConsistency
  
  pb.Progress = (progressLimit * (3 / steps))
  CheckConditionsTableConsistency
  
  pb.Progress = (progressLimit * (4 / steps))
  CheckFieldTypesTableConsistency
  
  pb.Progress = (progressLimit * (5 / steps))
  If ExistsDuplicateSourceWorksheet() Then
    err.raise Number:=ErrCode.InvalidProcedureCall, _
      Description:="Duplicate sources found in table """ & Setting(SourcesTableName) & """" & vbLf & _
      "Worksheet names in Sources must be unique!"
  End If
  
  pb.Progress = (progressLimit * (6 / steps))
  If Not ReportSettingsAreConsistentWithSources() Then
    err.raise Number:=ErrCode.NamedArgumentNotFound, _
      Description:="Sources do not match the input in worksheet Report Settings!" & vbLf & _
      "Please make sure that the Report Settigns are correctly defined according to the imported sources."
  End If
End Sub

Private Sub CheckSourcesTableConsistency()
  CheckListObjectConsistency ThisWorkbook, Setting(SourcesTableName), Me.SourcesFieldNames
End Sub

Private Sub CheckFieldMappingTableConsistency()
  CheckListObjectConsistency ThisWorkbook, Setting(ReportFieldMappingTableName), Me.MappingFieldNames
End Sub

Private Sub CheckConditionsTableConsistency()
  CheckListObjectConsistency ThisWorkbook, Setting(ReportFieldConditionsTableName), Me.ConditionsFieldNames
End Sub

Private Sub CheckFieldTypesTableConsistency()
  CheckListObjectConsistency ThisWorkbook, Setting(ReportFieldTypesTableName), Me.FieldTypesFieldNames
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : ReportSettingsAreConsistentWithSources
' Description : Returns false if worksheet names in report settings are used that are not
'   defined as worksheet names in the sources table
'---------------------------------------------------------------------------------------
Private Function ReportSettingsAreConsistentWithSources() As Boolean
  ReportSettingsAreConsistentWithSources = False
  
  Dim sources As ListObject, src As Variant, mapping As ListObject, map As Variant
  Set sources = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(SourcesTableName))
  Set mapping = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(ReportFieldMappingTableName))
  src = Application.Transpose(sources.ListColumns("Worksheet Name").Range.value)
  map = Application.Transpose(mapping.ListColumns("Worksheet Name").Range.value)
  map = Util.Arrays.RemoveDuplicates(map)
  
  Dim i As Integer
  For i = LBound(map) To UBound(map)
    If Not Util.Arrays.Contains(src, map(i)) Then
      ReportSettingsAreConsistentWithSources = False
      Exit Function
    End If
  Next i
  
  ReportSettingsAreConsistentWithSources = True
End Function

'---------------------------------------------------------------------------------------
' Procedure   : ExistsDuplicateSource
' Description : Checks for duplicate values on worksheet name and workbook path in sources table
'---------------------------------------------------------------------------------------
Private Function ExistsDuplicateSource() As Boolean
  Dim src As ListObject
  Set src = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(SourcesTableName))
  
  Dim sheetNames As Variant
  sheetNames = Application.Transpose(src.ListColumns("Worksheet Name").DataBodyRange.value)
  
  Dim dupl As Variant
  dupl = Util.Arrays.GetDuplicates(sheetNames)
  
  If Util.Arrays.IsEmpty(dupl) Then
    ExistsDuplicateSource = False
    Exit Function
  End If
  
  'compare workbook names based on worksheet names
  Dim i As Integer, rngs As Collection, r As Range, WS As Worksheet, bookNames As New ArrayList
  Set WS = src.Parent
  For i = LBound(dupl) To UBound(dupl) 'compare sheet names
    If Len(dupl(i)) > 0 Then
      Set rngs = Util.Excel.FindAllInRange(src.ListColumns("Worksheet Name").DataBodyRange, dupl(i))
      If rngs.count > 0 Then
        For Each r In rngs 'get workbooks for duplicate worksheet names
          bookNames.Append WS.Cells(r.row, src.ListColumns("Workbook Path").Range.Column).value
        Next r
        If Not Util.Arrays.IsEmpty(Util.Arrays.GetDuplicates(bookNames.Values)) Then
          ExistsDuplicateSource = True
          Exit Function
        End If
      End If
    End If
  Next i
End Function

'---------------------------------------------------------------------------------------
' Procedure   : ExistsDuplicateSourceWorksheet
' Description : Checks for duplicate values in worksheet names in sources table
'---------------------------------------------------------------------------------------
Private Function ExistsDuplicateSourceWorksheet() As Boolean
  ExistsDuplicateSourceWorksheet = Not Util.Arrays.IsEmpty( _
    Util.Arrays.GetDuplicates( _
    Application.Transpose( _
    Util.Excel.GetListObjectByName( _
    ThisWorkbook, Setting(SourcesTableName)).ListColumns("Worksheet Name").DataBodyRange.value)))
End Function


'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

Public Function GetConditions() As Collection
'  Dim lo As ListObject
'  Set lo = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(ReportFieldConditionsTableName))
'
'  pb.IncreseByPercent
'  pb.TextProgress = "Obtaining conditions"
'
'  Dim conds As New Collection
'  Dim fc As FieldCondition
'  Dim lr As ListRow
'  For Each lr In lo.ListRows
'    Set fc = New FieldCondition
'    fc.Headers = lo.DataBodyRange(lr.Index, lo.ListColumns("Column Headers").Index).Value
'    fc.WorksheetName = lo.DataBodyRange(lr.Index, lo.ListColumns("Worksheet Name").Index).Value
'    fc.Operator = lo.DataBodyRange(lr.Index, lo.ListColumns("Condition").Index).Value
'    fc.Value = lo.DataBodyRange(lr.Index, lo.ListColumns("Value").Index).Value
'    fc.alias = "cond" & CStr(lr.Index)
'    conds.Add fc, fc.alias
'  Next lr
'
'  Set GetConditions = conds
End Function

Public Function GetConditionsByWorksheet( _
  ByVal sheetName As String _
) As Collection
'  Dim lo As ListObject
'  Set lo = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(ReportFieldConditionsTableName))
'
'  Dim conds As New Collection
'  Dim fc As FieldCondition
'  Dim lr As ListRow
'  For Each lr In lo.ListRows
'    If StrComp( _
'      sheetName, _
'      lo.DataBodyRange(lr.Index, lo.ListColumns("Worksheet Name").Index).Value, _
'      vbTextCompare _
'    ) = 0 Then
'      Set fc = New FieldCondition
'      fc.Headers = lo.DataBodyRange(lr.Index, lo.ListColumns("Column Headers").Index).Value
'      fc.WorksheetName = lo.DataBodyRange(lr.Index, lo.ListColumns("Worksheet Name").Index).Value
'      fc.Operator = lo.DataBodyRange(lr.Index, lo.ListColumns("Condition").Index).Value
'      fc.Value = lo.DataBodyRange(lr.Index, lo.ListColumns("Value").Index).Value
'      conds.Add fc, "cond" & CStr(lr.Index)
'    End If
'  Next lr
'
'  Set GetConditionsByWorksheet = conds
End Function

Public Function GetSources( _
  ByRef app As Excel.Application _
) As Collection
  Dim res As New Collection
  Dim sourcesTable As ListObject
  Set sourcesTable = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(SourcesTableName))
  
  Dim lr As ListRow, src As WorksheetIO, wbPath As String, sheetName As String
  
  Dim progressLimit As Single, steps As Integer
  pb.IncreseByPercent
  progressLimit = pb.Progress + pb.ProgressLeft * 0.05 'set 5% limit for the overall progress on steps
  steps = sourcesTable.ListRows.count
  pb.TextProgress = "Obtaining sources"
  
  For Each lr In sourcesTable.ListRows
    wbPath = sourcesTable.DataBodyRange(lr.Index, sourcesTable.ListColumns("Workbook Path").Index)
    sheetName = sourcesTable.DataBodyRange(lr.Index, sourcesTable.ListColumns("Worksheet Name").Index)
    
    pb.TextProgress = "Obtaining source """ & sheetName & """"
    pb.Progress = (progressLimit * (lr.Index / steps))
    
    If Len(wbPath) > 0 And Len(sheetName) > 0 Then
      Set src = New WorksheetIO
      src.Application = app
      
      CheckWorkbookExistance wbPath
      If Not Util.Excel.IsWorkbookOpen(app, Util.File.GetFileName(wbPath)) Then
        src.OpenWorkbookFromPath wbPath
      Else
        src.Workbook = app.Workbooks(Util.File.GetFileName(wbPath))
      End If
      
      CheckWorksheetExistance src.Workbook, sheetName
      src.Worksheet = src.Workbook.Worksheets(sheetName)
      
      'Debug.Print "Opened " & src.Worksheet.name & " FROM " & src.Workbook.name
      
      pb.TextProgress = "Initializing source """ & sheetName & """"
      src.RevealHiddenData
      InitializeSourceWithSettingsAndConditions src
      
      'TODO may use hash(sheetname + workbookname) for the key value in sources collection
      'only in case we allow different source workbooks with same sheet names
      res.Add src, src.Worksheet.name
    End If
  Next lr
  
  Set GetSources = res
End Function

'---------------------------------------------------------------------------------------
' Procedure   : InitializeSourceWithSettingsAndConditions
' Description : obtains settings from the Report Settings sheet and initializes
'   WorksheetIO.HeadColumns with the retrieved settings for that specific source
'---------------------------------------------------------------------------------------
Private Sub InitializeSourceWithSettingsAndConditions( _
  ByRef src As WorksheetIO _
)
  Dim lr As ListRow
  Dim mapping As ListObject
  Set mapping = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(ReportFieldMappingTableName))
  Dim idx As Variant
  'Initialize HeadColumns from the FieldMapping Settings Table
  For Each lr In mapping.ListRows
    'match the source worksheet
    If StrComp( _
      mapping.DataBodyRange(lr.Index, mapping.ListColumns("Worksheet Name").Index).value, _
      src.Worksheet.name, _
      vbTextCompare _
    ) = 0 Then
      idx = mapping.DataBodyRange(lr.Index, mapping.ListColumns("ID").Index).value
      
      If Len(Nz(trim(idx))) = 0 Then
        err.raise Number:=ErrCode.NamedArgumentNotFound, _
          Description:="Missing field ID for Column Headers """ & _
          mapping.DataBodyRange(lr.Index, mapping.ListColumns("Column Headers").Index).value & """ " & _
          "in the Reports setings table """ & Setting(ReportFieldMappingTableName) & """"
      End If
      
      src.hcc.HeadColumns.Add New HeadColumn, CStr(idx)
      src.hcc.aliases.Add CStr(idx), CStr(idx)
      src.hcc.HeadColumn(CStr(idx)).Init _
        SearchDomain:=src.UsedRange, _
        HeadTitles:=mapping.DataBodyRange(lr.Index, mapping.ListColumns("Column Headers").Index).value, _
        dataType:=mapping.DataBodyRange(lr.Index, mapping.ListColumns("Field Type").Index).value, _
        NewHeaderValue:=mapping.DataBodyRange(lr.Index, mapping.ListColumns("New Column Header").Index).value
    End If
  Next lr
  
  'Initialize HeadColumns from the FieldConditions Settings Table
  Dim conditions As ListObject
  Set conditions = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(ReportFieldConditionsTableName))
  For Each lr In conditions.ListRows
    'match the source worksheet
    If StrComp( _
      conditions.DataBodyRange(lr.Index, conditions.ListColumns("Worksheet Name").Index).value, _
      src.Worksheet.name, _
      vbTextCompare _
    ) = 0 Then
      src.hcc.HeadColumns.Add New HeadColumn, "cond" & CStr(lr.Index)
      src.hcc.aliases.Add "cond" & CStr(lr.Index), "cond" & CStr(lr.Index)
      src.hcc.HeadColumn("cond" & CStr(lr.Index)).Init _
        SearchDomain:=src.UsedRange, _
        HeadTitles:=conditions.DataBodyRange(lr.Index, conditions.ListColumns("Column Headers").Index).value
    End If
  Next lr
  
  'report if not all headers where found in search domain
  Dim srcMissingHeaders As Variant
  srcMissingHeaders = src.hcc.GetTitlesOfHeadsNotFound()
  If Not Util.Arrays.IsEmpty(srcMissingHeaders) Then
    err.raise Number:=ErrCode.InvalidFileFormat, _
      Description:="Missing column headers """ & Util.Arrays.ToString(srcMissingHeaders) & _
      """ from worksheet " & _
      src.Worksheet.name & _
      """ in workbook " & _
      src.Workbook.name
  End If
  
  'If Util.Arrays.IsEmpty(srcMissingHeaders) Then Debug.Print "All heads found for " & src.Worksheet.name
End Sub

'Depricated
Public Function GetKeyHeadColumnsFromSources( _
  ByRef sources As Collection _
) As Collection
  Set GetKeyHeadColumnsFromSources = New Collection
  Dim hc As HeadColumn, src As WorksheetIO
  
  For Each src In sources
    For Each hc In src.hcc.HeadColumns
      If StrComp(hc.dataType, "Key", vbTextCompare) = 0 Then
        GetKeyHeadColumnsFromSources.Add hc
      End If
    Next hc
  Next src
End Function

Public Function GetUniquePrimaryKeys( _
  ByRef hcs As Collection _
) As Variant
  Dim hc As HeadColumn
  Dim al As New ArrayList
  
  pb.TextProgress = "Obtaining primary key values"
  
  For Each hc In hcs
    al.Append Application.Transpose(hc.AfterHeadRange.value)
  Next hc
  
  GetUniquePrimaryKeys = Util.Arrays.RemoveDuplicates(al.Values)
  pb.IncreseByPercent
End Function

Private Function GetDestination( _
  ByRef app As Application _
) As WorksheetIO
  Dim dst As New WorksheetIO
  dst.Application = app
  
  pb.TextProgress = "Initializing destination sheet"
  
  Dim wbPath As String
  wbPath = Util.File.GetFolderPath(ThisWorkbook.FullName) & "\" & _
    Util.DateTime.GetDateTimeStamp() & _
    "_Report_" & _
    Util.File.GetBaseName(ThisWorkbook.name) & ".xlsx"
  
  dst.Workbook = Util.Excel.CreateWorkbook(app, wbPath, Util.File.GetBaseName(wbPath))
  dst.Worksheet = Util.Excel.CreateWorksheet(dst.Workbook, "Report_" & Util.DateTime.GetDateTimeStamp())
  Util.Excel.CropWorksheet dst.Worksheet
  
  InitializeDestinationWithSettings dst
  
  Set GetDestination = dst
  pb.IncreseByPercent
End Function

Private Function IsKeyMapping( _
  ByRef lo As ListObject, _
  rowIdx As Long _
) As Boolean

  IsKeyMapping = (StrComp( _
      lo.DataBodyRange(rowIdx, lo.ListColumns("Field Type").Index), _
      "Key", _
      vbTextCompare) = 0)
End Function

Private Sub InitializeDestinationWithSettings( _
  ByRef dst As WorksheetIO _
)
  Dim mapping As ListObject, colOffset As Long, lr As ListRow
  Set mapping = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(ReportFieldMappingTableName))
  
  Dim rngStart As Range, rngHeader As Range, headerValue As Variant
  Set rngStart = dst.Worksheet.Cells(1, 1)
  colOffset = 0
  
  Dim idxNewHeader As Long, idxWSName As Long, idxHeaders As Long, idxId As Long, idx As String
  idxNewHeader = mapping.ListColumns("New Column Header").Index
  idxWSName = mapping.ListColumns("Worksheet Name").Index
  idxHeaders = mapping.ListColumns("Column Headers").Index
  idxId = mapping.ListColumns("ID").Index
  
  For Each lr In mapping.ListRows
    If IsKeyMapping(mapping, lr.Index) Then
      Set rngHeader = rngStart
    Else
      Set rngHeader = rngStart.offset(0, colOffset)
    End If
    
    'set header to either newly defined header or the first header option in source headers
    If Len(mapping.DataBodyRange(lr.Index, idxNewHeader).value) > 0 Then
      headerValue = mapping.DataBodyRange(lr.Index, idxNewHeader).value
    Else
      'default new name of column header in case none was provided
      Dim v As Variant
      v = Split(CStr(mapping.DataBodyRange(lr.Index, idxHeaders).value), ";")
      headerValue = v(LBound(v)) & "_" & CStr(mapping.DataBodyRange(lr.Index, idxWSName).value)
    End If
    
    rngHeader.value = headerValue
    idx = CStr(mapping.DataBodyRange(lr.Index, idxWSName).value) & CStr(mapping.DataBodyRange(lr.Index, idxId).value)
    
    dst.hcc.HeadColumns.Add New HeadColumn, idx
    dst.hcc.aliases.Add idx, idx
    dst.hcc.HeadColumn(idx).Init _
      SearchDomain:=rngHeader, _
      HeadTitles:=headerValue
      
    colOffset = colOffset + 1
  Next lr
  
  'add extra manual headers
  Set rngHeader = rngStart.offset(0, colOffset)
  rngHeader.value = "Duplicate in Sheet"
  dst.hcc.HeadColumns.Add New HeadColumn, "duplicate"
  dst.hcc.aliases.Add "duplicate", "duplicate"
  dst.hcc.HeadColumn("duplicate").Init _
    SearchDomain:=rngHeader, _
    HeadTitles:=rngHeader.value
End Sub

Public Function GetSheetKeyIndeces() As Collection
  Dim mapping As ListObject, lr As ListRow
On Error GoTo GetSheetKeyIndeces_Error
  
  Set GetSheetKeyIndeces = New Collection
  Set mapping = Util.Excel.GetListObjectByName(ThisWorkbook, Setting(ReportFieldMappingTableName))
  
  Dim idxType As Long, idxSheetName As Long, idxId As Long
  idxType = mapping.ListColumns("Field Type").Index
  idxSheetName = mapping.ListColumns("Worksheet Name").Index
  idxId = mapping.ListColumns("ID").Index
  
  Dim sheetKeys As New Collection
  For Each lr In mapping.ListRows
    If StrComp(mapping.DataBodyRange(lr.Index, idxType).value, "Key", vbTextCompare) = 0 Then
      sheetKeys.Add CStr(mapping.DataBodyRange(lr.Index, idxId).value), CStr(mapping.DataBodyRange(lr.Index, idxSheetName).value)
    End If
  Next lr
  
  Set GetSheetKeyIndeces = sheetKeys
  Exit Function
  
GetSheetKeyIndeces_Error:
  If err.Number = 457 Then 'duplicate collection key
    err.raise Number:=457, Description:="There can be only one uniquely defined field with Field Type = ""Key""" _
      & vbLf & "Please check the defined headers and field types in Report Settings Table """ & Setting(ReportFieldMappingTableName) & """"
  End If
End Function

Private Function GetSearchResultsForKeyInSources( _
  ByRef sources As Collection, _
  ByRef keyIndeces As Collection, _
  ByRef value As Variant _
) As Collection
  
  Set GetSearchResultsForKeyInSources = New Collection
  
  Dim src As WorksheetIO
  For Each src In sources
    GetSearchResultsForKeyInSources.Add _
      src.hcc.HeadColumn(keyIndeces(src.Worksheet.name)).FindAll(value), _
      src.Worksheet.name
  Next src
End Function

Private Function GetNumberOfNonEmptyCollections( _
  ByRef cols As Collection _
) As Integer
  
  Dim c As Collection, counter As Integer
  counter = 0
  For Each c In cols
    If c.count > 0 Then
      counter = counter + 1
    End If
  Next c

  GetNumberOfNonEmptyCollections = counter
End Function


Private Sub ImportSearchResults( _
  ByRef sources As Collection, _
  ByRef keyIndeces As Collection, _
  ByRef dst As WorksheetIO, _
  ByVal searchTerm As Variant, _
  ByRef conditions As Collection, _
  ByRef rowOffset As Long _
)
  Dim searchResults As Collection
  Set searchResults = GetSearchResultsForKeyInSources(sources, keyIndeces, searchTerm)
    
  Dim numSourcesWithMatches As Integer, src As WorksheetIO, nextRowFlag As Boolean
  numSourcesWithMatches = GetNumberOfNonEmptyCollections(searchResults)
  nextRowFlag = False
  
  'Debug.Print searchTerm
  
  If numSourcesWithMatches = 1 Then 'prim found in one sheet
    For Each src In sources
      nextRowFlag = nextRowFlag Or ImportMatchesFromSource(src, dst, searchResults(src.Worksheet.name), conditions, rowOffset)
    Next src
  ElseIf numSourcesWithMatches > 1 Then 'prim found in multiple sheets -> compare values to decide if import
    nextRowFlag = nextRowFlag Or ImportMultipleMatchesFromSources(sources, dst, searchResults, conditions, rowOffset)
  End If
  
  If nextRowFlag Then rowOffset = rowOffset + 1
End Sub

Private Function ImportMatchesFromSource( _
  ByRef src As WorksheetIO, _
  ByRef dst As WorksheetIO, _
  ByRef matches As Collection, _
  ByRef conditions As Collection, _
  ByRef rowOffset As Long _
) As Boolean
  ImportMatchesFromSource = False
  'Debug.Print matches.count
  If matches.count = 1 Then 'single value found in source
    ImportMatchesFromSource = ImportFromSource(src, dst, matches(1), conditions, rowOffset)
  ElseIf matches.count > 1 Then 'multiple values found in source
    'duplicates ImportFromSource
    ImportMatchesFromSource = ImportFromSource(src, dst, matches(1), conditions, rowOffset)
    'mark as duplicate
    dst.hcc.HeadColumn("duplicate").Head.offset(rowOffset, 0).value = src.Worksheet.name
  End If
End Function

Private Function ImportMultipleMatchesFromSources( _
  ByRef sources As Collection, _
  ByRef dst As WorksheetIO, _
  ByRef searchResults As Collection, _
  ByRef conditions As Collection, _
  ByRef rowOffset As Long _
) As Boolean
  ImportMultipleMatchesFromSources = False
  
  Dim i As Integer
  Dim importFlag As Boolean
  
  For i = 1 To searchResults.count
    If i < searchResults.count Then
      importFlag = importFlag Or FoundMatchesAreDifferent(sources, searchResults(i), searchResults(i + 1))
    End If
  Next i
  
  Dim src As WorksheetIO
  If importFlag Then
    For Each src In sources
      ImportMultipleMatchesFromSources = _
        ImportMultipleMatchesFromSources _
        Or ImportMatchesFromSource(src, dst, searchResults(src.Worksheet.name), conditions, rowOffset)
    Next src
  End If
End Function

Private Function FoundMatchesAreDifferent( _
  ByRef sources As Collection, _
  ByRef matchCol1 As Collection, _
  ByRef matchCol2 As Collection _
) As Boolean

  FoundMatchesAreDifferent = False
  Dim src1 As WorksheetIO, src2 As WorksheetIO, rng1 As Range, rng2 As Range
  
  For Each rng1 In matchCol1
    For Each rng2 In matchCol2
      Set src1 = GetSourceByWorksheetName(sources, rng1.Worksheet.name)
      Set src2 = GetSourceByWorksheetName(sources, rng2.Worksheet.name)
      FoundMatchesAreDifferent = FoundMatchesAreDifferent Or AreMatchesDifferent(src1, src2, rng1, rng2)
      If FoundMatchesAreDifferent Then Exit Function
    Next rng2
  Next rng1
End Function

Private Function AreMatchesDifferent( _
  ByRef src1 As WorksheetIO, _
  ByRef src2 As WorksheetIO, _
  ByRef rng1 As Range, _
  ByRef rng2 As Range _
) As Boolean
  AreMatchesDifferent = False
  
  Dim alias As Variant
  For Each alias In src1.hcc.aliases
    If Util.Strings.GetIndexOf(alias, "cond", vbTextCompare) = 0 Then
      If StrComp( _
        src1.Worksheet.Cells(rng1.row, src1.hcc.HeadColumn(alias).Head.Column).value, _
        src2.Worksheet.Cells(rng2.row, src2.hcc.HeadColumn(alias).Head.Column).value, _
        vbTextCompare _
      ) <> 0 Then
        AreMatchesDifferent = True
        Exit Function
      End If
    End If
  Next alias
  
End Function


Private Function GetSourceByWorksheetName( _
  ByRef sources As Collection, _
  ByVal sheetName As String _
) As WorksheetIO
  
  Dim src As WorksheetIO
  For Each src In sources
    If StrComp(src.Worksheet.name, sheetName, vbTextCompare) = 0 Then
      Set GetSourceByWorksheetName = src
      Exit Function
    End If
  Next src
End Function

Private Function ImportFromSource( _
  ByRef src As WorksheetIO, _
  ByRef dst As WorksheetIO, _
  ByRef rng As Range, _
  ByRef conditions As Collection, _
  ByRef rowOffset As Long _
) As Boolean
  ImportFromSource = False
  If AreFieldConditionsMet(rng, src, conditions) Then
    
    ImportFromSource = True
    
    Dim alias As Variant
    For Each alias In src.hcc.aliases
      If Util.Strings.GetIndexOf(alias, "cond", vbTextCompare) = 0 Then
        dst.hcc.HeadColumn(CStr(src.Worksheet.name) & CStr(alias)).Head.offset(rowOffset, 0).value = _
          src.Worksheet.Cells(rng.row, src.hcc.HeadColumn(alias).Head.Column).value
        
        FormatOnChosenDataType _
          dst.hcc.HeadColumn(CStr(src.Worksheet.name) & CStr(alias)).Head.offset(rowOffset, 0), _
          src.Worksheet.Cells(rng.row, src.hcc.HeadColumn(alias).Head.Column), _
          src.hcc.HeadColumn(alias).dataType
      End If
    Next alias
  End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckListObjectConsistency
' Author    : Ivanov, Bozhan
' Purpose   : Checks if a list object with that name exists in the provided workbook
'   and if the list column names are correctly defined
'   Raises appropriate error if the Check does not pass
'---------------------------------------------------------------------------------------
Private Sub CheckListObjectConsistency( _
  ByRef wb As Workbook, _
  ByVal loName As String, _
  ByVal columnNames As Variant _
)
  Dim lo As ListObject
  Set lo = Util.Excel.GetListObjectByName(wb, loName)
  
  If lo Is Nothing Then
    err.raise Number:=ErrCode.NamedArgumentNotFound, _
        Description:="No table with the name """ & _
        loName & """ was found!" & vbLf & _
        "Please Check your Settings!"
  End If
  
  Dim i As Integer, missingNames As New ArrayList
  For i = LBound(columnNames) To UBound(columnNames)
    If Util.Excel.GetListObjectColumnIndex(lo, columnNames(i)) <= 0 Then
      missingNames.Append columnNames(i)
    End If
  Next i
    
  If Not missingNames.IsEmpty Then
    Dim msg As String
    For i = missingNames.lb To missingNames.ub
      msg = msg & """" & missingNames(i) & """" & ", "
    Next i
    
    msg = Util.Strings.RemoveExcessSeperators(msg, ",")
    msg = "The column names " & msg & " in object list table """ & lo.name & """"
    msg = msg & " in worksheet """ & lo.Parent.name & """"
    msg = msg & " were not found!" & vbLf
    msg = msg & "Please Check the fields and column names definitions in """ & lo.name & """"
    err.raise Number:=ErrCode.NamedArgumentNotFound, Description:=msg
  End If
End Sub

Private Sub CheckWorkbookExistance( _
  ByVal path As Variant _
)
  If Not Util.File.IsWorkbookPath(path) _
  Or Not Util.File.FileExists(path) Then
    err.raise Number:=ErrCode.NamedArgumentNotFound, _
      Description:="The provided path is not a valid excel workbook or does not exist/you have no access permissions!" & _
        """" & path & """" & vbLf & _
        "Please, Check your settings." & vbLf & vbLf & _
        InfoErrMsg()
  End If
End Sub

Private Sub CheckWorksheetExistance( _
  ByRef wb As Workbook, _
  ByVal name As Variant _
)
  If Not Util.Excel.IsWorksheetName(wb, name) Then
    err.raise Number:=ErrCode.NamedArgumentNotFound, _
      Description:="Worksheet with the name """ & _
        name & """ was not found in workbook """ & _
        wb.name & """!" & vbLf & _
        "Please, Check your settings." & vbLf & vbLf & _
        InfoErrMsg()
  End If
End Sub

Private Function EvaluateFieldCondition( _
  value1 As Variant, _
  value2 As Variant, _
  cond As Variant _
) As Boolean
  EvaluateFieldCondition = False

  Select Case CStr(cond)
    Case "="
      If Util.Strings.CompareValues(value1, value2, vbTextCompare) = 0 Then
        EvaluateFieldCondition = True
        Exit Function
      End If
    Case ">"
      If Util.Strings.CompareValues(value1, value2, vbTextCompare) = 1 Then
        EvaluateFieldCondition = True
        Exit Function
      End If
    Case "<"
      If Util.Strings.CompareValues(value1, value2, vbTextCompare) = -1 Then
        EvaluateFieldCondition = True
        Exit Function
      End If
  End Select
End Function

Private Function AreFieldConditionsMet( _
  ByRef rng As Range, _
  ByRef src As WorksheetIO, _
  ByRef conds As Collection _
) As Boolean
'
'Dim fc As FieldCondition
'Dim result As Boolean
'result = True
'
'  For Each fc In conds
'    'check if cond are for that field condition
'    If StrComp(src.Worksheet.Name, fc.WorksheetName, vbTextCompare) = 0 Then
'      result = result _
'        And EvaluateFieldCondition( _
'          src.Worksheet.Cells(rng.row, src.hcc.HeadColumn(fc.alias).Head.Column).Value, _
'          fc.Value, _
'          fc.Operator)
'    End If
'  Next fc
'
'  AreFieldConditionsMet = result
End Function

'TODO move to a utility module
Public Function IsCollectionElement( _
  ByRef col As Collection, _
  ByRef key As Variant _
) As Boolean
  On Error Resume Next
  Call col(CStr(key))
  IsCollectionElement = (err.Number = 0)
  err.Clear
  On Error GoTo 0
End Function

Public Sub FormatOnChosenDataType( _
  ByRef dstRange As Range, _
  ByRef srcRange As Range, _
  ByVal dataType As Variant _
)
  dstRange.RowHeight = Setting(DefaultRowHeight)
  
  If Len(Nz(dataType)) > 0 Then
    Select Case dataType
    Case "Date"
      dstRange.NumberFormat = Setting(dateFormat)
    Case "Hyperlink"
      Util.Excel.CopyHyperlinks srcRange, dstRange
    End Select
  End If
End Sub

Public Sub FormatReportSheet( _
  ByRef dst As WorksheetIO _
)
  Dim lo As ListObject
  Set lo = dst.Worksheet.ListObjects.Add(xlSrcRange, dst.Worksheet.UsedRange, , xlYes)
  lo.name = "DiscrepancyReport"
  lo.TableStyle = "TableStyleMedium15"
  
  'Set source seperator columns
  Dim c As ListColumn, searchName As String, num As String
  searchName = "column"
  
  For Each c In lo.ListColumns
    c.Range.Columns.AutoFit
    
    If StrComp(searchName, Left(c.name, Len(searchName)), vbTextCompare) = 0 Then
      'check if after string "column" a number follows
      num = Util.Strings.GetRightOfIndex( _
        c.name, Util.Strings.GetIndexOf(c.name, searchName, vbTextCompare) + Len(searchName) - 1)
      If Util.Strings.IsNumber(num) Then
        c.Range.Interior.ColorIndex = ColorIndex.Orange
        c.Range.ColumnWidth = 2
      End If
    End If
  Next c
End Sub
