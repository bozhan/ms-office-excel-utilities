VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBarController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    :
' Author    : Ivanov, Bozhan
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit
Private progressBuffer As Single
Private frm As frmProgressBar
Attribute frm.VB_VarHelpID = -1
Private cmdCancel As MSforms.Control
Attribute cmdCancel.VB_VarHelpID = -1
Private bCancel As Boolean

Private Sub Class_Initialize()
  Set frm = VBA.UserForms.Add("frmProgressBar")
  With frm
    setControlGeoAndCaption .lblStatus, 6, 18, 6, 220
    setControlGeoAndCaption .ProgressFrame, .lblStatus.Top + .lblStatus.Height + 2, 20, 8, 200
    setControlGeoAndCaption .ProgressLabel, .ProgressFrame.Top + 4, 12, 210, 24
    setControlGeoAndCaption .ProgressBar, 0, .ProgressFrame.Height, 0, 0
    .ProgressBar.BackColor = ColorDec.Blue_Highlight
    setControlGeoAndCaption .lblResult, .ProgressFrame.Top + .ProgressFrame.Height + 8, 18, 8, 220
    .lblResult.AutoSize = True
    setControlGeoAndCaption .cmdOk, .lblResult.Top + .lblResult.Height + 12, 24, (.Width / 2) - (.cmdOk.Width / 2), .cmdOk.Width, "OK"
    setControlGeoAndCaption .cmdCancel, .cmdOk.Top, 24, .cmdOk.Left + .cmdOk.Width + 20, 48, "Cancel"
    setControlGeoAndCaption Me.Form, 0, .ProgressFrame.Top + .ProgressFrame.Height + 28, 0, 240
    .img_Alert.Visible = False
    .img_Success.Visible = False
    .cmdOk.Visible = False
    .cmdCancel.Visible = False
    .lblResult.Visible = False
  End With
  progressBuffer = 0
  bCancel = False
  HideResultControls
End Sub

Private Sub setControlGeoAndCaption( _
ByRef crl As Object, _
ByVal Top As Single, _
ByVal Height As Single, _
ByVal Left As Single, _
ByVal Width As Single, _
Optional ByVal Caption As String = vbNullString _
)
On Error Resume Next
  With crl
    .Top = Top
    .Height = Height
    .Left = Left
    .Width = Width
    If Len(Caption) > 0 Then
      .Caption = Caption
    End If
  End With
On Error GoTo 0
End Sub

Private Sub Class_Terminate()
  Set frm = Nothing
End Sub

Public Property Get IsCancelable() As Boolean
  IsCancelable = bCancel
End Property

Public Property Let IsCancelable(ByVal value As Boolean)
  bCancel = value
  If bCancel Then
    ShowCancelControl
  Else
    HideResultControls
  End If
End Property

Private Sub ShowCancelControl()
  With Me.Form
    setControlGeoAndCaption .cmdCancel, .ProgressFrame.Top + .ProgressFrame.Height + 12, 24, .ProgressFrame.Left + .ProgressFrame.Width - 48, 48
    .Height = .cmdCancel.Top + .cmdCancel.Height + 27
    .cmdCancel.Visible = True
  End With
End Sub

Public Property Get Form() As frmProgressBar
  Set Form = frm
End Property

Public Property Get Caption() As Variant
  Caption = Me.Form.Caption
End Property

Public Property Let Caption(ByVal value As Variant)
  Me.Form.Caption = CStr(value)
End Property

Public Property Get ProgressMax() As Single
  ProgressMax = Me.Form.ProgressFrame.Width
End Property

Public Property Get ProgressLeft() As Single
  ProgressLeft = Me.ProgressMax - Me.Progress
End Property

Public Property Get Progress() As Single
  Progress = Me.Form.ProgressBar.Width
End Property

'TODO set progress to the value not incrementally incerasing with the value
'impl another sub that will increaseprogress
Public Property Let Progress(ByVal value As Single)
  If value - Me.Form.ProgressBar.Width < 0.1 Then
    progressBuffer = progressBuffer + value - Me.Form.ProgressBar.Width
  Else
    Me.Form.ProgressBar.Width = value
  End If
  
  If progressBuffer >= 1 Then
    Me.Form.ProgressBar.Width = Me.Form.ProgressBar.Width + progressBuffer
    progressBuffer = 0
  End If
  
  Me.Form.ProgressLabel.Caption = Format((Me.Progress / Me.ProgressMax) * 100, "0") & "%"
  DoEvents
End Property

Public Property Get TextProgress() As Variant
  TextProgress = Me.Form.lblStatus.Caption
End Property

Public Property Let TextProgress(ByVal value As Variant)
  Me.Form.lblStatus.Caption = CStr(value)
  Me.Form.Repaint
End Property

Public Property Get TextResult() As Variant
  TextResult = Me.Form.lblResult.Caption
End Property

Public Property Let TextResult(ByVal value As Variant)
  Me.Form.lblResult.Caption = CStr(value)
End Property

Public Property Get OkButton() As Control
  Set OkButton = Me.Form.cmdOk
End Property

Public Property Get CancelButton() As Control
  Set CancelButton = Me.Form.cmdCancel
End Property

Public Property Get AlertIcon() As Control
  Set AlertIcon = Me.Form.img_Alert
End Property

Public Property Get SuccessIcon() As Control
  Set SuccessIcon = Me.Form.img_Success
End Property

Public Property Get IsCanceled() As Boolean
  IsCanceled = Me.Form.cancel
End Property

Public Sub CloseForm()
  Unload Me.Form
End Sub

Public Sub HideResultControls()
  With Me.Form
    .cmdOk.Visible = False
    .cmdCancel.Visible = False
    .lblResult.Visible = False
    .Height = .ProgressFrame.Top + .ProgressFrame.Height + 28
    .Repaint
  End With
End Sub

Public Sub ShowResultControls()
  With Me.Form
    setControlGeoAndCaption .cmdOk, .lblResult.Top + .lblResult.Height + 12, 24, (.Width / 2) - (.cmdOk.Width / 2), .cmdOk.Width
    setControlGeoAndCaption .cmdCancel, .cmdOk.Top, 24, .cmdOk.Left + .cmdOk.Width + 20, 48
    .Height = .cmdOk.Top + .cmdOk.Height + 27
    .cmdCancel.Visible = False
    .cmdOk.Visible = True
    .lblResult.Visible = True
    .Repaint
    .cmdOk.SetFocus
  End With
  
  If Not Me.Form.Visible Then Me.Show
End Sub

Public Sub IncreseByPercent(Optional ByVal value As Integer = 1)
  Me.Progress = Me.Progress + ((Me.ProgressMax) * (value / 100))
End Sub

Public Sub Show(Optional ByVal Modal As Boolean = False)
  If Modal Then
    Me.Form.Show vbModal
  Else
    Me.Form.Show vbModeless
  End If
  
  If Me.IsCancelable Then
    With Me.Form
      setControlGeoAndCaption .cmdCancel, .ProgressFrame.Top + .ProgressFrame.Height + 12, 24, .ProgressFrame.Left + .ProgressFrame.Width - 48, 48
      .Height = .cmdCancel.Top + .cmdCancel.Height + 27
      .cmdCancel.Visible = True
    End With
  End If
  
End Sub

Public Sub Hide()
  Me.Form.Hide
End Sub



