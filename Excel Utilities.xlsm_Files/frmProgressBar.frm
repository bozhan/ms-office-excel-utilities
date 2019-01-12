VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressBar 
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "frmProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cancel As Boolean

Private Sub cmdCancel_Click()
  cancel = True
End Sub

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub UserForm_Initialize()
  cancel = False
End Sub
