VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum MagnitudeOrder
  tera = 12
  giga = 9
  mega = 6
  kilo = 3
  hecto = 2
  deca = 1
  one = 0
  deci = -1
  centi = -2
  milli = -3
  micro = -6
  nano = -9
  pico = -12
End Enum

Private Const TPI As Single = 1440 'twips per inch
Private Const PPI As Single = 72 'pixels per inch
Private Const IPC As Single = 0.3937 'Inch per cm
Private Const TPC As Single = 567 'twips per cm
Private Const TPP As Single = 20 'twips per point for font uses

Property Get TwipsPerPoint() As Single
  TwipsPerPoint = TPP
End Property

Property Get TwipsPerInch() As Single
  TwipsPerInch = TPI
End Property

Property Get PixelsPerInch() As Single
  PixelsPerInch = PPI
End Property

Property Get InchesPerCm() As Single
  InchesPerCm = IPC
End Property

Property Get TwipsPerCm() As Single
  TwipsPerCm = TPC
End Property

Public Function TwipsToInch(value As Single) As Single
    TwipsToInch = value / TwipsPerInch
End Function

Public Function InchToTwips(value As Single) As Single
    InchToTwips = value * TwipsPerInch
End Function

Public Function PixelsToInch(value As Single) As Single
    PixelsToInch = value / PixelsPerInch
End Function

Public Function InchToPixels(value As Single) As Single
    InchToPixels = value * PixelsPerInch
End Function

Public Function TwipsToCm(value As Single) As Single
    TwipsToCm = value * TwipsPerCm
End Function

Public Function CmToTwips(value As Single) As Single
    CmToTwips = value / TwipsPerCm
End Function

Public Function PixelsToCm(value As Single) As Single
    PixelsToCm = value / (PixelsPerInch / InchesPerCm)
End Function

Public Function CmToPixels(value As Single) As Single
  If IsNumeric(value) Then
    CmToPixels = (value * InchesPerCm) * PixelsPerInch
  End If
End Function

Public Function SingleToString(ByVal value As Single) As String
  SingleToString = Replace(CStr(value), ",", ".")
End Function

Public Function StringToSingle(ByVal value As String) As Single
  StringToSingle = CSng(Replace(value, ".", ","))
End Function
