VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPhone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private intMaxLen As Integer
Private strMask As String
Private blnArea As Boolean
Private blnPars As Boolean
Private blnSpc As Boolean
Private blnDash As Boolean
Private blnExt  As Boolean

Public Property Get UseArea() As Boolean
    UseArea = blnArea
End Property

Public Property Let UseArea(ByVal Value As Boolean)
    blnArea = Value
End Property

Public Property Get UsePars() As Boolean
    UsePars = blnPars
End Property

Public Property Let UsePars(ByVal Value As Boolean)
    blnPars = Value
End Property

Public Property Get UseSpc() As Boolean
    UseSpc = blnSpc
End Property

Public Property Let UseSpc(ByVal Value As Boolean)
    blnSpc = Value
End Property

Public Property Get UseDash() As Boolean
    UseDash = blnDash
End Property

Public Property Let UseDash(ByVal Value As Boolean)
    blnDash = Value
End Property
Public Property Get UseExt() As Boolean
    UseExt = blnExt
End Property
Public Property Let UseExt(ByVal Value As Boolean)
    blnExt = Value
End Property

Public Property Get PhoneFormat() As String
    PhoneFormat = strMask
End Property

Public Property Let PhoneFormat(ByVal Value As String)
    strMask = Value
    intMaxLen = Len(strMask)
End Property

Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String
    Dim strTemp As String
    Dim intTemp As Integer
    
    If intKeyAscii <> vbKeyBack And (intKeyAscii < 48 Or intKeyAscii > 57) And (UCase(intKeyAscii) <> 88) Then
            Mask = strCurrentText
    ElseIf intKeyAscii <> vbKeyBack And (intKeyAscii < 48 Or intKeyAscii > 57) And (InStr(1, strCurrentText, "X") > 0) Then
            Mask = strCurrentText
    ElseIf intKeyAscii = vbKeyBack Then
        If Len(strCurrentText) > 1 Then
            If (Asc(Right(strCurrentText, 1)) < 48) Or (Asc(Right(strCurrentText, 1)) > 57) Then
                strTemp = Left(strCurrentText, Len(strCurrentText) - 2)
                Do While ((Asc(Right(strTemp, 1)) < 48) Or (Asc(Right(strTemp, 1)) > 57))
                    strTemp = Left(strCurrentText, Len(strCurrentText) - 3)
                Loop
            Else
                strTemp = Left(strCurrentText, Len(strCurrentText) - 1)
            End If
        End If
        Mask = strTemp
    Else
        If Len(strCurrentText) > intMaxLen - 1 Then
            Mask = strCurrentText
        Else
            intTemp = (intMaxLen - Len(strCurrentText)) - 2
            If intTemp < 0 Then intTemp = 0
            strTemp = strCurrentText & "LM" & String(intTemp, "0")
            strTemp = Replace(strTemp, "(", "")
            strTemp = Replace(strTemp, ")", "")
            strTemp = Replace(strTemp, " ", "")
            strTemp = Replace(strTemp, "-", "")
            strTemp = Replace(strTemp, "X", "")
            strTemp = Format(strTemp, strMask)
            strTemp = Left(strTemp, InStr(1, strTemp, "M") - 1)
            strTemp = Replace(strTemp, "L", Chr(intKeyAscii))
            Mask = strTemp
        End If
    End If
    intKeyAscii = 0
End Function

