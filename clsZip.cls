VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum ZipMaskType
    SetOnlyFive = True
    SetZipPlus4 = False
End Enum

Private blnOnlyFive As Boolean

Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String
    If (intKeyAscii < 48 Or intKeyAscii > 57) And (intKeyAscii <> vbKeyBack) And (intKeyAscii <> 45) Then
        Mask = strCurrentText
    ElseIf Len(Trim(strCurrentText)) < 5 Then
        If intKeyAscii = vbKeyBack Then
            If Len(Trim(strCurrentText)) > 1 Then Mask = Left(strCurrentText, Len(strCurrentText) - 1)
        ElseIf intKeyAscii <> 45 Then
            Mask = strCurrentText & Chr(intKeyAscii)
        Else
            Mask = strCurrentText
        End If
    Else
        If blnOnlyFive And intKeyAscii <> vbKeyBack Then
            Mask = strCurrentText
        ElseIf intKeyAscii = vbKeyBack Then
            Mask = Left(strCurrentText, Len(strCurrentText) - 1)
        ElseIf Len(strCurrentText) = 5 Then
            If intKeyAscii = 45 Then
                Mask = strCurrentText & "-"
            Else
                Mask = strCurrentText & "-" & Chr(intKeyAscii)
            End If
        ElseIf (Len(strCurrentText) = 10) And intKeyAscii <> vbKeyBack Then
            Mask = strCurrentText
        Else
            If intKeyAscii = 45 Then
                Mask = strCurrentText
            Else
                Mask = strCurrentText & Chr(intKeyAscii)
            End If
        End If
    End If
    intKeyAscii = 0
End Function

Public Property Get OnlyFive() As ZipMaskType
    OnlyFive = blnOnlyFive
End Property

Public Property Let OnlyFive(ByVal Value As ZipMaskType)
    blnOnlyFive = Value
End Property
