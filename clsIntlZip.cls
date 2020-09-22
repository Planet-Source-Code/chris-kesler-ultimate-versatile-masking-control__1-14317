VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIntlZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum IntlZipType
    SetIntlZipMask = True
    UnSetIntlZipMask = False
End Enum

Private strMask As String
Private intMaxLen As Integer
Private blnSpaces As Boolean

Public Property Get IntlZipMask() As IntlZipType
    IntlZipMask = blnSpaces
End Property

Public Property Let IntlZipMask(ByVal Value As IntlZipType)
    blnSpaces = Value
    If Not blnSpaces Then
        strMask = "&&&&&&"
    Else
        strMask = "&&& &&&"
    End If
    intMaxLen = Len(strMask)
End Property

Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String
    Dim strTemp As String
    Dim intTemp As Integer
    
    If intKeyAscii <> vbKeyBack And (intKeyAscii < 48 Or intKeyAscii > 57) And _
        (intKeyAscii < 65 Or intKeyAscii > 90) And (intKeyAscii < 97 Or intKeyAscii > 122) Then
        Mask = UCase(strCurrentText)
    ElseIf intKeyAscii = vbKeyBack Then
        If Len(strCurrentText) > 1 Then
            If Asc(Right(strCurrentText, 1)) = 0 Then
                strTemp = Left(strCurrentText, Len(strCurrentText) - 2)
            Else
                strTemp = Left(strCurrentText, Len(strCurrentText) - 1)
            End If
        End If
        Mask = UCase(strTemp)
    Else
        If Len(strCurrentText) > intMaxLen - 1 Then
            Mask = UCase(strCurrentText)
        Else
            intTemp = (intMaxLen - Len(strCurrentText)) - 2
            If intTemp < 0 Then intTemp = 0
            strTemp = strCurrentText & "XY" & String(intTemp, "0")
            strTemp = Replace(strTemp, " ", "")
            strTemp = Format(strTemp, strMask)
            strTemp = Left(strTemp, InStr(1, strTemp, "Y") - 1)
            strTemp = Replace(strTemp, "X", Chr(intKeyAscii))
            Mask = UCase(strTemp)
        End If
    End If
    intKeyAscii = 0
End Function

