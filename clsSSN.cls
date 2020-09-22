VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum SSNMaskType
    SetSSNMask = True
    UnSetSSNMask = False
End Enum

Private strMask As String
Private intMaxLen As Integer
Private blnDashes As Boolean

Public Property Get SSNMask() As SSNMaskType
    SSNMask = blnDashes
End Property

Public Property Let SSNMask(ByVal Value As SSNMaskType)
    blnDashes = Value
    If Not blnDashes Then
        strMask = "&&&&&&&&&"
    Else
        strMask = "&&&-&&-&&&&"
    End If
    intMaxLen = Len(strMask)
End Property

Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String
    Dim strTemp As String
    Dim intTemp As Integer
    
    If intKeyAscii <> vbKeyBack And (intKeyAscii < 48 Or intKeyAscii > 57) Then
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
            strTemp = strCurrentText & "XY" & String(intTemp, "0")
            strTemp = Replace(strTemp, "-", "")
            strTemp = Format(strTemp, strMask)
            strTemp = Left(strTemp, InStr(1, strTemp, "Y") - 1)
            strTemp = Replace(strTemp, "X", Chr(intKeyAscii))
            Mask = strTemp
        End If
    End If
    intKeyAscii = 0
End Function
