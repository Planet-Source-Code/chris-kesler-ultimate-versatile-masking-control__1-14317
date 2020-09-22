VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum UseSignType
    SetDollarSign = True
    UnSetDollarSign = False
End Enum

Private blnSign As Boolean

Public Property Get UseSign() As UseSignType
    UseSign = blnSign
End Property

Public Property Let UseSign(ByVal Value As UseSignType)
    blnSign = Value
End Property

Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String
    If intKeyAscii = vbKeyBack Then
        If Len(strCurrentText) > (1 + Abs(blnSign) + InStr(1, strCurrentText, "-")) Then
            Mask = Left(strCurrentText, Len(strCurrentText) - 1)
        End If
    ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) And intKeyAscii <> 46 And intKeyAscii <> 45 Then
        Mask = strCurrentText
    Else
        If intKeyAscii = 45 Then
            If InStr(1, strCurrentText, "-") > 0 Then 'remove - sign since it already exists
                Mask = Mid(strCurrentText, 2)
            Else
                Mask = Chr(intKeyAscii) & strCurrentText
            End If
        ElseIf intKeyAscii = 46 Then    'decimal
            If Len(strCurrentText) = (0 + (Abs(Left(strCurrentText, 1) = "-"))) Then
                If blnSign Then
                    Mask = strCurrentText & "$ 0."
                Else
                    Mask = strCurrentText & "0."
                End If
            ElseIf InStr(1, strCurrentText, ".") = 0 Then
                Mask = strCurrentText & "."
            Else
                Mask = strCurrentText
            End If
        Else    'number
            If Len(strCurrentText) = (0 + (Abs(Left(strCurrentText, 1) = "-"))) Then
                If blnSign Then
                    Mask = strCurrentText & "$ " & Chr(intKeyAscii)
                Else
                    Mask = strCurrentText & Chr(intKeyAscii)
                End If
            ElseIf InStr(1, strCurrentText, ".") = 0 Then
                Mask = strCurrentText & Chr(intKeyAscii)
            Else
                If Len(strCurrentText) - InStr(1, strCurrentText, ".") = 2 Then
                    Mask = strCurrentText
                Else
                    Mask = strCurrentText & Chr(intKeyAscii)
                End If
            End If
        End If
    End If
    intKeyAscii = 0
End Function
