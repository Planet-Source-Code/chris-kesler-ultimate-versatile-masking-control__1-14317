VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum EmailMaskType
    Standard = 0
    DecimalB4AtSign = 1
    MultiDecimalAfterAt = 2
    DecimalB4MultiAfterAt = 3
End Enum

Private strMask As String
Private intMaxLen As Integer
Private intEmailMask As Integer

Public Property Get EmailMask() As EmailMaskType
    EmailMask = intEmailMask
End Property

Public Property Let EmailMask(ByVal Value As EmailMaskType)
    intEmailMask = Value
End Property

Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String
    Dim intRev As Integer
    
    If intKeyAscii = vbKeyBack Then
        If Len(strCurrentText) > 1 Then
            Mask = Left(strCurrentText, Len(strCurrentText) - 1)
        End If
    ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) And ((Asc(UCase(Chr(intKeyAscii))) < 65) Or (Asc(UCase(Chr(intKeyAscii))) > 90)) And intKeyAscii <> 46 And intKeyAscii <> 64 And intKeyAscii <> 95 Then
        Mask = strCurrentText
    Else
        If EmailMask = 0 Then
            If (intKeyAscii = 46) And (InStr(1, strCurrentText, "@") = 0) Then
                Mask = strCurrentText
            ElseIf (intKeyAscii = 64) And (InStr(1, strCurrentText, "@") > 0) Then
                Mask = strCurrentText
            ElseIf (intKeyAscii = 46) And (InStr(1, strCurrentText, ".") > 0) Then
                Mask = strCurrentText
            Else
                Mask = strCurrentText & Chr(intKeyAscii)
            End If
        ElseIf EmailMask = 1 Then
            If (intKeyAscii = 46) And (InStr(1, strCurrentText, ".") = 0) And (InStr(1, strCurrentText, "@") = 0) Then
                Mask = strCurrentText & Chr(intKeyAscii)
            ElseIf (intKeyAscii = 46) And (InStr(1, strCurrentText, ".") > 0) And (InStr(1, strCurrentText, "@") = 0) Then
                Mask = strCurrentText
            ElseIf (intKeyAscii = 64) And (InStr(1, strCurrentText, "@") = 0) Then
                Mask = strCurrentText & Chr(intKeyAscii)
            ElseIf (intKeyAscii = 64) And (InStr(1, strCurrentText, "@") > 0) Then
                Mask = strCurrentText
            ElseIf (intKeyAscii = 46) And (InStr(1, strCurrentText, ".") > 0) And (InStr(1, strCurrentText, "@") > 0) Then
                If InStrRev(strCurrentText, ".", Len(strCurrentText) - 1) < InStr(1, strCurrentText, "@") Then
                    Mask = strCurrentText & Chr(intKeyAscii)
                Else
                    Mask = strCurrentText
                End If
            Else
                Mask = strCurrentText & Chr(intKeyAscii)
            End If
        ElseIf EmailMask = 2 Then
            If (intKeyAscii = 46) And (InStr(1, strCurrentText, "@") = 0) Then
                Mask = strCurrentText
            ElseIf (intKeyAscii = 64) And (InStr(1, strCurrentText, "@") > 0) Then
                Mask = strCurrentText
            Else
                Mask = strCurrentText & Chr(intKeyAscii)
            End If
        ElseIf EmailMask = 3 Then
            If (intKeyAscii = 46) And (InStr(1, strCurrentText, ".") = 0) And (InStr(1, strCurrentText, "@") = 0) Then
                Mask = strCurrentText & Chr(intKeyAscii)
            ElseIf (intKeyAscii = 46) And (InStr(1, strCurrentText, ".") > 0) And (InStr(1, strCurrentText, "@") = 0) Then
                Mask = strCurrentText
            ElseIf (intKeyAscii = 64) And (InStr(1, strCurrentText, "@") = 0) Then
                Mask = strCurrentText & Chr(intKeyAscii)
            ElseIf (intKeyAscii = 64) And (InStr(1, strCurrentText, "@") > 0) Then
                Mask = strCurrentText
            ElseIf (intKeyAscii = 46) And (InStr(1, strCurrentText, ".") > 0) And (InStr(1, strCurrentText, "@") > 0) Then
                Mask = strCurrentText & Chr(intKeyAscii)
            Else
                Mask = strCurrentText & Chr(intKeyAscii)
            End If
        End If
    End If
    intKeyAscii = 0
End Function

