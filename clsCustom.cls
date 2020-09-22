VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strCharAllowed As String
Private blnAlpha As Boolean
Private blnNumeric As Boolean
Private blnAllCaps As Boolean
Private intMaxAllowed As Integer

Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String
    If intKeyAscii = vbKeyBack Then
        If Len(strCurrentText) > 1 Then Mask = Left(strCurrentText, Len(strCurrentText) - 1)
    ElseIf (Len(strCurrentText) > intMaxAllowed - 1) And (intMaxAllowed <> 0) Then
        Mask = strCurrentText
    ElseIf (Not IsAlpha(intKeyAscii)) And (Not (IsNumber(intKeyAscii))) And (Not (IsAllowed(intKeyAscii))) Then
        Mask = strCurrentText
    ElseIf (IsAllowed(intKeyAscii)) Then
        If AllowAlpha And (IsAlpha(intKeyAscii)) Then
            If AllCaps Then
                Mask = UCase(strCurrentText & Chr(intKeyAscii))
            Else
                Mask = strCurrentText & Chr(intKeyAscii)
            End If
        ElseIf AllowNumeric And (IsNumber(intKeyAscii)) Then
            Mask = strCurrentText & Chr(intKeyAscii)
        Else
            Mask = strCurrentText
        End If
    ElseIf CharAllowed = "" Then
        If AllowAlpha And (IsAlpha(intKeyAscii)) Then
            If AllCaps Then
                Mask = UCase(strCurrentText & Chr(intKeyAscii))
            Else
                Mask = strCurrentText & Chr(intKeyAscii)
            End If
        ElseIf AllowNumeric And (IsNumber(intKeyAscii)) Then
            Mask = strCurrentText & Chr(intKeyAscii)
        Else
            Mask = strCurrentText
        End If
    Else
        Mask = strCurrentText
    End If
    intKeyAscii = 0
End Function

Private Function IsAlpha(intAscii As Integer) As Boolean
    Dim intTemp As Integer
    
    intTemp = Asc(UCase(Chr(intAscii)))
    IsAlpha = (blnAlpha And Not (intTemp < 65 Or intTemp > 90))
End Function

Private Function IsNumber(intAscii As Integer) As Boolean
    IsNumber = blnNumeric And (intAscii > 47 And intAscii < 58)
End Function

Private Function IsAllowed(intAscii) As Boolean
    IsAllowed = (InStr(1, strCharAllowed, Chr(intAscii)) > 0)
End Function

Public Property Get AllowAlpha() As Boolean
    AllowAlpha = blnAlpha
End Property

Public Property Let AllowAlpha(Value As Boolean)
    blnAlpha = Value
End Property

Public Property Get AllowNumeric() As Boolean
    AllowNumeric = blnNumeric
End Property

Public Property Let AllowNumeric(Value As Boolean)
    blnNumeric = Value
End Property

Public Property Get CharAllowed() As String
    CharAllowed = strCharAllowed
End Property

Public Property Let CharAllowed(Value As String)
    strCharAllowed = Value
End Property

Public Property Get MaxAllowed() As Integer
    MaxAllowed = intMaxAllowed
End Property

Public Property Let MaxAllowed(Value As Integer)
    intMaxAllowed = Value
End Property

Public Property Get AllCaps() As Boolean
    AllCaps = blnAllCaps
End Property
Public Property Let AllCaps(ByVal bAllCaps As Boolean)
    blnAllCaps = bAllCaps
End Property
