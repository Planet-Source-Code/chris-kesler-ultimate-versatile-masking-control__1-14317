VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWholeNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum UsePrecision
    SetPrecision = True
    UnSetPrecision = False
End Enum
Public Enum AllowNegatives
    SetNegative = True
    UnSetNegative = False
End Enum
Private intPrecision As Integer
Private blnAllowNegative As Boolean
Private blnPrecision As Boolean

Public Property Get UsePrecision() As UsePrecision
    UsePrecision = blnPrecision
End Property

Public Property Let UsePrecision(ByVal Value As UsePrecision)
    blnPrecision = Value
End Property

Public Property Get PrecisionValue() As Integer
    PrecisionValue = intPrecision
End Property

Public Property Let PrecisionValue(ByVal Value As Integer)
    intPrecision = Value
End Property

Public Property Get AllowNegatives() As AllowNegatives
    AllowNegatives = blnAllowNegative
End Property

Public Property Let AllowNegatives(ByVal Value As AllowNegatives)
    blnAllowNegative = Value
End Property
Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String
    If intKeyAscii = vbKeyBack Then
        If Len(strCurrentText) > (1 + Abs(UsePrecision) + InStr(1, strCurrentText, "-")) Then
            Mask = Left(strCurrentText, Len(strCurrentText) - 1)
        End If
    ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) And intKeyAscii <> 46 And intKeyAscii <> 45 Then
        Mask = strCurrentText
    Else
        If intKeyAscii = 45 Then
            If AllowNegatives = UnSetNegative Then
                Mask = strCurrentText
            ElseIf InStr(1, strCurrentText, "-") > 0 Then 'remove - sign since it already exists
                Mask = Mid(strCurrentText, 2)
            Else
                Mask = Chr(intKeyAscii) & strCurrentText
            End If
        ElseIf intKeyAscii = 46 Then    'decimal
            If PrecisionValue = 0 Then
                Mask = strCurrentText
            ElseIf InStr(1, strCurrentText, ".") = 0 Then
                Mask = strCurrentText & "."
            Else
                Mask = strCurrentText
            End If
        Else    'number
            If InStr(1, strCurrentText, ".") = 0 Then
                Mask = strCurrentText & Chr(intKeyAscii)
            Else
                If Len(strCurrentText) - InStr(1, strCurrentText, ".") = PrecisionValue Then
                    Mask = strCurrentText
                Else
                    Mask = strCurrentText & Chr(intKeyAscii)
                End If
            End If
        End If
    End If
    intKeyAscii = 0
End Function

