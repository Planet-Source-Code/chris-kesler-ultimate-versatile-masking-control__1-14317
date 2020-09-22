VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum DateMaskType
    ShortDate = 0
    MedDate = 1
    LongDate = 2
End Enum

Private intDateTypeMask As Integer
Private strMonthFormat(2) As String
Private strDayFormat(2) As String

Public Property Get DateTypeMask() As DateMaskType
    DateTypeMask = intDateTypeMask
End Property

Public Property Let DateTypeMask(ByVal Value As DateMaskType)
    intDateTypeMask = Value
End Property
Private Sub Class_Initialize()
    strMonthFormat(0) = "MM/"
    strMonthFormat(1) = "MMM "
    strMonthFormat(2) = "MMMM "
    strDayFormat(0) = "DD/"
    strDayFormat(1) = "DD, "
    strDayFormat(2) = "DD, "
End Sub

Public Function Mask(intKeyAscii As Integer, strCurrentText As String) As String
    Dim intTemp As Integer
    Select Case Len(strCurrentText)
        Case 0, 1, 2
            If intKeyAscii = vbKeyBack Then
                Mask = ""
            Else
                Mask = FinishMonth(strCurrentText & Chr(intKeyAscii))
            End If
        Case 3  'should only be here with short dates
            If intDateTypeMask > 0 Then 'an error occured erase everything and start over
                Mask = ""
            ElseIf intKeyAscii = vbKeyBack Then
                Mask = ""
            ElseIf intKeyAscii < 48 Or intKeyAscii > 57 Then
                Mask = strCurrentText
            Else
                Mask = FinishDay((Val(Left(strCurrentText, 2))), Val(Chr(intKeyAscii)))
            End If
        Case 4
            If intDateTypeMask = 0 Then 'short date
                intTemp = Val(Left(strCurrentText, 2))
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 3)
                ElseIf intKeyAscii < 48 Or intKeyAscii > 57 Then
                    Mask = strCurrentText
                Else
                    Mask = FinishDay(intTemp, Val(Mid(strCurrentText, 4, 1)), Val(Chr(intKeyAscii)))
                End If
            Else    'med date or long and month of May
                intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
                If (intDateTypeMask = 2) And intTemp <> 5 Then  'an error occured erase everything and start over
                    Mask = ""
                ElseIf intKeyAscii = vbKeyBack Then
                    Mask = ""
                ElseIf intKeyAscii < 48 Or intKeyAscii > 57 Then
                    Mask = strCurrentText
                Else
                    Mask = FinishDay(intTemp, Val(Chr(intKeyAscii)))
                End If
            End If
        Case 5  'should only be here for med and long dates
            If intDateTypeMask = 0 Then 'an error occured erase everything and start over
                Mask = ""
            ElseIf intDateTypeMask = 1 Then 'med date
                intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 4)
                ElseIf intKeyAscii < 48 Or intKeyAscii > 57 Then
                    Mask = strCurrentText
                Else
                    Mask = FinishDay(intTemp, Val(Mid(strCurrentText, 5, 1)), Val(Chr(intKeyAscii)))
                End If
            Else    'long date
                intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
                If intTemp <> 5 And intTemp <> 6 And intTemp <> 7 Then  'an error has occured erase it all and start over
                    Mask = ""
                ElseIf intKeyAscii = vbKeyBack Then
                    If intTemp = 5 Then
                        Mask = Left(strCurrentText, 4)
                    Else
                        Mask = ""
                    End If
                ElseIf intKeyAscii < 48 Or intKeyAscii > 57 Then
                    Mask = strCurrentText
                ElseIf intTemp = 5 Then 'long month of may
                    Mask = FinishDay(5, Val(Mid(strCurrentText, 5, 1)), Val(Chr(intKeyAscii)))
                ElseIf (intTemp = 6) Or (intTemp = 7) Then 'long month of june and july
                    Mask = FinishDay(intTemp, Val(Chr(intKeyAscii)))
                Else
                    Mask = strCurrentText
                End If
            End If
        Case 6  'six characters exist
            If intDateTypeMask = 1 Then 'an error occured erase everything and start over
                Mask = ""
            ElseIf intDateTypeMask = 0 Then 'Short Date
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 3)
                ElseIf (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    Mask = strCurrentText & Chr(intKeyAscii)
                End If
            Else    'long date
                intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
                If intKeyAscii = vbKeyBack Then
                    If intTemp = 6 Or intTemp = 7 Then
                        Mask = Left(strCurrentText, 5)
                    Else    'an error occured erase everything and start over
                        'also erase everything if it is march or april
                        Mask = ""
                    End If
                ElseIf (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    If intTemp = 3 Or intTemp = 4 Then  'mar or apr
                        Mask = FinishDay(intTemp, Val(Chr(intKeyAscii)))
                    ElseIf intTemp = 6 Or intTemp = 7 Then  'jun or jul
                        Mask = FinishDay(intTemp, Val(Mid(strCurrentText, 6, 1)), Val(Chr(intKeyAscii)))
                    Else    'an error occured erase everything and start over
                        Mask = ""
                    End If
                End If
            End If
        Case 7  'seven characters exist
            If intDateTypeMask = 1 Then 'an error occured erase everything and start over
                Mask = ""
            ElseIf intDateTypeMask = 0 Then 'short date
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 6)
                ElseIf (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    Mask = FinishYear(Val(Left(strCurrentText, 2)), Val(Mid(strCurrentText, 4, 2)), Val(Mid(strCurrentText, 7, 1)), Val(Chr(intKeyAscii)))
                End If
            Else    'long date
                intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
                If intKeyAscii = vbKeyBack Then
                    If intTemp = 3 Or intTemp = 4 Then
                        Mask = Left(strCurrentText, 6)
                    Else    'either an error occured or it's the month of august
                        Mask = ""
                    End If
                ElseIf (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else    'valid key was pressed
                    If intTemp = 3 Or intTemp = 4 Then  'march or april
                        Mask = FinishDay(intTemp, Val(Mid(strCurrentText, 7, 1)), Val(Chr(intKeyAscii)))
                    ElseIf intTemp = 8 Then 'august
                        If (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                            Mask = strCurrentText
                        Else
                            Mask = FinishDay(8, Val(Chr(intKeyAscii)))
                        End If
                    Else    'an error occured erase everything and start over
                        Mask = ""
                    End If
                End If
            End If
        Case 8  'eight characters
            If intDateTypeMask = 0 Then 'short date
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 7)
                ElseIf intKeyAscii < 48 Or intKeyAscii > 57 Then
                    Mask = strCurrentText
                Else
                    Mask = strCurrentText & Chr(intKeyAscii)
                End If
            ElseIf intDateTypeMask = 1 Then 'med date
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 4)
                ElseIf (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    Mask = strCurrentText & Chr(intKeyAscii)
                End If
            Else    'long date
                intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
                If intKeyAscii = vbKeyBack Then
                    If intTemp = 5 Then 'may
                        Mask = Left(strCurrentText, 5)
                    ElseIf intTemp = 8 Then 'august
                        Mask = "August "
                    Else    'an error occured or it's months Jan or Oct; erase everything
                        Mask = ""
                    End If
                ElseIf (intKeyAscii < 48) Or (intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    Select Case intTemp
                        Case 1, 5, 10  'jan, may or  oct
                            If intKeyAscii > 51 And intTemp <> 5 Then
                                Mask = Format(strCurrentText & "0" & Chr(intKeyAscii) & " 2000", "MMMM DD, ")
                            Else
                                Mask = strCurrentText & Chr(intKeyAscii)
                            End If
                        Case 8  'aug
                            If ((Val(Right(strCurrentText, 1)) = 3) And intKeyAscii > 49) Or ((Val(Right(strCurrentText, 1)) = 0) And intKeyAscii = 48) Then
                                Mask = strCurrentText
                            Else
                                Mask = strCurrentText & Chr(intKeyAscii) & ", "
                            End If
                        Case Else   'an error occured erase everything and start over
                            Mask = ""
                    End Select
                End If
            End If
        Case 9  'nine characters
            If intDateTypeMask = 0 Then 'short date
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 8)
                ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Or (Not (IsDate(Format(strCurrentText & Chr(intKeyAscii), "MM/DD/YYYY")))) Then
                    Mask = strCurrentText
                Else
                    Mask = strCurrentText & Chr(intKeyAscii)   'finished with short date
                End If
            ElseIf intDateTypeMask = 1 Then 'med date
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 8)
                ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    Mask = FinishYear(Month(strCurrentText & "000"), Day(strCurrentText & "000"), Val(Mid(strCurrentText, 9, 1)), Val(Chr(intKeyAscii)))
                End If
            Else    'long date
                intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
                If intKeyAscii = vbKeyBack Then
                    If intTemp = 1 Or intTemp = 10 Or intTemp = 5 Then 'jan or oct or may
                        Mask = Left(strCurrentText, Len(strCurrentText) - 1)
                    ElseIf intTemp = 6 Or intTemp = 7 Then  'june or july
                        Mask = Left(strCurrentText, 5)
                    Else    'an error has occured or months nov or dec
                        Mask = ""
                    End If
                ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    Select Case intTemp
                        Case 1, 10 'jan or oct
                            If (Val(Right(strCurrentText, 1)) = 0 And intKeyAscii = 48) Or (Val(Right(strCurrentText, 1)) = 3 And intKeyAscii > 49) Then
                                Mask = strCurrentText
                            Else
                                Mask = strCurrentText & Chr(intKeyAscii) & ", "
                            End If
                        Case 2  'feb
                            If intKeyAscii > 50 Then
                                Mask = strCurrentText & "0" & Chr(intKeyAscii) & ", "
                            Else
                                Mask = strCurrentText & Chr(intKeyAscii)
                            End If
                        Case 5, 6, 7 'may, june or july
                            If intTemp = 5 Then
                                Mask = FinishYear(intTemp, Val(Mid(strCurrentText, 5, 2)), Val(Right(strCurrentText, 1)), Val(Chr(intKeyAscii)))
                            Else
                                Mask = strCurrentText & Chr(intKeyAscii)
                            End If
                        Case 11, 12 'nov or dec
                            Mask = FinishDay(intTemp, Val(Chr(intKeyAscii)))
                        Case Else   'an error has occured so erase everything
                            Mask = "'"
                    End Select
                End If
            End If
        Case 10 'ten characters
            If intDateTypeMask = 0 Then 'short date...only BS allowed
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 9)
                Else
                    Mask = strCurrentText
                End If
            ElseIf intDateTypeMask = 1 Then 'med date
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 9)
                ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    Mask = strCurrentText & Chr(intKeyAscii)
                End If
            Else    'long date
                intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
                If intKeyAscii = vbKeyBack Then
                    Select Case intTemp
                        Case 2, 5, 6, 7, 11, 12 'feb, may, jun, jul, nov, or dec
                            Mask = Left(strCurrentText, Len(strCurrentText) - 1)
                        Case 3, 4   'mar or apr
                            Mask = Left(strCurrentText, 6)
                        Case Else   'error or sept
                            Mask = ""
                    End Select
                ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    Select Case intTemp
                        Case 2, 11, 12  'feb, nov, dec
                            Mask = FinishDay(intTemp, Val(Right(strCurrentText, 1)), Val(Chr(intKeyAscii)))
                        Case 6, 7   'jun, jul
                            Mask = FinishYear(intTemp, Day(Format(strCurrentText & "999", "MMMM DD, YYYY")), Val(Right(strCurrentText, 1)), Val(Chr(intKeyAscii)))
                        Case 3, 4, 5    'mar, apr, may
                            Mask = strCurrentText & Chr(intKeyAscii)
                        Case 9  'sep
                            Mask = FinishDay(9, Val(Chr(intKeyAscii)))
                        Case Else   'error
                            Mask = ""
                    End Select
                End If
            End If
        Case 11 'eleven characters
            If intDateTypeMask = 1 Then 'med date finished
                If intKeyAscii = vbKeyBack Then
                    Mask = Left(strCurrentText, 10)
                ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                    Mask = strCurrentText
                ElseIf IsDate(strCurrentText & Chr(intKeyAscii)) Then
                    Mask = strCurrentText & Chr(intKeyAscii)
                Else
                    Mask = strCurrentText
                End If
            Else    'long date
                intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
                If intKeyAscii = vbKeyBack Then
                    Select Case intTemp
                        Case 3, 4, 5, 6, 7, 9
                            Mask = Left(strCurrentText, Len(strCurrentText) - 1)
                        Case 8
                            Mask = Left(strCurrentText, 8)
                        Case Else
                            Mask = ""
                    End Select
                ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    Select Case intTemp
                        Case 3, 4
                            Mask = FinishYear(intTemp, Day(strCurrentText & "0"), Val(Right(strCurrentText, 1)), Val(Chr(intKeyAscii)))
                        Case 5
                            If IsDate(strCurrentText & Chr(intKeyAscii)) Then
                                Mask = strCurrentText & Chr(intKeyAscii)
                            Else
                                Mask = strCurrentText
                            End If
                        Case 6, 7, 8
                            Mask = strCurrentText & Chr(intKeyAscii)
                        Case 9
                            Mask = FinishDay(intTemp, Val(Right(strCurrentText, 1)), Val(Chr(intKeyAscii)))
                        Case Else
                            Mask = "'"
                    End Select
                End If
            End If
        Case 12 'twelve characters
            If intDateTypeMask = 1 Then 'med date back spaces only
                If intKeyAscii <> vbKeyBack Then
                    Mask = strCurrentText
                Else
                    Mask = Left(strCurrentText, Len(strCurrentText) - 1)
                End If
            Else    'long date
                intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
                If intKeyAscii = vbKeyBack Then
                    Select Case intTemp
                        Case 1
                            Mask = Left(strCurrentText, 9)
                        Case 3, 4, 5, 6, 7, 8
                            Mask = Left(strCurrentText, Len(strCurrentText) - 1)
                        Case 10
                            Mask = Left(strCurrentText, 9)
                        Case Else
                            Mask = ""
                    End Select
                ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                    Mask = strCurrentText
                Else
                    Select Case intTemp
                        Case 1, 3, 4, 6, 7, 10
                            Mask = strCurrentText & Chr(intKeyAscii)
                        Case 8
                            Mask = FinishYear(8, Day(strCurrentText & "000"), Val(Right(strCurrentText, 1)), Val(Chr(intKeyAscii)))
                        Case Else
                            Mask = strCurrentText
                    End Select
                End If
            End If
        Case 13 '13 characters...only long dates should be here at this point
            intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
            If intKeyAscii = vbKeyBack Then
                Select Case intTemp
                    Case 1, 3, 4, 6, 7, 8, 10
                        Mask = Left(strCurrentText, Len(strCurrentText) - 1)
                    Case 2, 11, 12
                        Mask = Left(strCurrentText, Len(strCurrentText) - 4)
                    Case Else
                        Mask = ""
                End Select
            ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                Mask = strCurrentText
            Else
                Select Case intTemp
                    Case 1, 10
                        Mask = FinishYear(intTemp, Day(strCurrentText & "000"), Val(Right(strCurrentText, 1)), Val(Chr(intKeyAscii)))
                    Case 2, 3, 4, 8, 11, 12
                        Mask = strCurrentText & Chr(intKeyAscii)
                    Case Else
                        Mask = strCurrentText
                End Select
            End If
        Case 14 '14 characters
            intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
            If intKeyAscii = vbKeyBack Then
                Select Case intTemp
                    Case 1, 2, 3, 4, 5, 8, 10, 11, 12
                        Mask = Left(strCurrentText, Len(strCurrentText) - 1)
                    Case 9
                        Mask = Left(strCurrentText, Len(strCurrentText) - 4)
                    Case Else
                        Mask = ""
                End Select
            ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                Mask = strCurrentText
            Else
                Select Case intTemp
                    Case 1, 8, 9, 10
                        Mask = strCurrentText & Chr(intKeyAscii)
                    Case 2, 11, 12
                        Mask = FinishYear(intTemp, Day(strCurrentText & "000"), Val(Right(strCurrentText, 1)), Val(Chr(intKeyAscii)))
                    Case Else
                        Mask = strCurrentText
                End Select
            End If
        Case 15 '15 characters
            intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
            If intKeyAscii = vbKeyBack Then
                Select Case intTemp
                    Case 1, 2, 8, 9, 10, 11, 12
                        Mask = Left(strCurrentText, Len(strCurrentText) - 1)
                    Case Else
                        Mask = "'"
                End Select
            ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                Mask = strCurrentText
            Else
                Select Case intTemp
                    Case 1, 2, 10, 11, 12
                        Mask = strCurrentText & Chr(intKeyAscii)
                    Case 9
                        Mask = FinishYear(9, Day(strCurrentText & "000"), Val(Right(strCurrentText, 1)), Val(Chr(intKeyAscii)))
                    Case Else
                        Mask = strCurrentText
                End Select
            End If
        Case 16 '16 cahracters
            intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
            If intKeyAscii = vbKeyBack Then
                Select Case intTemp
                    Case 1, 2, 9, 10, 11, 12
                        Mask = Left(strCurrentText, Len(strCurrentText) - 1)
                    Case Else
                        Mask = ""
                End Select
            ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                Mask = strCurrentText
            Else
                If IsDate(strCurrentText & Chr(intKeyAscii)) Then
                    Mask = strCurrentText & Chr(intKeyAscii)
                Else
                    Mask = strCurrentText
                End If
            End If
        Case 17 '17 characters
            intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
            If intKeyAscii = vbKeyBack Then
                Select Case intTemp
                    Case 2, 9, 11, 12
                        Mask = Left(strCurrentText, Len(strCurrentText) - 1)
                    Case Else
                        Mask = ""
                End Select
            ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) Then
                Mask = strCurrentText
            Else
                If intTemp = 9 Then
                    Mask = strCurrentText & Chr(intKeyAscii)
                Else
                    Mask = strCurrentText
                End If
            End If
        Case 18 '18 characters
            intTemp = Month(Left(strCurrentText, 3) & " 01,2000")
            If intKeyAscii <> vbKeyBack Then
                Mask = strCurrentText
            ElseIf intTemp = 9 Then
                Mask = Left(strCurrentText, Len(strCurrentText) - 1)
            Else
                Mask = strCurrentText
            End If
    End Select
    intKeyAscii = 0
End Function

Private Function FinishMonth(strMonth As String) As String
    Dim intMonth As Integer
    intMonth = Val(strMonth)
    If (intMonth > 0) Or (Asc(strMonth) = 48) Then
        If (intMonth > 1) And (intMonth < 13) Then
            If IsDate(Trim(intMonth) & "/01/2000") Then FinishMonth = Format(Trim(intMonth) & "/01/2000", strMonthFormat(intDateTypeMask))
        ElseIf intMonth > 12 Then
            FinishMonth = Format("12/01/2000", strMonthFormat(intDateTypeMask))
        ElseIf Len(strMonth) = 2 Then
            If intMonth = 0 Then intMonth = 1
            FinishMonth = Format(Trim(intMonth) & "/01/2000", strMonthFormat(intDateTypeMask))
        Else
            FinishMonth = Trim(intMonth)
        End If
    Else
        Select Case Len(strMonth)
            Case 1
                If UCase(Left(strMonth, 1)) = "F" Then
                    FinishMonth = Format("02/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf UCase(Left(strMonth, 1)) = "S" Then
                    FinishMonth = Format("09/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf UCase(Left(strMonth, 1)) = "O" Then
                    FinishMonth = Format("10/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf UCase(Left(strMonth, 1)) = "N" Then
                    FinishMonth = Format("11/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf UCase(Left(strMonth, 1)) = "D" Then
                    FinishMonth = Format("12/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf UCase(Left(strMonth, 1)) = "M" Then
                    FinishMonth = "Ma"
                ElseIf (UCase(Left(strMonth, 1)) = "J") Or (UCase(Left(strMonth, 1)) = "A") Then
                    FinishMonth = UCase(strMonth)
                Else
                    FinishMonth = ""
                End If
            Case 2
                If UCase(Left(strMonth, 2)) = "JA" Then
                    FinishMonth = Format("01/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf UCase(Left(strMonth, 2)) = "AP" Then
                    FinishMonth = Format("04/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf UCase(Left(strMonth, 2)) = "AU" Then
                    FinishMonth = Format("08/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf (UCase(Left(strMonth, 2)) = "JU") Then
                    FinishMonth = "Ju"
                Else
                    FinishMonth = Left(strMonth, 1)
                End If
            Case 3
                If UCase(Left(strMonth, 3)) = "JUL" Then
                    FinishMonth = Format("07/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf UCase(Left(strMonth, 3)) = "JUN" Then
                    FinishMonth = Format("06/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf UCase(Left(strMonth, 3)) = "MAR" Then
                    FinishMonth = Format("03/01/2000", strMonthFormat(intDateTypeMask))
                ElseIf UCase(Left(strMonth, 3)) = "MAY" Then
                    FinishMonth = Format("05/01/2000", strMonthFormat(intDateTypeMask))
                Else
                    FinishMonth = Left(strMonth, 2)
                End If
        End Select
    End If
End Function

Private Function FinishYear(intMonth As Integer, intDay As Integer, intFirstChar As Integer, intNextChar As Integer) As String
    Dim intLen As Integer
    Dim strMonth As String
    Dim strDay As String
    
    strMonth = Format(intMonth, "00")
    strDay = Format(intDay, "00")
    
    If intFirstChar = 0 Then
        If IsDate(Format(strMonth & "/" & strDay & "/200" & Trim(Str(intNextChar)), "MM/DD/YYYY")) Then
            FinishYear = Format(strMonth & "/" & strDay & "/200" & Trim(Str(intNextChar)), strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask) & "YYYY")
        Else
            intLen = Len(Format(strMonth & "/" & strDay & "/2000", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask) & "YYYY")) - 4
            FinishYear = Left(Format(strMonth & "/" & strDay & "/2000", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask) & "YYYY"), intLen)
        End If
    ElseIf intFirstChar > 2 Then
        If IsDate(Format(strMonth & "/" & strDay & "/" & "19" & Trim(Str(intFirstChar)) & Trim(Str(intNextChar)), "MM/DD/YYYY")) Then
            FinishYear = Format(strMonth & "/" & strDay & "/19" & Trim(Str(intFirstChar)) & Trim(Str(intNextChar)), strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask) & "YYYY")
        Else
            intLen = Len(Format(strMonth & "/" & strDay & "/" & Trim(Str(intFirstChar)) & Trim(Str(intNextChar)) & "00", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask) & "YYYY")) - 2
            FinishYear = Left(Format(strMonth & "/" & strDay & "/" & Trim(Str(intFirstChar)) & Trim(Str(intNextChar)) & "00", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask) & "YYYY"), intLen)
        End If
    Else
        intLen = Len(Format(strMonth & "/" & strDay & "/" & Trim(Str(intFirstChar)) & Trim(Str(intNextChar)) & "00", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask) & "YYYY")) - 2
        FinishYear = Left(Format(strMonth & "/" & strDay & "/" & Trim(Str(intFirstChar)) & Trim(Str(intNextChar)) & "00", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask) & "YYYY"), intLen)
    End If
End Function

Private Function FinishDay(intMonth As Integer, intFirstChar As Integer, Optional intNextChar As Integer = -1) As String
    Dim strMonth As String
    Dim strDay As String
    
    strMonth = Format(intMonth, "00")
    
    If intNextChar > -1 Then strDay = Format(intFirstChar & intNextChar, "00")
    
    Select Case intMonth
        Case 4, 6, 9, 11    '30 days
            If intFirstChar = 3 Then    'must be the 30th
                FinishDay = Format(strMonth & "/30/2000", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask))
            ElseIf intFirstChar > 3 Then    'must be the 0Nth
                FinishDay = Format(strMonth & "/" & Format(intFirstChar, "00") & "/2000", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask))
            ElseIf intNextChar = -1 Then    'must be 2 ,1 or 0
                FinishDay = Format(intMonth & "/01/2000", strMonthFormat(intDateTypeMask)) & Trim(intFirstChar)
            ElseIf IsDate(Format(strMonth & "/" & strDay & "/2000")) Then   'must be 2Nth, 1Nth, or 0Nth and valid
                FinishDay = Format(strMonth & "/" & strDay & "/2000", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask))
            Else    '00 was entered
                FinishDay = Format(intMonth & "/01/2000", strMonthFormat(intDateTypeMask)) & Trim(intFirstChar)
            End If
        Case 2
            If intFirstChar > 2 Then    'must be the 0Nth
                FinishDay = Format(strMonth & "/" & Format(intFirstChar, "00") & "/2000", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask))
            ElseIf intNextChar = -1 Then    'must be 2 ,1 or 0
                FinishDay = Format(intMonth & "/01/2000", strMonthFormat(intDateTypeMask)) & Trim(intFirstChar)
            ElseIf IsDate(Format(strMonth & "/" & strDay & "/2000")) Then   'must be 2Nth, 1Nth, or 0Nth and valid
                FinishDay = Format(strMonth & "/" & strDay & "/2000", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask))
            Else    '00 was entered
                FinishDay = Format(intMonth & "/01/2000", strMonthFormat(intDateTypeMask)) & Trim(intFirstChar)
            End If
        Case Else   '31 days
            If intFirstChar > 3 Then    'must be the 0Nth
                FinishDay = Format(strMonth & "/" & Format(intFirstChar, "00") & "/2000", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask))
            ElseIf intNextChar = -1 Then    'must be 3, 2 ,1 or 0
                FinishDay = Format(intMonth & "/01/2000", strMonthFormat(intDateTypeMask)) & Trim(intFirstChar)
            ElseIf IsDate(Format(strMonth & "/" & strDay & "/2000")) Then   'must be 3Nth, 2Nth, 1Nth, or 0Nth and valid
                FinishDay = Format(strMonth & "/" & strDay & "/2000", strMonthFormat(intDateTypeMask) & strDayFormat(intDateTypeMask))
            Else    '00 was entered
                FinishDay = Format(intMonth & "/01/2000", strMonthFormat(intDateTypeMask)) & Trim(intFirstChar)
            End If
    End Select
End Function
