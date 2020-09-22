Attribute VB_Name = "modProcedures"
Option Explicit

Declare Function GetPrivateProfileStringbyKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszreturnbuffer$, ByVal cchreturnbuffer&, ByVal lpszFile$)

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String)

Public Declare Function GetCurrentDirectoryA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public strCTRLName As String
Public strVersion As String

Public Function GetSetting(strPropertyPage As String, strProperty As String) As String
    Dim lCharacters As Long
    Dim strTemp As String
    
    strTemp = String$(128, 0)
    lCharacters = GetPrivateProfileStringbyKeyName(strPropertyPage, strProperty, "", strTemp, 127, strCTRLName)
        
    GetSetting = Left$(strTemp, lCharacters)
End Function

Public Sub SaveSetting(strPropertyPage As String, strProperty As String, ByVal strValue As String)
    WritePrivateProfileString strPropertyPage, strProperty, strValue, strCTRLName
End Sub

