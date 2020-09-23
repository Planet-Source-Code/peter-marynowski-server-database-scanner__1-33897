Attribute VB_Name = "Module1"
Option Explicit

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetUserLogin() As String

    Dim cBuffer As String
    Dim nSize As Long
    Dim nReturnVal As Long

    cBuffer = "                              "
    nSize = 30
    nReturnVal = GetUserName(cBuffer, nSize)
    If nReturnVal > 0 Then
        GetUserLogin = Trim(Left(cBuffer, nSize - 1))
    Else
        GetUserLogin = "Unknown User"
    End If

End Function
