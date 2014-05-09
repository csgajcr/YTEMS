Attribute VB_Name = "modCommon"
Option Explicit
'YTEMS全局变量
Public AppPath As String
Public ConfigPath As String
Public YTEMSServerIP As String
Public YTEMSServerPort As Long
Public YTEMSConnnection As Boolean
Public Function LoadServerIP(sConfigPath As String) As String
    LoadServerIP = ReadFromINI("YTEMS Common Config", "ServerIP", sConfigPath)
End Function
Public Function LoadServerPort(sConfigPath As String) As String
    LoadServerPort = ReadFromINI("YTEMS Common Config", "ServerPort", sConfigPath)
End Function
Public Function SaveServerIP(sConfigPath As String, IP As String)
    WriteToINI "YTEMS Common Config", "ServerIP", IP, sConfigPath
End Function
Public Function SaveServerPort(sConfigPath As String, PortID As Long)
    WriteToINI "YTEMS Common Config", "ServerPort", PortID, sConfigPath
End Function
Public Function IsNumber(str As String) As Boolean
    Dim i As Long
    If str = "" Then
        IsNumber = False
        Exit Function
    End If
    For i = 0 To Len(str) - 1
        If Asc(Mid(str, i + 1, 1)) > 57 Or Asc(Mid(str, i + 1, 1)) < 49 Then
            IsNumber = False
            Exit Function
        End If
    Next
    IsNumber = True
End Function
