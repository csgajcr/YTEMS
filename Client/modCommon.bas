Attribute VB_Name = "modCommon"
Option Explicit
'YTEMS全局变量
Public IsBinaryTransfer As Boolean
Public BinaryTransferFileName As String
Public BinaryFileLength As Long
Public IsHeadPicture As Boolean
Public IsExamFile As Boolean
'-------
Public AppPath As String
Public ConfigPath As String
Public YTEMSServerIP As String
Public YTEMSServerPort As Long
Public YTEMSConnnection As Boolean
Public StuInfo As StudentInformation
Public NewPassword As String
Public Type ExamInformation
    ExamName As String * 20
    ExamID As String * 10
    ExamDataTime As String * 30
    ExamTime As String * 10
End Type
Public Type StudentInformation
    UID As String * 10
    StuName As String * 10
    StuSex As String * 10
    StuPw As String * 32
    DeptNo As String * 10
    ClassNo As String * 10
    S_JoinYear As String * 4
End Type
Public Type StudentMoreInfo
    ClassName As String * 10
    ClassDtor As String * 10
    Dept As String * 10
    DeptDtor As String * 10
End Type
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
Public Function SocketReceiveHeadPic(img As Image, sck As Winsock)
    'On Error GoTo myerr
    Dim byt() As Byte
    Dim FileNum As Integer, FileLength As Long, i As Long
    Dim c As Long
    c = 0
    FileNum = FreeFile
    sck.GetData FileLength, vbLong, 4
    Open AppPath & "temp\Head.jpg" For Binary As #FileNum
    Do
        sck.GetData byt, , 1024
        If FileLength - c < 1024 Then
            For i = 0 To FileLength - c - 1
                Put #FileNum, , byt(i)
            Next
            Exit Do
        End If
        c = c + 1024
        Put #FileNum, , byt
    Loop
    Close FileNum
    img.Stretch = True
    img.Picture = LoadPicture(AppPath & "temp\Head.jpg")
    Kill (AppPath & "temp\Head.jpg")
    Exit Function
myerr:
    MsgBox Err.Number & Err.Description
End Function
Public Function SocketReceiveBinaryFile(FilePath As String, sck As Winsock)
    'On Error GoTo myerr
    Dim byt() As Byte
    Dim FileNum As Integer, FileLength As Long, i As Long
    Dim c As Long
    c = 0
    FileNum = FreeFile
    sck.GetData FileLength, vbLong, 4
    Open FilePath For Binary As #FileNum
    Do
        sck.GetData byt, , 1024
        If FileLength - c < 1024 Then
            For i = 0 To FileLength - c - 1
                Put #FileNum, , byt(i)
            Next
            Exit Do
        End If
        c = c + 1024
        Put #FileNum, , byt
    Loop
    Close FileNum
    Exit Function
myerr:
    MsgBox Err.Number & Err.Description
End Function
Public Function RemoveMask(str As String) As String
    Dim i As Long, sTmp As String
    If InStr(1, str, Chr(32)) Then
        
        sTmp = Left(str, InStr(1, str, Chr(32)) - 1)
        RemoveMask = sTmp
    Else
        sTmp = str
        RemoveMask = sTmp
    End If
End Function
