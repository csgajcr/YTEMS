Attribute VB_Name = "modCommon"
Option Explicit
'YTEMS全局变量
Public IsBinaryTransfer As Boolean                                              '是否为传输状态
Public BinaryTransferFileName As String                                         '文件路径
Public BinaryFileLength As Long                                                 '文件长度
Public IsHeadPicture As Boolean
Public IsExamFile As Boolean                                                    '考试文件
Public CurrentLength As Long
Public ChoiceCfg() As ChoiceConfig
'-------
Public AppPath As String
Public ConfigPath As String
Public YTEMSServerIP As String
Public YTEMSServerPort As Long
Public YTEMSConnnection As Boolean
Public StuInfo As StudentInformation
Public TcInfo As TeacherInformation
Public NewPassword As String
'--------
Public Type TeacherInformation
    UID As String
    TeacherName As String
    TeacherSex As String
    Password As String
    DeptNo As String
    JoinYear As String
End Type
Public Type ExamInformation
    ExamName As String
    ExamID As String
    ExamDataTime As String
    ExamTime As String
End Type
Public Type StudentInformation
    UID As String
    StuName As String
    StuSex As String
    StuPw As String
    DeptNo As String
    ClassNo As String
    S_JoinYear As String
End Type
Public Type StudentMoreInfo
    ClassName As String
    ClassDtor As String
    Dept As String
    DeptDtor As String
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
    Dim Filenum As Integer, FileLength As Long, i As Long
    Dim c As Long
    c = 0
    Filenum = FreeFile
    sck.GetData FileLength, vbLong, 4
    Open AppPath & "temp\Head.jpg" For Binary As #Filenum
    Do
        sck.GetData byt, , 1024
        If FileLength - c < 1024 Then
            For i = 0 To FileLength - c - 1
                Put #Filenum, , byt(i)
            Next
            Exit Do
        End If
        c = c + 1024
        Put #Filenum, , byt
    Loop
    Close Filenum
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
    Dim Filenum As Integer, FileLength As Long, i As Long
    Dim c As Long
    c = 0
    Filenum = FreeFile
    sck.GetData FileLength, vbLong, 4
    Open FilePath For Binary As #Filenum
    Do
        sck.GetData byt, , 1024
        If FileLength - c < 1024 Then
            For i = 0 To FileLength - c - 1
                Put #Filenum, , byt(i)
            Next
            Exit Do
        End If
        c = c + 1024
        Put #Filenum, , byt
    Loop
    Close Filenum
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
Public Function GetPathFromFileName(FileName As String) As String
    Dim sTmp As String
    Dim i As Integer
    For i = Len(FileName) To 1 Step -1
        If Mid(FileName, i, 1) = "\" Then
            sTmp = Left(FileName, i)
            GetPathFromFileName = sTmp
            Exit Function
        End If
    Next
End Function
