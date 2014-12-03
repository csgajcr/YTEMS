Attribute VB_Name = "ModCommon"
Option Explicit
'YTEMS全局变量
Public AppPath As String
Public ConfigPath As String
Public YTEMSServerIP As String
Public YTEMSServerPort As Long
Public mysql_conn As New ADODB.Connection
Public mysql_rs As New ADODB.Recordset
Public Type SQLConnectionInfo
    IPAddress As String
    DBName As String
    UID As String
    Password As String
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
Public Type ExamInformation
    ExamName As String
    ExamID As String
    ExamDataTime As String
    ExamTime As String
End Type
Public Type TeacherInformation
    UID As String * 10
    TeacherName As String * 10
    TeacherSex As String * 10
    Password As String * 32
    DeptNo As String * 10
    JoinYear As String * 4
End Type
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
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
Public Function LoadSQLConnectionInfo(sConfigPath As String) As SQLConnectionInfo
    Dim tmp As SQLConnectionInfo
    tmp.DBName = ReadFromINI("YTEMS Common Config", "MySQLDBName", sConfigPath)
    tmp.IPAddress = ReadFromINI("YTEMS Common Config", "MySQLIPAddress", sConfigPath)
    tmp.UID = ReadFromINI("YTEMS Common Config", "MySQLUsername", sConfigPath)
    tmp.Password = AESDecodeStr(ReadFromINI("YTEMS Common Config", "MySQLPassword", sConfigPath), tmp.UID, 256, 256)
    LoadSQLConnectionInfo = tmp
End Function
Public Function SaveSQLConnectionInfo(sConfigPath As String, SQLInfo As SQLConnectionInfo)
    WriteToINI "YTEMS Common Config", "MySQLDBName", SQLInfo.DBName, sConfigPath
    WriteToINI "YTEMS Common Config", "MySQLIPAddress", SQLInfo.IPAddress, sConfigPath
    WriteToINI "YTEMS Common Config", "MySQLUsername", SQLInfo.UID, sConfigPath
    WriteToINI "YTEMS Common Config", "MySQLPassword", AESEncodeStr(SQLInfo.Password, SQLInfo.UID, 256, 256), sConfigPath
End Function
Public Function SocketSendWideChar(ByVal str As String, ByVal Length As Long, sck As Winsock)
    Dim i As Integer
    Dim byt() As Byte
    byt = StrConv(str, vbFromUnicode)
    For i = 0 To Length - 1
        sck.SendData byt(i)
    Next
End Function
Public Function SocketSendBinaryFile(FilePath As String, sck As Winsock)
    Dim Filenum As Integer
    Dim byt() As Byte
    Dim FileLength As Long
    Dim ExtraLength As Long
    Dim i As Long
    Dim SendTimes As Long                                                       '发送次数
    Dim FileName As String
    Const StandardLength As Long = 1000
    Filenum = FreeFile
    Open FilePath For Binary As #Filenum
    FileLength = LOF(Filenum)
    FileName = GetFileNameFromPath(FilePath)
    sck.SendData SC_MSG_FILE_TRANSFER                                           '发送指令
    sck.SendData FileLength                                                     '发送文件大小
    sck.SendData CLng(Len(FileName))                                            '发送文件名长度
    sck.SendData FileName                                                       '发送文件名
    '----------------
    If FileLength <= StandardLength Then
        ReDim byt(FileLength - 1)
        Get #Filenum, , byt
        sck.SendData byt
    Else
        SendTimes = FileLength / StandardLength
        ExtraLength = FileLength Mod StandardLength
        For i = 1 To SendTimes
            ReDim byt(StandardLength - 1)
            Get #Filenum, , byt
            sck.SendData SC_MSG_FILE_DATA
            sck.SendData StandardLength
            sck.SendData byt
        Next
        If ExtraLength > 0 Then
            ReDim byt(ExtraLength - 1)
            Get #Filenum, , byt
            sck.SendData SC_MSG_FILE_DATA
            sck.SendData ExtraLength
            sck.SendData byt
        End If
        
    End If
    '------------------
    
    
    
    Close #Filenum
End Function
'---------------
Public Function SocketSendHeadPic(ByVal PicPath As String, sck As Winsock)
    Dim Filenum As Integer
    Dim byt() As Byte
    Dim FileLength As Long
    Filenum = FreeFile
    Open PicPath For Binary As #Filenum
    FileLength = LOF(Filenum)
    sck.SendData FileLength
    If FileLength < 1000 Then
        ReDim byt(FileLength - 1)
        Get #Filenum, , byt
        sck.SendData byt
    End If
    Do While Not EOF(Filenum)
        Get #Filenum, , byt
        sck.SendData byt
    Loop
    Close #Filenum
End Function
Public Function SocketSendExamInformation(ExamInfo() As ExamInformation, sck As Winsock)
    Dim ExamInfoCount As Long, i As Integer
    ExamInfoCount = (UBound(ExamInfo) + 1)
    sck.SendData CStr(ExamInfoCount)
    For i = 0 To UBound(ExamInfo)
        sck.SendData "|" & ExamInfo(i).ExamDataTime & "|"
        sck.SendData ExamInfo(i).ExamID & "|"
        'SocketSendWideChar ExamInfo(i).ExamName, 20, sck
        sck.SendData ExamInfo(i).ExamName & "|"
        sck.SendData ExamInfo(i).ExamTime
    Next
End Function
Public Function WaitForMysqlConnection()
    Do While (mysql_rs.State = 1)
        DoEvents
        mysql_rs.Close
        Sleep (10)
    Loop
End Function
Public Function AddQueto(ByRef str As String) As String
    
    Dim i As Long, sTmp As String
    If InStr(1, str, Chr(32)) Then
        
        sTmp = Chr(39) & Left(str, InStr(1, str, Chr(32)) - 1) & Chr(39)
        AddQueto = sTmp
    Else
        sTmp = Chr(39) & str & Chr(39)
        AddQueto = sTmp
    End If
    
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

Public Function GetFileNameFromPath(FilePath As String) As String
    Dim sTmp As String
    Dim sName As String
    Dim i  As Integer
    For i = Len(FilePath) To 1 Step -1
        If Mid(FilePath, i, 1) = "\" Then
            sName = Right(FilePath, Len(FilePath) - i)
            Exit For
        End If
    Next
    GetFileNameFromPath = sName
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
