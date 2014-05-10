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
Public Function SocketSendBinary(byt() As Byte, sck As Winsock)
    
End Function
Public Function SocketSendHeadPic(ByVal PicPath As String, sck As Winsock)
    Dim FileNum As Integer
    Dim byt(1023) As Byte
    Dim FileLength As Long
    FileNum = FreeFile
    Open PicPath For Binary As #FileNum
    FileLength = LOF(FileNum)
    sck.SendData FileLength
    Do While Not EOF(FileNum)
        Get #FileNum, , byt
        sck.SendData byt
    Loop
    Close #FileNum
End Function
