VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "Codejock.SkinFramework.Unicode.v12.0.1.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "双体考试系统 服务端"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8085
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdConfig 
      Caption         =   "服务器设置"
      Height          =   615
      Left            =   6720
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstUser 
      Height          =   3120
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   7560
      Top             =   1800
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   7560
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   7560
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "已连接的客户端："
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub InitiazationConfig()
    '初始化部分全局变量
    AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    ConfigPath = AppPath & "Config.ini"
    '加载皮肤
    SkinFramework.LoadSkin App.Path & "\Styles\iTunes", "Normalitunes.ini"
    SkinFramework.ApplyWindow Me.hWnd
    sckListen.Close
    sckListen.Bind LoadServerPort(ConfigPath)
    sckListen.Listen
    '连接数据库
    ConnectMySQL
End Sub
Sub ConnectMySQL()
    On Error GoTo myerr
    Dim sServer As String, sDBName As String, sUID As String, sPassword As String
    Dim SQLInfo As SQLConnectionInfo
    SQLInfo = LoadSQLConnectionInfo(ConfigPath)
    mysql_conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
    & "SERVER=" & SQLInfo.IPAddress & ";" _
    & " DATABASE=" & SQLInfo.DBName & ";" _
    & "UID=" & SQLInfo.UID & ";PWD=" & SQLInfo.Password & "; OPTION=3"
    mysql_conn.Open
    'mysql_rs.CursorLocation = adUseClient
    'mysql_rs.Open "SELECT * FROM mytesttable ", mysql_conn
    'mysql_rs.MoveFirst
    'Do While Not mysql_rs.EOF
    'MsgBox mysql_rs("id")
    'mysql_rs.MoveNext
    'Loop
    Exit Sub
myerr:
    MsgBox "数据库连接出错，请检查数据库设置！", vbCritical
    End
End Sub


Private Sub cmdConfig_Click()
    frmConfig.Show 1
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub cmdTest_Click()
    
    SocketSendBinaryFile "C:\Users\Jcr\Desktop\bd_logo1.png", sckServer(0)
End Sub

Private Sub Form_Initialize()
    If App.PrevInstance = True Then
        MsgBox "服务端已运行，请勿重复运行", vbInformation
        End
    End If
    
End Sub

Private Sub Form_Load()
    
    InitiazationConfig
    'tmp.DBName = "YTEMS"
    'tmp.IPAddress = "127.0.0.1"
    'tmp.UID = "root"
    'tmp.Password = "670510"
    'SaveSQLConnectionInfo ConfigPath, tmp
    'Dim a() As ExamInformation
    'SQLQueryExamInformation "tb_exammanage", "tb_examminfo", "02111301", a
    'MsgBox a(2).ExamDataTime
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    sckListen.Close
    Dim i As Long
    For i = 0 To sckServer.UBound
        sckServer(i).Close
    Next
    End
End Sub





Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
    Dim i As Long
    'MsgBox "Accept Success!" & sckServer(0).RemoteHostIP
    For i = 0 To sckServer.UBound
        If sckServer(i).State = sckClosed Then
            sckServer(i).Accept requestID
            '------
            'lstUser.AddItem i & "-" & sckServer(i).RemoteHostIP
            
            '------
            Exit Sub
        End If
    Next
    Load sckServer(i)
    sckServer(i).Accept requestID
    '-----
    'lstUser.AddItem i & "-" & sckServer(i).RemoteHostIP
    
    '-----
End Sub

Private Sub sckServer_Close(Index As Integer)
    
    Dim i As Long
    For i = 0 To lstUser.ListCount - 1
        If Left(lstUser.List(i), Len(sckServer(Index).RemoteHostIP)) = sckServer(Index).RemoteHostIP Then
            lstUser.RemoveItem i
        End If
    Next
    sckServer(Index).Close
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Cmd As Byte
    Dim sTmp() As String
    Dim sData As String
    Dim StuInfo As StudentInformation
    Dim ExamInfo() As ExamInformation
    Dim ExamDate As Date
    Dim TimeLength As Long
    Dim TcInfo As TeacherInformation
    sData = ""
    Cmd = 0
    sckServer(Index).GetData Cmd, , 1
    Select Case Cmd
    Case CS_MSG_STU_REQUEST_LOGIN
        sckServer(Index).GetData sData, , bytesTotal - 1
        sTmp = Split(sData, "|")                                                'sTmp数组为用户名和密码
        If SQLQueryStudentInfo("tb_student", sTmp(0), StuInfo) Then
            If Left(sTmp(1), 24) = Left(StuInfo.StuPw, 24) Then
                '登陆成功,发送学生基本信息
                sckServer(Index).SendData SC_MSG_LOGIN_SUCCESS
                sckServer(Index).SendData StuInfo.ClassNo & "|"
                sckServer(Index).SendData StuInfo.DeptNo & "|"
                sckServer(Index).SendData StuInfo.S_JoinYear & "|"
                sckServer(Index).SendData StuInfo.StuName & "|"
                sckServer(Index).SendData StuInfo.StuPw & "|"
                sckServer(Index).SendData StuInfo.StuSex & "|"
                sckServer(Index).SendData StuInfo.UID & "|"
                '--------------发送学生考试信息
                If SQLQueryExamInformation("tb_exammanage", "tb_examminfo", StuInfo.ClassNo, ExamInfo) Then
                    SocketSendExamInformation ExamInfo, sckServer(Index)
                Else
                    Dim ExamInfoCount As Long
                    ExamInfoCount = 0
                    sckServer(Index).SendData ExamInfoCount
                End If
                '-----------------------------------------------------------
                
            Else
                '登录失败
                sckServer(Index).SendData SC_MSG_LOGIN_FAILED
            End If
        Else
            '登录失败
            sckServer(Index).SendData SC_MSG_LOGIN_FAILED
            
        End If
        
        
        
    Case CS_MSG_REQUEST_STUDENT_MORE_INFORMATION
        sckServer(Index).GetData sData, , bytesTotal - 1
        sTmp = Split(sData, "|")
        
        Dim StuMoreInfo As StudentMoreInfo
        If SQLQueryStudentMoreInfo("tb_class", "tb_Dept", sTmp(0), sTmp(1), StuMoreInfo) Then
            sckServer(Index).SendData SC_MSG_STUDENT_MORE_INFORMATION
            sckServer(Index).SendData StuMoreInfo.ClassDtor & "|"
            sckServer(Index).SendData StuMoreInfo.ClassName & "|"
            sckServer(Index).SendData StuMoreInfo.Dept & "|"
            sckServer(Index).SendData StuMoreInfo.DeptDtor
        End If
        
    Case CS_MSG_FILE_TRANSFER
        
        
        
    Case CS_MSG_SET_PASSWORD
        sckServer(Index).GetData sData, , bytesTotal - 1
        sTmp = Split(sData, "|")
        If SQLSetStudentPassword("tb_student", sTmp(0), sTmp(1)) Then
            
            sckServer(Index).SendData SC_MSG_SET_PASSWORD_SUCCESS               '密码修改成功
        Else
            
            sckServer(Index).SendData SC_MSG_SET_PASSWORD_FAILED                '密码修改失败
        End If
    Case CS_MSG_REQUEST_ENTER_EXAM
        '请求进入考试
        sckServer(Index).GetData sData, , bytesTotal - 1
        sTmp = Split(sData, "|")
        ExamDate = sTmp(1)
        TimeLength = CLng(sTmp(2))
        If Now >= ExamDate And Now <= DateAdd("n", TimeLength, ExamDate) Then
            
            sckServer(Index).SendData SC_MSG_ALLOW_ENTER_EXAM                   '允许进入考试
            'SocketSendBinaryFile AppPath & "Examination Paper\" & RemoveMask(sTmp(0)) & ".bin", sckServer(Index)
        Else
            
            sckServer(Index).SendData SC_MSG_NOT_ALLOW_ENTER_EXAM               '不允许进入考试
        End If
    Case CS_MSG_TEACHER_REQUEST_LOGIN                                           '教师请求登录
        sckServer(Index).GetData sData, , bytesTotal - 1
        sTmp = Split(sData, "|")                                                'sTmp数组为用户名和密码
        If SQLQueryTeacherInfo("tb_teacher", sTmp(0), TcInfo) Then
            If Left(sTmp(1), 24) = Left(TcInfo.Password, 24) Then
                sckServer(Index).SendData SC_MSG_TEACHER_LOGIN_SUCCESS
                sckServer(Index).SendData TcInfo.DeptNo & "|"
                sckServer(Index).SendData TcInfo.JoinYear & "|"
                sckServer(Index).SendData TcInfo.Password & "|"
                sckServer(Index).SendData TcInfo.TeacherName & "|"
                sckServer(Index).SendData TcInfo.TeacherSex & "|"
                sckServer(Index).SendData TcInfo.UID
                
            Else
                sckServer(Index).SendData SC_MSG_TEACHER_LOGIN_FAILED
            End If
        Else
            sckServer(Index).SendData SC_MSG_TEACHER_LOGIN_FAILED
        End If
    Case CS_MSG_TEACHER_SET_PASSWORD
        sckServer(Index).GetData sData, , bytesTotal - 1
        sTmp = Split(sData, "|")
        If SQLSetTeacherPassword("tb_teacher", sTmp(0), sTmp(1)) Then
            
            sckServer(Index).SendData SC_MSG_TEACHER_SET_PASSWORD_SUCCESS       '密码修改成功
        Else
            
            sckServer(Index).SendData SC_MSG_TEACHER_SET_PASSWORD_FAILED        '密码修改失败
            
        End If
    End Select
    
    
    
End Sub
