VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "Codejock.SkinFramework.v12.0.1.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "移通考试系统 服务端"
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
    Dim RetData() As Byte, sData As String * 100, sTmp() As String
    Dim StuInfo As StudentInformation
    Dim TcInfo As TeacherInformation
    Dim sTmp2 As String * 100
    sckServer(Index).GetData sData, vbString
    '------------------用户请求登陆----------------------------
    If Left(sData, 25) = "YTEMSClientCommand-Login:" Then
        sTmp = Split(Mid(sData, 26, Len(sData) - 25), "|")
        If SQLQueryStudentInfo("tb_student", sTmp(0), StuInfo) Then
            If Left(sTmp(1), 24) = Left(StuInfo.StuPw, 24) Then
                
                sTmp2 = "YTEMSCommand:Login Success!"
                sckServer(Index).SendData sTmp2
                sckServer(Index).SendData StuInfo.ClassNo
                sckServer(Index).SendData StuInfo.DeptNo
                sckServer(Index).SendData StuInfo.S_JoinYear
                SocketSendWideChar StuInfo.StuName, 10, sckServer(Index)
                sckServer(Index).SendData StuInfo.StuPw
                SocketSendWideChar StuInfo.StuSex, 10, sckServer(Index)
                sckServer(Index).SendData StuInfo.UID
                '-----------发送考试信息
                Dim Examinfo() As ExamInformation
                If SQLQueryExamInformation("tb_exammanage", "tb_examminfo", StuInfo.ClassNo, Examinfo) Then
                    SocketSendExamInformation Examinfo, sckServer(Index)
                Else
                    Dim ExamInfoLength As Long
                    ExamInfoLength = 0
                    sckServer(Index).SendData ExamInfoLength
                End If
                '--------发送图片
                If Dir(AppPath & "UserPicture\" & RemoveMask(StuInfo.UID) & ".jpg") <> "" Then
                    SocketSendBinaryFile AppPath & "UserPicture\" & RemoveMask(StuInfo.UID) & ".jpg", sckServer(Index)
                Else
                    SocketSendBinaryFile AppPath & "UserPicture\" & "Default.jpg", sckServer(Index)
                End If
                
                '-------------------------
                lstUser.AddItem sckServer(Index).RemoteHostIP & " | " & StuInfo.UID
            Else
                sckServer(Index).SendData "YTEMSCommand:Login Failed!Error:Username Or Password Wrong!"
                
            End If
        Else
            sckServer(Index).SendData "YTEMSCommand:Login Failed!Error:Username Or Password Wrong!"
            
        End If
        '------------------教师请求登陆----------------------------
    ElseIf Left(sData, 32) = "YTEMSClientCommand-TeacherLogin:" Then
        sTmp = Split(Mid(sData, 33, Len(sData) - 32), "|")
        If SQLQueryTeacherInfo("tb_teacher", sTmp(0), TcInfo) Then
            If Left(sTmp(1), 24) = Left(TcInfo.Password, 24) Then
                sTmp2 = "YTEMSCommand:Login Success!"
                sckServer(Index).SendData sTmp2
                sckServer(Index).SendData TcInfo.DeptNo
                sckServer(Index).SendData TcInfo.JoinYear
                sckServer(Index).SendData TcInfo.Password
                SocketSendWideChar TcInfo.TeacherName, 10, sckServer(Index)
                SocketSendWideChar TcInfo.TeacherSex, 10, sckServer(Index)
                sckServer(Index).SendData TcInfo.UID
                '-----------发送考试信息
                
                '--------发送图片
                
            End If
            
        Else
            sckServer(Index).SendData "YTEMSCommand:Login Failed!Error:Username Or Password Wrong!"
        End If
        
        '------------------用户请求获取更多学生信息----------------------------
    ElseIf Left(sData, 38) = "YTEMSClientCommand:GetMoreInformation:" Then
        Dim ClassNo As String * 10, DeptNo As String * 10
        sTmp = Split(Mid(sData, 39, Len(sData) - 38), "|")
        Dim StuMoreInfo As StudentMoreInfo
        If SQLQueryStudentMoreInfo("tb_class", "tb_Dept", sTmp(0), sTmp(1), StuMoreInfo) Then
            sTmp2 = "YTEMSCommand:StudentMoreInfo"
            sckServer(Index).SendData sTmp2
            SocketSendWideChar StuMoreInfo.ClassDtor, 10, sckServer(Index)
            SocketSendWideChar StuMoreInfo.ClassName, 10, sckServer(Index)
            SocketSendWideChar StuMoreInfo.Dept, 10, sckServer(Index)
            SocketSendWideChar StuMoreInfo.DeptDtor, 10, sckServer(Index)
        End If
        '------------------用户请求获取修改密码----------------------------
    ElseIf Left(sData, 34) = "YTEMSClientCommand:ChangePassword:" Then
        sTmp = Split(Mid(sData, 35, Len(sData) - 34), "|")
        If SQLSetStudentPassword("tb_student", sTmp(0), sTmp(1)) Then
            sTmp2 = "YTEMSCommand:SetPasswordSuccess!"
            sckServer(Index).SendData sTmp2
        Else
            sTmp2 = "YTEMSCommand:SetPasswordFailed!"
            sckServer(Index).SendData sTmp2
        End If
        '------------------用户请求进入考试----------------------------
    ElseIf Left(sData, 29) = "YTEMSClientCommand:EnterExam:" Then
        sTmp = Split(Mid(sData, 30, Len(sData) - 29), "|")
        Dim ExamDate As Date, TimeLength As Long
        ExamDate = sTmp(1)
        TimeLength = CLng(sTmp(2))
        If Now >= ExamDate And Now <= DateAdd("n", TimeLength, ExamDate) Then
            sTmp2 = "YTEMSCommand:AllowToEnterExam"
            sckServer(Index).SendData sTmp2
            SocketSendBinaryFile AppPath & "Examination Paper\" & RemoveMask(sTmp(0)) & ".bin", sckServer(Index)
        Else
            sTmp2 = "YTEMSCommand:Can Not Enter Exam!Date & Time Invalid"
            sckServer(Index).SendData sTmp2
        End If
        
    End If
End Sub
