VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "Codejock.SkinFramework.v12.0.1.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "移通考试系统 登陆"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4605
   StartUpPosition =   2  '屏幕中心
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   5160
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "设置(&S)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1720
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&X)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "登陆(&L)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   0
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Line Line1 
      X1              =   80
      X2              =   4560
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   2520
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "当前版本："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   2520
      Width           =   900
   End
   Begin VB.Image imgLogo 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "密   码："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   1410
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用户名："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   930
      Width           =   960
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   60
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "当前状态："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   900
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'sckClosed 0 关闭状态
'sckOpen 1 打开状态
'sckListening 2 侦听状态
'sckConnectionPending 3 连接挂起
'sckResolvingHost 4 解析域名
'sckHostResolved 5 已识别主机
'sckConnecting 6 正在连接
'sckConnected 7 已连接
'sckClosing 8 同级人员正在关闭连接
'sckError 9 错误
Dim IsStuInfo As Boolean

Private Sub cmdConfig_Click()
    frmConfig.Show 1
End Sub

Private Sub cmdExit_Click()
    UnloadConfig
    End
End Sub

Private Sub cmdLogin_Click()
    'frmMain.Show
    'Me.Hide
    If txtUserName.Text = "" Or txtPassword.Text = "" Then
        MsgBox "请输入用户名或密码", vbCritical
        Exit Sub
    End If
    
    sckClient.Close
    sckClient.Connect YTEMSServerIP, YTEMSServerPort
    lblStatus.Caption = "正在尝试连接服务器......"
    cmdLogin.Enabled = False
    cmdConfig.Enabled = False
End Sub

Private Sub Form_Load()
    InitializationConfig
End Sub
Sub InitializationConfig()
    '初始化部分全局变量
    AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    ConfigPath = AppPath & "Config.ini"
    IsStuInfo = False
    '加载皮肤
    SkinFramework.LoadSkin App.Path & "\Styles\iTunes", "Normalitunes.ini"
    SkinFramework.ApplyWindow Me.hwnd
    '初始化IP和端口
    YTEMSServerIP = LoadServerIP(ConfigPath)
    YTEMSServerPort = LoadServerPort(ConfigPath)
    '加载LOGO
    lblStatus.Caption = "程序初始化完毕"
    lblVersion.Caption = App.Major & "." & App.Minor
End Sub
Sub UnloadConfig()
    If sckClient.State <> sckClosed Then sckClient.Close
End Sub



Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub sckClient_Close()
    sckClient.Close
    If frmLoading.Visible = True Then Unload frmLoading
    MsgBox "与服务器连接断开", vbInformation
    cmdLogin.Enabled = True
    cmdConfig.Enabled = True
    lblStatus.Caption = "与服务器连接断开"
    
    frmLogin.Show
    Unload frmMain
End Sub

Private Sub sckClient_Connect()
    lblStatus.Caption = "连接成功....正在等待验证..."
    sckClient.SendData "YTEMSClientCommand-Login:" & txtUserName.Text & "|" & MD5(txtPassword.Text)
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String * 100
    sckClient.GetData sData, vbString, 100
    
    'MsgBox Hex(StrPtr(sData))
    If Left(sData, 27) = "YTEMSCommand:Login Success!" Then
        frmLoading.Show
        '----------------------------接受学生信息
        Dim StuName() As Byte
        sckClient.GetData StuInfo.ClassNo, vbString, 10
        sckClient.GetData StuInfo.DeptNo, vbString, 10
        sckClient.GetData StuInfo.S_JoinYear, vbString, 4
        sckClient.GetData StuInfo.StuName, vbString, 10
        'sckClient.GetData StuInfo.StuPw, vbString, 32
        sckClient.GetData StuInfo.StuSex, vbString, 10
        sckClient.GetData StuInfo.UID, vbString, 10
        cmdLogin.Enabled = True
        cmdConfig.Enabled = True
        Me.Hide
        frmMain.Show
        frmMain.txtUserName = StuInfo.StuName
        frmMain.txtClassNo = StuInfo.ClassNo
        frmMain.txtJoinYear = StuInfo.S_JoinYear
        frmMain.txtSex = StuInfo.StuSex
        frmMain.txtUID = StuInfo.UID
        '----------接受图片
        SocketReceiveHeadPic frmMain.imgHead, sckClient
        '------------------------
        Unload frmLoading
    ElseIf Left(sData, 59) = "YTEMSCommand:Login Failed!Error:Username Or Password Wrong!" Then
        MsgBox "用户名或密码错误", vbCritical
        lblStatus.Caption = "用户名或密码错误"
        cmdLogin.Enabled = True
        cmdConfig.Enabled = True
        Exit Sub
    ElseIf Left(sData, 28) = "YTEMSCommand:StudentMoreInfo" Then
        Dim StuMoreInfo As StudentMoreInfo
        sckClient.GetData StuMoreInfo.ClassDtor, vbString, 10
        sckClient.GetData StuMoreInfo.ClassName, vbString, 10
        sckClient.GetData StuMoreInfo.Dept, vbString, 10
        sckClient.GetData StuMoreInfo.DeptDtor, vbString, 10
        frmMoreInfo.txtClassName = StuMoreInfo.ClassName
        frmMoreInfo.txtClassDtor = StuMoreInfo.ClassDtor
        frmMoreInfo.txtDept = StuMoreInfo.Dept
        frmMoreInfo.txtDeptDtor = StuMoreInfo.DeptDtor
        Unload frmLoading
    End If
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'MsgBox Number & Description & Scode & Source
    If Number = 10060 Then
        MsgBox "服务端未运行", vbInformation
        cmdLogin.Enabled = True
        cmdConfig.Enabled = True
        lblStatus.Caption = "连接失败"
        Exit Sub
    Else
        MsgBox "请检查网络连接", vbInformation
        lblStatus.Caption = "连接失败"
        cmdLogin.Enabled = True
        cmdConfig.Enabled = True
        Exit Sub
    End If
End Sub


