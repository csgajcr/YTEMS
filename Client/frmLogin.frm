VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "Codejock.SkinFramework.Unicode.v12.0.1.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "˫�忼��ϵͳ ��½"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4605
   StartUpPosition =   2  '��Ļ����
   Begin VB.OptionButton optTeacher 
      Caption         =   "������ʦ"
      Height          =   300
      Left            =   3000
      TabIndex        =   12
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton optStudent 
      Caption         =   "����ѧ��"
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Value           =   -1  'True
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   5160
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "����(&S)"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&X)"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "��½(&L)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      X1              =   75
      X2              =   4555
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ�汾��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   3000
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
      Caption         =   "��   �룺"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�û�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Top             =   3000
      Width           =   60
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ״̬��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   3000
      Width           =   900
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'sckClosed 0 �ر�״̬
'sckOpen 1 ��״̬
'sckListening 2 ����״̬
'sckConnectionPending 3 ���ӹ���
'sckResolvingHost 4 ��������
'sckHostResolved 5 ��ʶ������
'sckConnecting 6 ��������
'sckConnected 7 ������
'sckClosing 8 ͬ����Ա���ڹر�����
'sckError 9 ����


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
        MsgBox "�������û���������", vbCritical
        Exit Sub
    End If
    
    sckClient.Close
    sckClient.Connect YTEMSServerIP, YTEMSServerPort
    lblStatus.Caption = "���ڳ������ӷ�����......"
    cmdLogin.Enabled = False
    cmdConfig.Enabled = False
End Sub

Private Sub Form_Load()
    InitializationConfig
End Sub
Sub InitializationConfig()
    '��ʼ������ȫ�ֱ���
    AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    ConfigPath = AppPath & "Config.ini"
    IsBinaryTransfer = False
    IsHeadPicture = False
    IsExamFile = False
    '����Ƥ��
    SkinFramework.LoadSkin App.Path & "\Styles\iTunes", "Normalitunes.ini"
    SkinFramework.ApplyWindow Me.hwnd
    '��ʼ��IP�Ͷ˿�
    YTEMSServerIP = LoadServerIP(ConfigPath)
    YTEMSServerPort = LoadServerPort(ConfigPath)
    '����LOGO
    lblStatus.Caption = "�����ʼ�����"
    lblVersion.Caption = App.Major & "." & App.Minor
End Sub
Sub UnloadConfig()
    If sckClient.State <> sckClosed Then sckClient.Close
End Sub



Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub sckClient_Close()
    On Error Resume Next
    sckClient.Close
    If frmLoading.Visible = True Then Unload frmLoading
    MsgBox "����������ӶϿ�", vbInformation
    cmdLogin.Enabled = True
    cmdConfig.Enabled = True
    lblStatus.Caption = "����������ӶϿ�"
    
    frmLogin.Show
    Unload frmMain
    Unload frmAdmin
End Sub

Private Sub sckClient_Connect()
    lblStatus.Caption = "���ӳɹ�....���ڵȴ���֤..."
    If optStudent.Value Then
        'sckClient.SendData "YTEMSClientCommand-Login:" & txtUserName.Text & "|" & MD5(txtPassword.Text)
        sckClient.SendData CS_MSG_STU_REQUEST_LOGIN
        sckClient.SendData txtUserName.Text & "|" & MD5(txtPassword.Text)
    Else
        'sckClient.SendData "YTEMSClientCommand-TeacherLogin:" & txtUserName.Text & "|" & MD5(txtPassword.Text)
        sckClient.SendData CS_MSG_TEACHER_REQUEST_LOGIN
        sckClient.SendData txtUserName.Text & "|" & MD5(txtPassword.Text)
    End If
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim Cmd As Byte, sTmp() As String
    Dim sData As String
    Dim lTmp As Long
    'Dim StuInfo As StudentInformation
    Dim i As Integer
    Dim ExamInfo() As ExamInformation
    Dim Filenum As Integer
    Dim byt() As Byte
Start:
    Cmd = 0
    sckClient.GetData Cmd, , 1
    '1�ֽ�ָ��Cmd
    Select Case Cmd
    Case SC_MSG_LOGIN_SUCCESS
        frmLoading.Show
        '---------------------------��½�ɹ�����ѧ��������Ϣ------------------------------
        sckClient.GetData sData, , bytesTotal - 1
        
        sTmp() = Split(sData, "|")
        cmdLogin.Enabled = True
        cmdConfig.Enabled = True
        StuInfo.ClassNo = sTmp(0)
        StuInfo.DeptNo = sTmp(1)
        StuInfo.S_JoinYear = sTmp(2)
        StuInfo.StuName = sTmp(3)
        StuInfo.StuPw = sTmp(4)
        StuInfo.StuSex = sTmp(5)
        StuInfo.UID = sTmp(6)
        Me.Hide
        frmMain.Show
        frmMain.txtUserName = RemoveMask(StuInfo.StuName)
        frmMain.txtClassNo = RemoveMask(StuInfo.ClassNo)
        frmMain.txtJoinYear = RemoveMask(StuInfo.S_JoinYear)
        frmMain.txtSex = RemoveMask(StuInfo.StuSex)
        frmMain.txtUID = RemoveMask(StuInfo.UID)
        '---------------------------------����ѧ����ؿ�����Ϣ-----------------------------
        'sTmp(7)Ϊ������Ϣ����
        ReDim Preserve ExamInfo(CInt(sTmp(7)) - 1)
        For i = 0 To CInt(sTmp(7)) - 1
            ExamInfo(i).ExamDataTime = sTmp(7 + i * 4 + 1)
            ExamInfo(i).ExamID = sTmp(7 + i * 4 + 2)
            ExamInfo(i).ExamName = sTmp(7 + i * 4 + 3)
            ExamInfo(i).ExamTime = sTmp(7 + i * 4 + 4)
            frmMain.lstExamInformation.ListItems.Add , , RemoveMask(ExamInfo(i).ExamName)
            frmMain.lstExamInformation.ListItems(frmMain.lstExamInformation.ListItems.Count).SubItems(1) = RemoveMask(ExamInfo(i).ExamID)
            frmMain.lstExamInformation.ListItems(frmMain.lstExamInformation.ListItems.Count).SubItems(2) = ExamInfo(i).ExamDataTime
            frmMain.lstExamInformation.ListItems(frmMain.lstExamInformation.ListItems.Count).SubItems(3) = RemoveMask(ExamInfo(i).ExamTime)
        Next
        
        
        Unload frmLoading
    Case SC_MSG_LOGIN_FAILED
        '��¼ʧ��
        MsgBox "�û������������", vbCritical
        lblStatus.Caption = "�û������������"
        cmdLogin.Enabled = True
        cmdConfig.Enabled = True
        Exit Sub
    Case SC_MSG_STUDENT_MORE_INFORMATION
        'ѧ��������Ϣ
        sckClient.GetData sData, , bytesTotal - 1
        sTmp() = Split(sData, "|")
        Dim StuMoreInfo As StudentMoreInfo
        StuMoreInfo.ClassDtor = sTmp(0)
        StuMoreInfo.ClassName = sTmp(1)
        StuMoreInfo.Dept = sTmp(2)
        StuMoreInfo.DeptDtor = sTmp(3)
        frmMoreInfo.txtClassName = RemoveMask(StuMoreInfo.ClassName)
        frmMoreInfo.txtClassDtor = RemoveMask(StuMoreInfo.ClassDtor)
        frmMoreInfo.txtDept = RemoveMask(StuMoreInfo.Dept)
        frmMoreInfo.txtDeptDtor = RemoveMask(StuMoreInfo.DeptDtor)
        Unload frmLoading
    Case SC_MSG_FILE_TRANSFER
        '��Ҫ���ļ�
        sckClient.GetData lTmp, , 4
        BinaryFileLength = lTmp
        sckClient.GetData lTmp, 4
        sckClient.GetData sData, , lTmp
        BinaryTransferFileName = sData
        If Dir(AppPath & "temp\" & BinaryTransferFileName) <> "" Then
            Kill AppPath & "temp\" & BinaryTransferFileName
        End If
        Select Case sData
            
        End Select
        CurrentLength = 0
        sData = ""
        lTmp = 0
        GoTo Start
    Case SC_MSG_FILE_DATA
        '�ļ�����
        Filenum = FreeFile
        Open AppPath & "temp\" & BinaryTransferFileName For Binary As #Filenum
        sckClient.GetData lTmp, , 4
        CurrentLength = CurrentLength + lTmp
        ReDim Preserve byt(lTmp - 1)
        sckClient.GetData byt, , lTmp
        Put #Filenum, LOF(Filenum) + 1, byt
        Close #Filenum
        If CurrentLength < BinaryFileLength Then                                ' �ж��ļ��Ƿ������
            GoTo Start
        End If
    Case SC_MSG_SET_PASSWORD_SUCCESS
        StuInfo.StuPw = NewPassword
        MsgBox "�����޸ĳɹ�", vbInformation
    Case SC_MSG_SET_PASSWORD_FAILED
        MsgBox "�����޸�ʧ��", vbCritical
    Case SC_MSG_ALLOW_ENTER_EXAM                                                '���뿼��
        MsgBox "������뿼��", vbInformation
        
        '--------------------------
    Case SC_MSG_NOT_ALLOW_ENTER_EXAM
        MsgBox "Ŀǰ��������뿼��", vbCritical
    Case SC_MSG_TEACHER_LOGIN_SUCCESS
        cmdLogin.Enabled = True
        cmdConfig.Enabled = True
        '--------------------------------------------���ܽ�ʦ�����Ϣ-------------------
        sckClient.GetData sData, , bytesTotal - 1
        sTmp() = Split(sData, "|")
        TcInfo.DeptNo = sTmp(0)
        TcInfo.JoinYear = sTmp(1)
        TcInfo.Password = sTmp(2)
        TcInfo.TeacherName = sTmp(3)
        TcInfo.TeacherSex = sTmp(4)
        TcInfo.UID = sTmp(5)
        frmAdmin.Show
        Me.Hide
        frmAdmin.txtJoinYear.Text = TcInfo.JoinYear
        frmAdmin.txtUserName.Text = TcInfo.TeacherName
        frmAdmin.txtUID.Text = TcInfo.UID
    Case SC_MSG_TEACHER_LOGIN_FAILED
        '��¼ʧ��
        MsgBox "�û������������", vbCritical
        lblStatus.Caption = "�û������������"
        cmdLogin.Enabled = True
        cmdConfig.Enabled = True
        Exit Sub
    Case SC_MSG_TEACHER_SET_PASSWORD_SUCCESS
        TcInfo.Password = NewPassword
        MsgBox "�����޸ĳɹ�", vbInformation
    Case SC_MSG_TEACHER_SET_PASSWORD_FAILED
        MsgBox "�����޸�ʧ��", vbCritical
    End Select
    
    
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'MsgBox Number & Description & Scode & Source
    If Number = 10060 Then
        MsgBox "�����δ����", vbInformation
        cmdLogin.Enabled = True
        cmdConfig.Enabled = True
        lblStatus.Caption = "����ʧ��"
        Exit Sub
    Else
        MsgBox "������������", vbInformation
        lblStatus.Caption = "����ʧ��"
        cmdLogin.Enabled = True
        cmdConfig.Enabled = True
        Exit Sub
    End If
End Sub


