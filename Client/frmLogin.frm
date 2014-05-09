VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "Codejock.SkinFramework.v12.0.1.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ͨ����ϵͳ ��½"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5625
   StartUpPosition =   2  '��Ļ����
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   5160
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "����"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "��½"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1200
      TabIndex        =   4
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   0
      Top             =   0
      _Version        =   786432
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Image imgLogo 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��  �룺"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   1695
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�û�����"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   975
      Width           =   735
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   960
      TabIndex        =   1
      Top             =   2880
      Width           =   90
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ״̬��"
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   2880
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
    '����Ƥ��
    SkinFramework.LoadSkin App.Path & "\Styles\iTunes", "Normalitunes.ini"
    SkinFramework.ApplyWindow Me.hWnd
    '��ʼ��IP�Ͷ˿�
    YTEMSServerIP = LoadServerIP(ConfigPath)
    YTEMSServerPort = LoadServerPort(ConfigPath)
    '����LOGO
    lblStatus.Caption = "�����ʼ�����"
End Sub
Sub UnloadConfig()
    If sckClient.State <> sckClosed Then sckClient.Close
End Sub

Private Sub sckClient_Close()
    sckClient.Close
    MsgBox "����������ӶϿ�", vbInformation
    frmLogin.Show
    Unload frmMain
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    sckClient.GetData sData, vbString
    If sData = "YTEMSCommand:Login Success!" Then
        Me.Hide
        frmMain.Show
    End If
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


