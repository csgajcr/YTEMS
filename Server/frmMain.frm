VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "Codejock.SkinFramework.v12.0.1.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ͨ����ϵͳ �����"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   8085
   StartUpPosition =   3  '����ȱʡ
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
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   7560
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����ӵĿͻ��ˣ�"
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
Dim mysql_conn As New ADODB.Connection
Dim mysql_rs As New ADODB.Recordset
Dim mysql_filed As ADODB.Field
Sub InitiazationConfig()
    '��ʼ������ȫ�ֱ���
    AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    ConfigPath = AppPath & "Config.ini"
    '����Ƥ��
    SkinFramework.LoadSkin App.Path & "\Styles\iTunes", "Normalitunes.ini"
    SkinFramework.ApplyWindow Me.hWnd
    sckListen.Close
    sckListen.Bind LoadServerPort(ConfigPath)
    sckListen.Listen
    '�������ݿ�
    mysql_conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
    & "SERVER=127.0.0.1;" _
    & " DATABASE=test;" _
    & "UID=root;PWD=670510; OPTION=3"
    mysql_conn.Open
    mysql_rs.CursorLocation = adUseClient
    mysql_rs.Open "SELECT * FROM mytesttable ", mysql_conn
    mysql_rs.MoveFirst
    Do While Not mysql_rs.EOF
        MsgBox mysql_rs("id")
        mysql_rs.MoveNext
    Loop
    
End Sub

Private Sub Form_Load()
    InitiazationConfig
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    sckListen.Close
    Dim i As Long
    For i = 0 To sckServer.UBound
        sckServer(i).Close
    Next
End Sub



Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
    Dim i As Long
    'MsgBox "Accept Success!" & sckServer(0).RemoteHostIP
    For i = 0 To sckServer.UBound
        If sckServer(i).State = sckClosed Then
            sckServer(i).Accept requestID
            '------
            lstUser.AddItem i & "-" & sckServer(i).RemoteHostIP
            
            '------
            Exit Sub
        End If
    Next
    Load sckServer(i)
    sckServer(i).Accept requestID
    '-----
    lstUser.AddItem i & "-" & sckServer(i).RemoteHostIP
    
    '-----
End Sub

Private Sub sckServer_Close(Index As Integer)
    sckServer(Index).Close
    Dim i As Long
    For i = 0 To lstUser.ListCount - 1
        If CLng(Left(lstUser.List(i), 1)) = Index Then
            lstUser.RemoveItem i
        End If
    Next
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim RetData() As Byte, sData As String, sTmp() As String
    sckServer(Index).GetData sData, vbString
    If Left(sData, 25) = "YTEMSClientCommand-Login:" Then
        sTmp = Split(Mid(sData, 26, Len(sData) - 25), "|")
        If sTmp(0) = "jcr" And sTmp(1) = MD5("123456") Then
            sckServer(Index).SendData "YTEMSCommand:Login Success!"
        Else
            sckServer(Index).SendData "YTEMSCommand:Login Failed!Error:Username Or Password Wrong!"
        End If
    End If
End Sub
