VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#12.0#0"; "Codejock.SkinFramework.v12.0.1.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "移通考试系统 服务端"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   8085
   StartUpPosition =   3  '窗口缺省
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
            lstUser.AddItem sckServer(i).RemoteHostIP
            sckServer(i).SendData "YTEMSCommand:Login Success!"
            Exit Sub
        End If
    Next
    Load sckServer(i)
    sckServer(i).Accept requestID
    lstUser.AddItem sckServer(i).RemoteHostIP
    sckServer(i).SendData "YTEMSCommand:Login Success!"
End Sub

Private Sub sckServer_Close(Index As Integer)
    
End Sub

