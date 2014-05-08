VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "移通考试系统 服务端"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   8085
   StartUpPosition =   3  '窗口缺省
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   7440
      Top             =   480
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub InitiazationConfig()
    sckListen.Close
    sckListen.Bind GetPort()
    sckListen.Listen
End Sub

Private Sub Form_Load()
    InitiazationConfig
End Sub

Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
    sckServer(0).Accept requestID
    MsgBox "Accept Success!" & sckServer(0).RemoteHostIP
End Sub

Private Sub sckServer_Close(Index As Integer)
    sckServer(Index).Close
End Sub

