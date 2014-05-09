VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "配置"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3420
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtPort 
      Height          =   270
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtIP 
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "端    口："
      Height          =   180
      Left            =   165
      TabIndex        =   3
      Top             =   975
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "服务器IP："
      Height          =   180
      Left            =   165
      TabIndex        =   1
      Top             =   375
      Width           =   900
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo myerr
    If txtIP.Text = "" Or txtPort.Text = "" Then Exit Sub
    YTEMSServerIP = txtIP.Text
    YTEMSServerPort = CLng(txtPort.Text)
    SaveServerIP ConfigPath, txtIP.Text
    SaveServerPort ConfigPath, CLng(txtPort.Text)
    MsgBox "设置完成", vbInformation
    Unload Me
myerr:
    If Err.Number = 13 Then
        MsgBox "请正确输入端口号", vbCritical
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    txtIP.Text = YTEMSServerIP
    txtPort.Text = YTEMSServerPort
End Sub
