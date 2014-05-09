VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "移通考试系统 客户端"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   5145
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame FraUserInformation 
      Caption         =   "个人信息"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.Image imgHead 
         Height          =   975
         Left            =   240
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    If frmLogin.Visible = False Then
        End
    End If
End Sub
