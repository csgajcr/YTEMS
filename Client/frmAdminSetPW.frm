VERSION 5.00
Begin VB.Form frmAdminSetPW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "–ﬁ∏ƒ√‹¬Î"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3270
   StartUpPosition =   2  '∆¡ƒª÷––ƒ
   Begin VB.TextBox txtOld 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   135
      Width           =   1335
   End
   Begin VB.TextBox txtNew 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   855
      Width           =   1335
   End
   Begin VB.TextBox txtNew2 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1575
      Width           =   1335
   End
   Begin VB.CommandButton cmdSetPassword 
      Caption         =   "–ﬁ∏ƒ"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2295
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "∑µªÿ"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2295
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "‘≠√‹¬Î£∫"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "–¬√‹¬Î£∫"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "»∑»œ√‹¬Î£∫"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1320
   End
End
Attribute VB_Name = "frmAdminSetPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSetPassword_Click()
    If txtOld.Text = "" Or txtNew.Text = "" Or txtNew2.Text = "" Then
        MsgBox "«Î ‰»Î√‹¬Î", vbCritical
        Exit Sub
    ElseIf txtNew.Text <> txtNew2.Text Then
        MsgBox "¡Ω¥Œ ‰»Î√‹¬Î≤ª’˝»∑", vbCritical
        Exit Sub
    ElseIf Left(MD5(txtOld.Text), 24) <> Left(TcInfo.Password, 24) Then
        MsgBox "‘≠√‹¬Î≤ª’˝»∑", vbCritical
        Exit Sub
    ElseIf InStr(1, txtNew.Text, "|") Then
        MsgBox "√‹¬Îƒ⁄∫¨”–∑«∑®◊÷∑˚", vbCritical
        Exit Sub
    End If
    
    frmLogin.sckClient.SendData CS_MSG_TEACHER_SET_PASSWORD
    frmLogin.sckClient.SendData TcInfo.UID & "|" & MD5(txtNew.Text)
    NewPassword = MD5(txtNew.Text)
    Unload Me
End Sub

