VERSION 5.00
Begin VB.Form frmSetPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�޸�����"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3225
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdExit 
      Caption         =   "����"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSetPassword 
      Caption         =   "�޸�"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtNew2 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtNew 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtOld 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ȷ�����룺"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Top             =   1665
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�����룺"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   945
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ԭ���룺"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   225
      Width           =   960
   End
End
Attribute VB_Name = "frmSetPassword"
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
        MsgBox "����������", vbCritical
        Exit Sub
    ElseIf txtNew.Text <> txtNew2.Text Then
        MsgBox "�����������벻��ȷ", vbCritical
        Exit Sub
    ElseIf Left(MD5(txtOld.Text), 24) <> Left(StuInfo.StuPw, 24) Then
        MsgBox "ԭ���벻��ȷ", vbCritical
        Exit Sub
    ElseIf InStr(1, txtNew.Text, "|") Then
        MsgBox "�����ں��зǷ��ַ�", vbCritical
        Exit Sub
    End If
    
    'sCommand = "YTEMSClientCommand:ChangePassword:" & StuInfo.UID & "|" & MD5(txtNew.Text)
    'frmLogin.sckClient.SendData sCommand
    frmLogin.sckClient.SendData CS_MSG_SET_PASSWORD
    frmLogin.sckClient.SendData StuInfo.UID & "|" & MD5(txtNew.Text)
    NewPassword = MD5(txtNew.Text)
    Unload Me
    
End Sub

