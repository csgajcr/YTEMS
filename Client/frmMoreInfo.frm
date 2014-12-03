VERSION 5.00
Begin VB.Form frmMoreInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "¸ü¶àÐÅÏ¢"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3180
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame1 
      Caption         =   "¸ü¶àÐÅÏ¢"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.TextBox txtDeptDtor 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1815
         Width           =   1335
      End
      Begin VB.TextBox txtClassDtor 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   855
         Width           =   1335
      End
      Begin VB.TextBox txtDept 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1335
         Width           =   1335
      End
      Begin VB.TextBox txtClassName 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   375
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ÏµÖ÷ÈÎ£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "¸¨µ¼Ô±£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ôº   Ïµ£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "°à   ¼¶£º"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmMoreInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtClassName.BackColor = &HB6B6B6
    txtClassDtor.BackColor = &HB6B6B6
    txtDept.BackColor = &HB6B6B6
    txtDeptDtor.BackColor = &HB6B6B6
    'frmLoading.Show
    
    'sCommand = "YTEMSClientCommand:GetMoreInformation:" & StuInfo.ClassNo & "|" & StuInfo.DeptNo
    'frmLogin.sckClient.SendData sCommand
    frmLogin.sckClient.SendData CS_MSG_REQUEST_STUDENT_MORE_INFORMATION
    frmLogin.sckClient.SendData StuInfo.ClassNo & "|" & StuInfo.DeptNo
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

