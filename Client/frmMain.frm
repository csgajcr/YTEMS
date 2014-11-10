VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "移通考试系统 客户端"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   5145
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame FraUserInformation 
      Caption         =   "个人信息"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.CommandButton Command1 
         Caption         =   "修改登录密码"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdMoreInfo 
         Caption         =   "查看详细信息"
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Frame FraExaminfo 
         Caption         =   "考试信息"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   4455
         Begin MSComctlLib.ListView lstExamInformation 
            Height          =   3255
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   5741
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "科目"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "考试号"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "时间"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "时长(分)"
               Object.Width           =   1411
            EndProperty
         End
         Begin VB.Image cmdEnterExam 
            Height          =   1575
            Left            =   2400
            Picture         =   "frmMain.frx":0000
            Stretch         =   -1  'True
            Top             =   3720
            Width           =   1815
         End
      End
      Begin VB.TextBox txtJoinYear 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1455
         Width           =   1095
      End
      Begin VB.TextBox txtClassNo 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1815
         Width           =   1095
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtUID 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   371
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "年   级："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "班级号："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "性   别："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "学   号："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓   名："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   945
      End
      Begin VB.Image imgHead 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   240
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnterExam_Click()
    Dim sCommand As String * 100
    Dim msg As Long
    If lstExamInformation.SelectedItem.Index = -1 Then
        MsgBox "请选择要进入的考试", vbCritical
        Exit Sub
    End If
    msg = MsgBox("即将进入" & lstExamInformation.SelectedItem.Text & "科目的考试，确认进入？", vbYesNo + vbInformation)
    If msg = vbYes Then
        
        sCommand = "YTEMSClientCommand:EnterExam:" & lstExamInformation.SelectedItem.SubItems(1) & "|" & lstExamInformation.SelectedItem.SubItems(2) & "|" & lstExamInformation.SelectedItem.SubItems(3)
        frmLogin.sckClient.SendData sCommand
        Unload frmLoading
    Else
        Exit Sub
    End If
    
End Sub

Private Sub cmdEnterExam_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        
        
    End If
End Sub

Private Sub cmdEnterExam_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then

    End If
End Sub

Private Sub cmdMoreInfo_Click()
    frmMoreInfo.Show 1
End Sub

Private Sub Command1_Click()
    frmSetPassword.Show 1
End Sub

Private Sub Form_Load()
    
    txtUserName.BackColor = &HB6B6B6
    txtSex.BackColor = &HB6B6B6
    txtUID.BackColor = &HB6B6B6
    txtJoinYear.BackColor = &HB6B6B6
    txtClassNo.BackColor = &HB6B6B6
    lstExamInformation.BackColor = &HB6B6B6
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmLogin.Visible = False Then
        End
    End If
End Sub



Private Sub FraUserInformation_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
