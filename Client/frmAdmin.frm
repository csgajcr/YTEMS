VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "管理员/教师 客户端"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9045
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame3 
      Caption         =   "考试管理"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   4560
      TabIndex        =   9
      Top             =   960
      Width           =   4455
      Begin VB.CommandButton cmdDelExam 
         Caption         =   "删除考试"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdGradeTest 
         Caption         =   "批改试题"
         Height          =   495
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdEditExam 
         Caption         =   "修改考试"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddExam 
         Caption         =   "添加考试"
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin MSComctlLib.ListView lstExamInformation 
         Height          =   2175
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3836
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "信息管理"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "添加学生"
         Height          =   495
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddStudent 
         Caption         =   "添加学生"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "个人信息"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.CommandButton cmdSetPassword 
         Caption         =   "修改密码"
         Height          =   375
         Left            =   7560
         TabIndex        =   7
         Top             =   360
         Width           =   1335
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "年级："
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
         Left            =   4800
         TabIndex        =   3
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "工号："
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
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名："
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddExam_Click()
    frmAddExam.Show 1
End Sub

Private Sub cmdSetPassword_Click()
    frmAdminSetPW.Show 1
End Sub

Private Sub Form_Load()
    txtUserName.BackColor = &HB6B6B6
    txtUID.BackColor = &HB6B6B6
    txtJoinYear.BackColor = &HB6B6B6
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmLogin.Visible = False Then
        End
    End If
End Sub

