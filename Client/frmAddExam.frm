VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.ocx"
Begin VB.Form frmAddExam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ӿ���"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9015
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame2 
      Caption         =   "������ϸ��Ϣ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   9015
      Begin VB.TextBox txtChoiceCount 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��"
         Height          =   375
         Left            =   2520
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox cboChoice 
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox chkMutiSelect 
         Caption         =   "�Ƿ��ѡ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtTrueAnswer 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   7440
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdSaveChoice 
         Caption         =   "����"
         Height          =   375
         Left            =   8160
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtChooseCount 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   5760
         TabIndex        =   21
         Top             =   225
         Width           =   855
      End
      Begin VB.TextBox txtFillBlankCount 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   20
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtAnswerCount 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3960
         TabIndex        =   19
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtChoiceScore 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         TabIndex        =   18
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtFillBlankScore 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         TabIndex        =   17
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtAnswerScore 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2160
         TabIndex        =   16
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "�ύ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   14
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ѡ���������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   38
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   37
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "�𰸣�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   36
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "ѡ�������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   35
         Top             =   240
         Width           =   900
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8880
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "����������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   1440
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "����������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   33
         Top             =   1080
         Width           =   1440
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   8880
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "����ѡ���������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   1920
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "��������������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   1920
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "��������������"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   30
         Top             =   2640
         Width           =   1920
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   8880
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   29
         Top             =   3720
         Width           =   7815
      End
      Begin VB.Label Label16 
         Caption         =   "��ǰ״̬��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   3720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "���������Ϣ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.TextBox txtSubject 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtSubjectNo 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtTestPath 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   4815
      End
      Begin VB.CommandButton cmdbrowse 
         Caption         =   "���"
         Height          =   375
         Left            =   7800
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtExamDate 
         Height          =   375
         Left            =   5520
         TabIndex        =   3
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   9502721
         CurrentDate     =   41976
      End
      Begin MSComCtl2.DTPicker dtExamDate2 
         Height          =   375
         Left            =   7200
         TabIndex        =   7
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   9502722
         CurrentDate     =   41976
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ѧ   �ƣ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "���Ժţ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ʱ   �䣺"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4560
         TabIndex        =   10
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ʱ   ����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4560
         TabIndex        =   9
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "����(�����word�ļ�)��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   2745
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8280
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAddExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboChoice_Change()
    txtTrueAnswer.Text = ""
End Sub

Private Sub cmdbrowse_Click()
    cd.Filter = "*.doc|*.doc|*.*|*.*"
    cd.FileName = ""
    cd.ShowOpen
    If cd.CancelError Or cd.FileName = "" Then Exit Sub
    txtTestPath.Text = cd.FileName
End Sub

Private Sub cmdOK_Click()
    On Error GoTo myerr
    Dim i As Integer
    For i = 1 To CLng(txtChoiceCount.Text)
        cboChoice.AddItem i
    Next
    ReDim Preserve ChoiceCfg(CLng(txtChoiceCount.Text))
    Exit Sub
myerr:
    MsgBox "��������ȷ����ֵ", vbCritical
End Sub

Private Sub cmdSaveChoice_Click()
    On Error GoTo myerr
    ChoiceCfg(CLng(cboChoice.Text)).MutiSelect = chkMutiSelect.Value
    ChoiceCfg(CLng(cboChoice.Text)).TrueAnswer = txtTrueAnswer.Text
    ChoiceCfg(CLng(cboChoice.Text)).ChooseCount = CLng(txtChooseCount.Text)
    lblStatus.Caption = "����ɹ�"
    Exit Sub
myerr:
    MsgBox "��������ȷ����ֵ", vbCritical
End Sub

Private Sub cmdSubmit_Click()

End Sub

Private Sub Form_Load()
    
End Sub

