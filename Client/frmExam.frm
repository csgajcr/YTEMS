VERSION 5.00
Begin VB.Form frmExam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "正在考试..."
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   10935
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   2535
      Left            =   6960
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3615
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Me.Width = Screen.Width
    'Me.Height = Screen.Height
    'Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2, Screen.Width, Screen.Height
    'Me.WindowState = 2
End Sub
