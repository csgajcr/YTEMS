VERSION 5.00
Begin VB.Form frmLoading 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   120
   End
   Begin VB.Label Label1 
      Caption         =   "Loading....."
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Timer1_Timer()
    Select Case Label1.Caption
    Case "Loading"
        Label1.Caption = "Loading."
    Case "Loading."
        Label1.Caption = "Loading.."
    Case "Loading.."
        Label1.Caption = "Loading..."
    Case "Loading..."
        Label1.Caption = "Loading...."
    Case "Loading...."
        Label1.Caption = "Loading....."
    Case "Loading....."
        Label1.Caption = "Loading"
    End Select
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub
