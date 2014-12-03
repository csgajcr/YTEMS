VERSION 5.00
Object = "{DA729162-C84F-11D4-A9EA-00A0C9199875}#1.60#0"; "mpqctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin MPQCONTROLLib.MpqControl a 
      Left            =   120
      Top             =   120
      _Version        =   65542
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    'MsgBox GetPathFromFileName("C:\1dsad\asdasd\111.txt")
    Dim b As AnswerCardInformation
    b.AnswerCount = 3
    b.ChoiceCount = 5
    b.DateTime = "2014/12/4 12:30:00"
    b.ExamTime = 120
    b.FillBlankCount = 5
    b.Subject = "大学英语"
    b.SubjectNo = "001"
    Dim c(4) As String, cc() As String
    Dim f(4) As String, ff() As String
    Dim aa(2) As String, aaa() As String
    Dim d() As String
    c(0) = "A"
    c(1) = "AC"
    c(2) = "AD"
    c(3) = "AE"
    c(4) = "AF"
    aa(0) = "sadhaskjdhasjhda  shdkjashhdashdkjashdkash" & vbCrLf & "asdajskhdkjashdhasjhdkjh"
    aa(1) = "11111111111111"
    aa(2) = "11111222222"
    'WriteAnswerCardFile "C:\Users\Jcr\Desktop\1.ini", b, c, f, aa, d
    'CompressAnswerFile "C:\Users\Jcr\Desktop\1.ini", d, "C:\Users\Jcr\Desktop\1.asr", a
    UnCompressAnswerFile "C:\Users\Jcr\Desktop\1.asr", "C:\Users\Jcr\Desktop\", a
    ReadAnswerCardFile "C:\Users\Jcr\Desktop\Answer.ini", b, cc, ff, aaa
End Sub

