VERSION 5.00
Object = "{DA729162-C84F-11D4-A9EA-00A0C9199875}#1.60#0"; "mpqctl.ocx"
Begin VB.Form frmTest 
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
      Left            =   360
      Top             =   360
      _Version        =   65542
      _ExtentX        =   450
      _ExtentY        =   661
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'CompressTestFile "C:\Users\Jcr\Desktop\1.ini", "C:\Users\Jcr\Desktop\1.doc", "C:\Users\Jcr\Desktop\Exam.bin", a
    'UnCompressTestFile "C:\Users\Jcr\Desktop\Exam.bin", "C:\Users\Jcr\Desktop", a
    Dim t As TestInformation
    Dim a() As ChoiceConfig
    ReadTestConfigFile "C:\Users\Jcr\Desktop\Config1.ini", t, a
    
    t.AnswerCount = 10
    t.AnswerScore = 50
    t.ChoiceCount = 3
    t.ChoiceScore = 20
    t.DateTime = "2014/12/2 14:30:00"
    t.ExamTime = 120
    t.FillBlankCount = 10
    t.FillBlankScore = 20
    t.Subject = "大学英语"
    t.SubjectNo = "001"
    Dim b(2) As ChoiceConfig
    b(0).ChooseCount = 4
    b(0).MutiSelect = True
    b(0).TrueAnswer = "AB"
    b(1).ChooseCount = 6
    b(1).MutiSelect = True
    b(1).TrueAnswer = "AC"
    b(2).ChooseCount = 3
    b(2).MutiSelect = False
    b(2).TrueAnswer = "B"
    WriteTestConfigFile "C:\Users\Jcr\Desktop\Config1.ini", t, b
    
End Sub




