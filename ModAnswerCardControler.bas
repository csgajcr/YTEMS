Attribute VB_Name = "ModAnswerCardControler"
Option Explicit
Public Type AnswerCardInformation
    Subject As String
    SubjectNo As String
    DateTime As Date
    ExamTime As String
    ChoiceCount As Long
    FillBlankCount As Long
    AnswerCount As Long
End Type
'AnswerFilePath:答题卡文件
'DestPath 目标文件夹
Public Function UnCompressAnswerFile(AnswerFilePath As String, DestPath As String, mpqctl As MpqControl)
    Dim sTmp As String, iCount As Integer, i As Integer
    mpqctl.GetFile AnswerFilePath, "Answer.ini", DestPath, 1
    iCount = GetSectionKeyCount("Answer", IIf(Right(DestPath, 1) = "\", DestPath, DestPath & "\") & "Answer.ini")
    For i = 1 To iCount
        mpqctl.GetFile AnswerFilePath, "Answer" & CStr(i) & ".txt", DestPath, 1
    Next
End Function
'AnswerINI:答题卡信息ini文件
'AnswerTXT():解答题对应的txt文件路径
'DestFile:目标答题卡文件
Public Function CompressAnswerFile(AnswerINI As String, AnswerTXT() As String, DestFile As String, mpqctl As MpqControl)
    Dim i As Integer
    mpqctl.AddFile DestFile, AnswerINI, "Answer.ini", 1
    Kill AnswerINI
    For i = 0 To UBound(AnswerTXT)
        mpqctl.AddFile DestFile, AnswerTXT(i), "Answer" & CStr(i + 1) & ".txt", 1
        Kill AnswerTXT(i)
    Next
    
End Function
'AnswerINI:答题卡信息ini文件
'out_AnswerTXT():解答题对应的txt文件路径
'AnsInfo:答题卡信息变量
'Choice(),FillBlank(),Answer():3种类型题目的解答
Public Function WriteAnswerCardFile(AnswerINI As String, AnsInfo As AnswerCardInformation, Choice() As String, FillBlank() As String, Answer() As String, out_AnswerTXT() As String) As Boolean
    On Error GoTo myerr
    Dim i As Integer
    
    
    WriteToINI "Examinfo", "Subject", AnsInfo.Subject, AnswerINI
    WriteToINI "Examinfo", "SubjectNo", AnsInfo.SubjectNo, AnswerINI
    WriteToINI "Examinfo", "DateTime", AnsInfo.DateTime, AnswerINI
    WriteToINI "Examinfo", "ExamTime", AnsInfo.ExamTime, AnswerINI
    If AnsInfo.ChoiceCount > 0 Then
        For i = 0 To UBound(Choice)
            WriteToINI "Choice", "Choice" & CStr(i + 1), Choice(i), AnswerINI
        Next
    End If
    If AnsInfo.FillBlankCount > 0 Then
        For i = 0 To UBound(FillBlank)
            WriteToINI "FillBlank", "FillBlank" & CStr(i + 1), FillBlank(i), AnswerINI
        Next
    End If
    Dim sPath As String
    sPath = GetPathFromFileName(AnswerINI)
    Dim Filenum As Integer
    If AnsInfo.AnswerCount > 0 Then
        ReDim Preserve out_AnswerTXT(UBound(Answer))
        For i = 0 To UBound(Answer)
            WriteToINI "Answer", "Answer" & CStr(i + 1), "Answer" & CStr(i + 1) & ".txt", AnswerINI
            Filenum = FreeFile
            
            Open sPath & "Answer" & CStr(i + 1) & ".txt" For Output As #Filenum
            out_AnswerTXT(i) = sPath & "Answer" & CStr(i + 1) & ".txt"
            Print #Filenum, Answer(i)
            Close #Filenum
        Next
    End If
    
    Exit Function
myerr:
    
    
End Function
'AnswerINI:答题卡信息ini文件
'AnsInfo:答题卡信息变量
'Choice(),FillBlank(),Answer():3种类型题目的解答
Public Function ReadAnswerCardFile(AnswerINI As String, AnsInfo As AnswerCardInformation, Choice() As String, FillBlank() As String, Answer() As String) As Boolean
    'On Error GoTo myerr
    Dim i As Integer
    Dim sData  As String
    AnsInfo.Subject = ReadFromINI("Examinfo", "Subject", AnswerINI)
    AnsInfo.SubjectNo = ReadFromINI("Examinfo", "SubjectNo", AnswerINI)
    AnsInfo.DateTime = ReadFromINI("Examinfo", "DateTime", AnswerINI)
    AnsInfo.ExamTime = ReadFromINI("Examinfo", "ExamTime", AnswerINI)
    AnsInfo.ChoiceCount = GetSectionKeyCount("Choice", AnswerINI)
    AnsInfo.FillBlankCount = GetSectionKeyCount("FillBlank", AnswerINI)
    AnsInfo.AnswerCount = GetSectionKeyCount("Answer", AnswerINI)
    ReDim Preserve Choice(AnsInfo.ChoiceCount)
    ReDim Preserve FillBlank(AnsInfo.FillBlankCount)
    ReDim Preserve Answer(AnsInfo.AnswerCount)
    For i = 1 To UBound(Choice)
        Choice(i - 1) = ReadFromINI("Choice", "Choice" & CStr(i), AnswerINI)
    Next
    
    For i = 1 To UBound(FillBlank)
        FillBlank(i - 1) = ReadFromINI("FillBlank", "Fillblank" & CStr(i), AnswerINI)
    Next
    Dim Filenum As Integer
    For i = 1 To UBound(Answer)
        Filenum = FreeFile
        Open GetPathFromFileName(AnswerINI) & "Answer" & CStr(i) & ".txt" For Input As #Filenum
        Do While EOF(Filenum) = False
            Line Input #Filenum, sData
            Answer(i - 1) = Answer(i - 1) & sData & vbCrLf
        Loop
        Close #Filenum
    Next
    
    Exit Function
myerr:
End Function


