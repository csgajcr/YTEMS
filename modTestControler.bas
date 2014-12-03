Attribute VB_Name = "modTestControler"
Public Type TestInformation
    Subject As String
    SubjectNo As String
    DateTime As Date
    ExamTime As String
    ChoiceCount As Long                                                         '选择题数量
    FillBlankCount As Long                                                      '填空题数量
    AnswerCount As Long                                                         '解答题数量
    ChoiceScore As Long
    FillBlankScore As Long
    AnswerScore As Long
End Type
Public Type ChoiceConfig                                                        '单个选择题属性
    MutiSelect As Boolean                                                       '是否多选
    ChooseCount As Long                                                         '选项个数
    TrueAnswer As String
End Type

Public Function UnCompressTestFile(TestFilePath As String, DestPath As String, mpqctl As MpqControl)
    mpqctl.GetFile TestFilePath, "Config.ini", DestPath, True
    mpqctl.GetFile TestFilePath, "Exam.doc", DestPath, True
End Function

Public Function CompressTestFile(TestConfigFile As String, TestFile As String, DestFileName As String, mpqctl As MpqControl)
    mpqctl.AddFile DestFileName, TestConfigFile, "Config.ini", 1
    mpqctl.AddFile DestFileName, TestFile, "Exam.doc", 1
End Function

Public Function ReadTestConfigFile(ConfigPath As String, TestInfo As TestInformation, ChoiceCfg() As ChoiceConfig) As Boolean
    On Error GoTo myerr
    Dim i As Integer
    Dim sTmp() As String
    Dim sData As String
    TestInfo.AnswerCount = ReadFromINI("Test", "Answer", ConfigPath)
    TestInfo.AnswerScore = ReadFromINI("Test", "AnswerScore", ConfigPath)
    TestInfo.ChoiceCount = ReadFromINI("Test", "Choice", ConfigPath)
    TestInfo.ChoiceScore = ReadFromINI("Test", "ChoiceScore", ConfigPath)
    TestInfo.DateTime = ReadFromINI("Examinfo", "DateTime", ConfigPath)
    TestInfo.ExamTime = ReadFromINI("Examinfo", "ExamTime", ConfigPath)
    TestInfo.FillBlankCount = ReadFromINI("Test", "FillBlank", ConfigPath)
    TestInfo.FillBlankScore = ReadFromINI("Test", "FillBlankScore", ConfigPath)
    TestInfo.Subject = ReadFromINI("Examinfo", "Subject", ConfigPath)
    TestInfo.SubjectNo = ReadFromINI("Examinfo", "SubjectNo", ConfigPath)
    ReDim Preserve ChoiceCfg(TestInfo.ChoiceCount - 1)
    For i = 0 To TestInfo.ChoiceCount - 1
        sData = ""
        sData = ReadFromINI("Choice", "Choice" & CStr(i + 1), ConfigPath)
        sTmp = Split(sData, "|")
        ChoiceCfg(i).ChooseCount = CLng(sTmp(0))
        ChoiceCfg(i).MutiSelect = IIf(CLng(sTmp(1)) = 0, False, True)
        ChoiceCfg(i).TrueAnswer = sTmp(2)
    Next
    ReadTestConfigFile = True
    Exit Function
myerr:
    ReadTestConfigFile = False
End Function

Public Function WriteTestConfigFile(ConfigPath As String, TestInfo As TestInformation, ChoiceCfg() As ChoiceConfig) As Boolean
    On Error GoTo myerr
    Dim i As Integer
    Dim sTmp() As String
    Dim sData As String
    WriteToINI "Test", "Answer", TestInfo.AnswerCount, ConfigPath
    WriteToINI "Test", "AnswerScore", TestInfo.AnswerScore, ConfigPath
    WriteToINI "Test", "Choice", TestInfo.ChoiceCount, ConfigPath
    WriteToINI "Test", "ChoiceScore", TestInfo.ChoiceScore, ConfigPath
    WriteToINI "Examinfo", "DateTime", TestInfo.DateTime, ConfigPath
    WriteToINI "Examinfo", "ExamTime", TestInfo.ExamTime, ConfigPath
    WriteToINI "Test", "FillBlank", TestInfo.FillBlankCount, ConfigPath
    WriteToINI "Test", "FillBlankScore", TestInfo.FillBlankScore, ConfigPath
    WriteToINI "Examinfo", "Subject", TestInfo.Subject, ConfigPath
    WriteToINI "Examinfo", "SubjectNo", TestInfo.SubjectNo, ConfigPath
    For i = 0 To UBound(ChoiceCfg)
        WriteToINI "Choice", "Choice" & CStr(i + 1), ChoiceCfg(i).ChooseCount & "|" & IIf(ChoiceCfg(i).MutiSelect, "1", "0") & "|" & ChoiceCfg(i).TrueAnswer, ConfigPath
    Next
    
    WriteTestConfigFile = True
    Exit Function
myerr:
    WriteTestConfigFile = False
End Function
