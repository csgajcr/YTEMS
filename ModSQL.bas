Attribute VB_Name = "ModSQL"
Option Explicit
Public Function SQLQueryStudentInfo(TableName As String, UID As String, StuInfo As StudentInformation) As Boolean
    On Error GoTo myerr
    WaitForMysqlConnection
    Dim mysql_rs As New ADODB.Recordset
    mysql_rs.CursorLocation = adUseClient
    '--------
    mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "SELECT * FROM " & TableName & " WHERE StuNo = " & UID, mysql_conn
    
    StuInfo.ClassNo = mysql_rs(5)
    StuInfo.DeptNo = mysql_rs(4)
    StuInfo.S_JoinYear = mysql_rs(6)
    StuInfo.StuName = mysql_rs(1)
    StuInfo.StuPw = mysql_rs(3)
    StuInfo.StuSex = mysql_rs(2)
    StuInfo.UID = mysql_rs(0)
    mysql_rs.Close
    
    SQLQueryStudentInfo = True
    Exit Function
myerr:
    'MsgBox Err.Number & Err.Description
    SQLQueryStudentInfo = False
End Function
Public Function SQLQueryTeacherInfo(TableName As String, UID As String, TcInfo As TeacherInformation) As Boolean
    On Error GoTo myerr
    WaitForMysqlConnection
    Dim mysql_rs As New ADODB.Recordset
    mysql_rs.CursorLocation = adUseClient
    '--------
    mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "SELECT * FROM " & TableName & " WHERE TeaNo = " & UID, mysql_conn
    
    TcInfo.DeptNo = mysql_rs(4)
    TcInfo.JoinYear = mysql_rs(5)
    TcInfo.Password = mysql_rs(3)
    TcInfo.TeacherName = mysql_rs(1)
    TcInfo.TeacherSex = mysql_rs(2)
    TcInfo.UID = mysql_rs(0)
    mysql_rs.Close
    SQLQueryTeacherInfo = True
    Exit Function
myerr:
    'MsgBox Err.Number & Err.Description
    SQLQueryTeacherInfo = False
End Function
Public Function SQLQueryStudentMoreInfo(ClassTableName As String, DeptTableName As String, ClassNo As String, DeptNo As String, StuMoreInfo As StudentMoreInfo) As Boolean
    On Error GoTo myerr
    WaitForMysqlConnection
    Dim mysql_rs As New ADODB.Recordset
    mysql_rs.CursorLocation = adUseClient
    '-------------
    
    mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "SELECT * FROM " & ClassTableName & " WHERE ClassNo = " & ClassNo, mysql_conn
    StuMoreInfo.ClassDtor = mysql_rs(3)
    StuMoreInfo.ClassName = mysql_rs(1)
    mysql_rs.Close
    mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "SELECT * FROM " & DeptTableName & " WHERE DeptNo = " & DeptNo, mysql_conn
    StuMoreInfo.DeptDtor = mysql_rs(3)
    StuMoreInfo.Dept = mysql_rs(1)
    mysql_rs.Close
    SQLQueryStudentMoreInfo = True
    Exit Function
myerr:
    'MsgBox Err.Number & Err.Description
    SQLQueryStudentMoreInfo = False
End Function
Public Function SQLSetStudentPassword(TableName As String, UID As String, newPassword As String) As Boolean
    On Error GoTo myerr
    WaitForMysqlConnection
    Dim mysql_rs As New ADODB.Recordset
    mysql_rs.CursorLocation = adUseClient
    '--------------------------
    mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "UPDATE " & TableName & " SET StuPw=" & AddQueto(newPassword) & " WHERE StuNo=" & UID, mysql_conn
    
    SQLSetStudentPassword = True
    Exit Function
myerr:
    'MsgBox Err.Number & Err.Description
    SQLSetStudentPassword = False
End Function
Public Function SQLQueryExamInformation(ManageTableName As String, InfoTableName As String, ClassNo As String, ExamInfo() As ExamInformation) As Boolean
    On Error GoTo myerr
    WaitForMysqlConnection
    Dim mysql_rs As New ADODB.Recordset
    mysql_rs.CursorLocation = adUseClient
    '--------------------------
    mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "SELECT * FROM " & ManageTableName, mysql_conn                '& " WHERE ClassNo = " & AddQueto(ClassNo), mysql_conn
    Dim i As Long
    If mysql_rs.RecordCount > 0 Then
        For i = 1 To mysql_rs.RecordCount
            ReDim Preserve ExamInfo(i - 1)
            ExamInfo(i - 1).ExamID = mysql_rs(0)
            ExamInfo(i - 1).ExamTime = mysql_rs(2)
            ExamInfo(i - 1).ExamName = mysql_rs(1)
            mysql_rs.MoveNext
        Next                                                                    '
        mysql_rs.Close
        For i = 0 To UBound(ExamInfo)
            mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
            mysql_rs.Open "SELECT * FROM " & InfoTableName & " WHERE SubjectNo = " & AddQueto(ExamInfo(i).ExamID), mysql_conn
            ExamInfo(i).ExamDataTime = mysql_rs(1)
            mysql_rs.Close
        Next
    Else
        SQLQueryExamInformation = False
        Exit Function
    End If
    
    SQLQueryExamInformation = True
    Exit Function
myerr:
    'MsgBox Err.Number & Err.Description
    SQLQueryExamInformation = False
End Function

Public Function SQLQueryStudentScore(ScoreTableName As String, StuNo As String, SubjectNo As String, ret_StuScore As Long) As Boolean
    On Error GoTo myerr
    WaitForMysqlConnection
    Dim mysql_rs As New ADODB.Recordset
    mysql_rs.CursorLocation = adUseClient
    '--------------------------
    mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "SELECT * FROM " & ScoreTableName & " WHERE SubjectNo = " & AddQueto(SubjectNo) & " AND " & "StuNo = " & AddQueto(StuNo), mysql_conn
    ret_StuScore = mysql_rs(3)
    mysql_rs.Close
    SQLQueryStudentScore = True
    Exit Function
myerr:
    'MsgBox Err.Number & Err.Description
    SQLQueryStudentScore = False
End Function
Public Function SQLSetTeacherPassword(TableName As String, UID As String, newPassword As String) As Boolean
    On Error GoTo myerr
    WaitForMysqlConnection
    Dim mysql_rs As New ADODB.Recordset
    mysql_rs.CursorLocation = adUseClient
    '--------------------------
    mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "UPDATE " & TableName & " SET TeaPw=" & AddQueto(newPassword) & " WHERE TeaNo=" & UID, mysql_conn
    
    SQLSetTeacherPassword = True
    Exit Function
myerr:
    'MsgBox Err.Number & Err.Description
    SQLSetTeacherPassword = False
End Function

Public Function SQLAddExamInformation(ManageTableName As String, InfoTableName As String, ExamInfo As ExamInformation) As Boolean
    On Error GoTo myerr
    WaitForMysqlConnection
    Dim mysql_rs As New ADODB.Recordset
    mysql_rs.CursorLocation = adUseClient
    '--------------------------
    mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "INSERT INTO " & ManageTableName & " (`SubjectNo`, `SubjectName`, `ExamTime`, `ClassNo`) VALUES (" & AddQueto(ExamInfo.ExamID) & "," & AddQueto(ExamInfo.ExamName) & "," & AddQueto(ExamInfo.ExamTime) & "," & AddQueto("00000003") & ");", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "INSERT INTO " & InfoTableName & "(`SubjectNo`, `ExamDate`, `ExamRoom`) VALUES (" & AddQueto(ExamInfo.ExamID) & "," & AddQueto(ExamInfo.ExamDataTime) & "," & AddQueto("2201") & ");", mysql_conn, adOpenKeyset, adLockPessimistic
    SQLAddExamInformation = True
    Exit Function
myerr:
    MsgBox Err.Number & Err.Description
    SQLAddExamInformation = False
End Function

Public Function SQLDeleteExamInformation(ManageTableName As String, InfoTableName As String, ExamID As String) As Boolean
    On Error GoTo myerr
    WaitForMysqlConnection
    Dim mysql_rs As New ADODB.Recordset
    mysql_rs.CursorLocation = adUseClient
    '--------------------------
    mysql_rs.Open "SET NAMES GBK", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "DELETE FROM " & ManageTableName & " WHERE (`SubjectNo`='" & ExamID & "')", mysql_conn, adOpenKeyset, adLockPessimistic
    mysql_rs.Open "DELETE FROM " & InfoTableName & " WHERE (`SubjectNo`='" & ExamID & "')", mysql_conn, adOpenKeyset, adLockPessimistic
    SQLDeleteExamInformation = True
    Exit Function
myerr:
    MsgBox Err.Number & Err.Description
    SQLDeleteExamInformation = False
    
End Function
Public Function SQLAddStudentInformation()
    
End Function




