Attribute VB_Name = "ModSQL"
Option Explicit
Public Function SQLQueryStudentInfo(TableName As String, UID As String, StuInfo As StudentInformation) As Boolean
    On Error GoTo myerr
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
    MsgBox Err.Number & Err.Description
    SQLQueryStudentInfo = False
End Function
Public Function SQLQueryStudentMoreInfo(ClassTableName As String, DeptTableName As String, ClassNo As String, DeptNo As String, StuMoreInfo As StudentMoreInfo) As Boolean
    On Error GoTo myerr
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
    MsgBox Err.Number & Err.Description
    SQLQueryStudentMoreInfo = False
End Function
