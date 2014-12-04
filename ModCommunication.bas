Attribute VB_Name = "ModCommunication"
Option Explicit

'例：CS 代表 C->S MSG代表Message REQUEST_LOGIN代表此条常量的用途：请求登陆
Public Const CS_MSG_STU_REQUEST_LOGIN As Byte = &H1
Public Const CS_MSG_TEACHER_REQUEST_LOGIN As Byte = &H2
Public Const SC_MSG_LOGIN_SUCCESS As Byte = &H3
Public Const SC_MSG_LOGIN_FAILED As Byte = &H4
Public Const CS_MSG_REQUEST_STUDENT_MORE_INFORMATION As Byte = &H5
Public Const SC_MSG_STUDENT_MORE_INFORMATION As Byte = &H6
Public Const CS_MSG_FILE_TRANSFER As Byte = &H8
Public Const SC_MSG_FILE_TRANSFER As Byte = &H9
Public Const CS_MSG_FILE_DATA As Byte = &HA
Public Const SC_MSG_FILE_DATA As Byte = &HB
Public Const CS_MSG_SET_PASSWORD As Byte = &HC
Public Const SC_MSG_SET_PASSWORD_SUCCESS As Byte = &HD
Public Const SC_MSG_SET_PASSWORD_FAILED As Byte = &HE
Public Const CS_MSG_REQUEST_ENTER_EXAM As Byte = &HF
Public Const SC_MSG_ALLOW_ENTER_EXAM As Byte = &H10
Public Const SC_MSG_NOT_ALLOW_ENTER_EXAM As Byte = &H11
Public Const SC_MSG_TEACHER_LOGIN_SUCCESS As Byte = &H12
Public Const SC_MSG_TEACHER_LOGIN_FAILED As Byte = &H13
Public Const CS_MSG_TEACHER_SET_PASSWORD As Byte = &H14
Public Const SC_MSG_TEACHER_SET_PASSWORD_SUCCESS As Byte = &H15
Public Const SC_MSG_TEACHER_SET_PASSWORD_FAILED As Byte = &H16
Public Const CS_MSG_CHECK_EXAM_INFO As Byte = &H17
Public Const SC_MSG_EXAM_INFO_VALID As Byte = &H18
Public Const SC_MSG_EXAM_INFO_INVALID As Byte = &H19
Public Const CS_MSG_ADD_EXAM As Byte = &H20
