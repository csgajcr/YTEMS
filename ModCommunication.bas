Attribute VB_Name = "ModCommunication"
Option Explicit

'例：CS 代表 C->S MSG代表Message REQUEST_LOGIN代表此条常量的用途：请求登陆
Public Const CS_MSG_STU_REQUEST_LOGIN As Byte = &H1
Public Const CS_MSG_TEACHER_REQUEST_LOGIN As Byte = &H2
Public Const SC_MSG_LOGIN_SUCCESS As Byte = &H3
Public Const SC_MSG_LOGIN_FAILED As Byte = &H4
