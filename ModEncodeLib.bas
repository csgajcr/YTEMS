Attribute VB_Name = "ModEncodeLib"
Option Explicit

Public Declare Function MD5Encode Lib "Encode.dll" (ByVal sstring As String, ByVal bbyte As Integer) As String
Public Declare Function AESEncodeStr Lib "Encode.dll" (ByVal sStr As String, ByVal sPassword As String, ByVal byt1 As Integer, ByVal byt2 As Integer) As String
Public Declare Function AESDecodeStr Lib "Encode.dll" (ByVal sHashText As String, ByVal sPassword As String, ByVal byt1 As Integer, ByVal byt2 As Integer) As String
Public Declare Function AESEncodeFile Lib "Encode.dll" (ByVal SrcFileName As String, ByVal HashFileName As String, ByVal sPassword As String, ByVal byt1 As Integer, ByVal byt2 As Integer) As Boolean
Public Declare Function AESDecodeFile Lib "Encode.dll" (ByVal EncryptFileName As String, ByVal DecryptFileName As String, ByVal sPassword As String, ByVal byt1 As Integer, ByVal byt2 As Integer) As Boolean
Public Declare Function DESEncodeStr Lib "Encode.dll" (ByVal sStr As String, ByVal sPassword As String) As String
Public Declare Function DESDecodeStr Lib "Encode.dll" (ByVal sStr As String, ByVal sPassword As String) As String
Public Declare Function DESDecodeFile Lib "Encode.dll" (ByVal SourceFile As String, ByVal OutFile As String, ByVal sPassword As String)
Public Declare Function DESEncodeFile Lib "Encode.dll" (ByVal SourceFile As String, ByVal OutFile As String, ByVal sPassword As String)
Public Declare Function RSAEncodeStr Lib "Encode.dll" (ByVal sStr As String) As String
Public Declare Function RSADecodeStr Lib "Encode.dll" (ByVal sStr As String) As String
Public Declare Function ShA1EncodeStr Lib "Encode.dll" (ByVal sStr As String) As String
Public Declare Function CalcFileSHA1 Lib "Encode.dll" (ByVal sFilePath As String) As String

