Attribute VB_Name = "ModAES"
Option Explicit
#Const SUPPORT_LEVEL = 0
Public Function AESEncodeStr(ByVal str As String, ByVal password As String, ByVal bits As Integer) As String
    Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long
    Dim m_Rijndael As New cRijndael
    str = StrConv(str, vbUnicode)
    password = StrConv(password, vbUnicode)
    If Len(str) = 0 Then
        AESEncodeStr = ""
    Else
        If Len(password) = 0 Then
            AESEncodeStr = ""
        Else
            KeyBits = bits
            BlockBits = 128
            pass = GetPassword(password, bits)
            
            If HexDisplayRev(str, plaintext) = 0 Then
                AESEncodeStr = ""
                Exit Function
            End If
            #If SUPPORT_LEVEL Then
                m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
                m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0, BlockBits
            #Else
                m_Rijndael.SetCipherKey pass, KeyBits
                m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0
            #End If
            
            AESEncodeStr = StrConv(HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8), vbFromUnicode)
            MsgBox HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
        End If
    End If
    Set m_Rijndael = Nothing
End Function
Public Function AESDecodeStr(ByVal str As String, ByVal password As String, ByVal bits As Integer) As String
    Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long
    Dim m_Rijndael As New cRijndael
    str = StrConv(str, vbUnicode)
    password = StrConv(password, vbUnicode)
    If Len(str) = 0 Then
        AESDecodeStr = ""
    Else
        If Len(password) = 0 Then
            AESDecodeStr = ""
        Else
            KeyBits = bits
            BlockBits = 128
            pass = GetPassword(password, bits)
            
            If HexDisplayRev(str, ciphertext) = 0 Then
                AESDecodeStr = ""
                Exit Function
            End If
            #If SUPPORT_LEVEL Then
                m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
                If m_Rijndael.ArrayDecrypt(plaintext, ciphertext, 0, BlockBits) <> 0 Then
                    Exit Function
                End If
            #Else
                m_Rijndael.SetCipherKey pass, KeyBits
                If m_Rijndael.ArrayDecrypt(plaintext, ciphertext, 0) <> 0 Then
                    Exit Function
                End If
            #End If
            
            AESDecodeStr = StrConv(plaintext, vbUnicode)
            
        End If
    End If
    Set m_Rijndael = Nothing
End Function
Private Function HexDisplayRev(TheString As String, data() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim d As Long
    Dim n As Long
    Dim data2() As Byte
    
    n = 2 * Len(TheString)
    data2 = TheString
    
    ReDim data(n \ 4 - 1)
    
    d = 0
    i = 0
    j = 0
    Do While j < n
        c = data2(j)
        Select Case c
        Case 48 To 57                                                           '"0" ... "9"
            If d = 0 Then                                                       'high
                d = c
            Else                                                                'low
                data(i) = (c - 48) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 65 To 70                                                           '"A" ... "F"
            If d = 0 Then                                                       'high
                d = c - 7
            Else                                                                'low
                data(i) = (c - 55) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 97 To 102                                                          '"a" ... "f"
            If d = 0 Then                                                       'high
                d = c - 39
            Else                                                                'low
                data(i) = (c - 87) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        End Select
        j = j + 2
    Loop
    n = i
    If n = 0 Then
        Erase data
    Else
        ReDim Preserve data(n - 1)
    End If
    HexDisplayRev = n
    
End Function
Private Function HexDisplay(data() As Byte, n As Long, k As Long) As String
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim data2() As Byte
    
    If LBound(data) = 0 Then
        ReDim data2(n * 4 - 1 + ((n - 1) \ k) * 4)
        j = 0
        For i = 0 To n - 1
            If i Mod k = 0 Then
                If i <> 0 Then
                    data2(j) = 32
                    data2(j + 2) = 32
                    j = j + 4
                End If
            End If
            c = data(i) \ 16&
            If c < 10 Then
                data2(j) = c + 48                                               ' "0"..."9"
            Else
                data2(j) = c + 55                                               ' "A"..."F"
            End If
            c = data(i) And 15&
            If c < 10 Then
                data2(j + 2) = c + 48                                           ' "0"..."9"
            Else
                data2(j + 2) = c + 55                                           ' "A"..."F"
            End If
            j = j + 4
        Next i
        Debug.Assert j = UBound(data2) + 1
        HexDisplay = data2
    End If
    
End Function
Private Function GetPassword(password As String, ByVal bits As Integer) As Byte()
    Dim data() As Byte
    
    
    If HexDisplayRev(password, data) <> (bits \ 8) Then
        data = StrConv(password, vbFromUnicode)
        ReDim Preserve data(31)
    End If
    GetPassword = data
End Function


