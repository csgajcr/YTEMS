Attribute VB_Name = "ModDllMain"
'===================================================================================================
'| 模 块 名 | ModDllMain
'| 描    述 | Win32 Dynamic-Link Library's entry function "DllMain"
'| 说    明 | 在'dllMain'中除注释加函数外尽可能保持原代码\如需引用请保留模块原作者注释\构建时使用amicForvb插件
'| 创 建 人 | amicy QQ:35723195 MSN:xb@live.it
'| 日    期 | 2008-10-10 16:55:07
'| 版    本 | 1.0.0
'===================================================================================================
Option Explicit
Private Type tagGlobal
    Value     As Long
End Type
Dim ofac        As Object
Dim mGlobal     As tagGlobal
Const MEM_COMMIT = 4096
Const PAGE_EXECUTE_READWRITE = 64

Public Function DllMain(ByVal hinstDLL As Long, ByVal fdwReason As Long, ByVal lpvReserved As Long) As Long
    Dim Lib         As amicBase.ITypeLib
    Dim Info        As amicBase.ITypeInfo
    Dim sfile       As String
    Dim sClsName    As String
    Dim iFound      As Integer
    Dim riid        As amicBase.UUID
    Dim Attr        As amicBase.TYPEATTR
    Dim lpasmCode   As Long
    Dim i           As Long
    Dim lLibMSVBVM60    As Long
    Dim lpUserDllMain   As Long
    Dim lpDllGetObj     As Long
    
    Select Case fdwReason
    Case DLL_PROCESS_DETACH
        lpDllGetObj = mGlobal.Value
        '
        'Destroy add [function]=====>>>>>>>>>>>>>>>>>>>
        '
        GoTo VBDLLMAIN
    Case DLL_PROCESS_ATTACH
        
        sfile = Space$(MAX_PATH)
        amicBase.GetModuleFileName hinstDLL, sfile, MAX_PATH
        Set Lib = amicBase.LoadTypeLib(sfile)
        If Lib Is Nothing Then Exit Function
        iFound = 1
        sClsName = "CInitWinDll"
        Lib.FindName sClsName, 0&, Info, 0&, iFound
        If iFound Then
            amicBase.CopyMemory Attr, ByVal Info.GetTypeAttr, Len(Attr)
            With riid
                .Data1 = 1
                .Data4(0) = &HC0
                .Data4(7) = &H46
            End With
            
            lpDllGetObj = amicBase.GetProcAddress(hinstDLL, "DllGetClassObject")
            
            
VBDLLMAIN:
            lLibMSVBVM60 = amicBase.LoadLibrary("MSVBVM60.dll")
            lpUserDllMain = amicBase.GetProcAddress(lLibMSVBVM60, "UserDllMain")
            lpasmCode = VirtualAlloc(0, 4096, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
            FillMemory ByVal lpasmCode, 4096, &HCC
            amicBase.CopyMemory ByVal lpasmCode + 0, &H68, 1
            amicBase.CopyMemory ByVal lpasmCode + 1, lpvReserved, 4
            amicBase.CopyMemory ByVal lpasmCode + 5, &H68, 1
            amicBase.CopyMemory ByVal lpasmCode + 6, fdwReason, 4
            amicBase.CopyMemory ByVal lpasmCode + 10, &H68, 1
            amicBase.CopyMemory ByVal lpasmCode + 11, hinstDLL, 4
            amicBase.CopyMemory ByVal lpasmCode + 15, &H68, 1
            amicBase.CopyMemory ByVal lpasmCode + 16, ByVal lpDllGetObj + 7, 4
            amicBase.CopyMemory ByVal lpasmCode + 20, &H68, 1
            amicBase.CopyMemory ByVal lpasmCode + 21, ByVal lpDllGetObj + 12, 4
            amicBase.CopyMemory ByVal lpasmCode + 25, &HE8, 1
            amicBase.CopyMemory ByVal lpasmCode + 26, lpUserDllMain - (lpasmCode + 26) - 4, 4
            amicBase.CopyMemory ByVal lpasmCode + 30, &HC2, 1
            amicBase.CopyMemory ByVal lpasmCode + 31, &H10, 1
            amicBase.CopyMemory ByVal lpasmCode + 32, &H0, 1
            
            If fdwReason = DLL_PROCESS_DETACH Then GoTo DLLDETACH
            
            DllMain = CallWindowProc(lpasmCode, 0, 0, 0, 0)
            
            FillMemory ByVal lpasmCode, 4096, &HCC
            amicBase.CopyMemory ByVal lpasmCode + 0, &H68, 1
            amicBase.CopyMemory ByVal lpasmCode + 1, VarPtr(ofac), 4
            amicBase.CopyMemory ByVal lpasmCode + 5, &H68, 1
            amicBase.CopyMemory ByVal lpasmCode + 6, VarPtr(riid), 4
            amicBase.CopyMemory ByVal lpasmCode + 10, &H68, 1
            amicBase.CopyMemory ByVal lpasmCode + 11, VarPtr(Attr.IID), 4
            amicBase.CopyMemory ByVal lpasmCode + 15, &HE8, 1
            amicBase.CopyMemory ByVal lpasmCode + 16, lpDllGetObj - (lpasmCode + 16) - 4, 4
            amicBase.CopyMemory ByVal lpasmCode + 20, &HC2, 1
            amicBase.CopyMemory ByVal lpasmCode + 21, &H10, 1
            amicBase.CopyMemory ByVal lpasmCode + 22, &H0, 1
            If CallWindowProc(lpasmCode, 0, 0, 0, 0) = 0 Then
                mGlobal.Value = lpDllGetObj
                Set Info = Nothing
                Set Lib = Nothing
                'Initialization finished... add [function]=====>>>>>>>>>>>>>>>>>>>
            End If
        End If
        
    End Select
    Exit Function
DLLDETACH:
    amicBase.CopyMemory ByVal lpasmCode + 1, lpvReserved, 4
    amicBase.CopyMemory ByVal lpasmCode + 6, fdwReason, 4
    amicBase.CopyMemory ByVal lpasmCode + 11, hinstDLL, 4
    DllMain = CallWindowProc(lpasmCode, 0, 0, 0, 0)
    VirtualFree lpasmCode, 0, &H8000&
    Set ofac = Nothing
End Function



