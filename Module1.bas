Attribute VB_Name = "ByPassHShield"
Private Declare Function OpenProcessAPI Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemoryAPI Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CloseHandleAPI Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const MEM_COMMIT = &H1000
Private Const PAGE_READWRITE = &H4



Public Function Crack_Hs() As Long
    Static BeUsed As Boolean
    If BeUsed = False Then
        Dim hProcess As Long
        Dim lpImagePath As String
        lpImagePath = "C:\WINDOWS\system32\taskmgr.exe"
        hProcess = OpenProcessAPI(PROCESS_ALL_ACCESS, 0, GetCurrentProcessId)
        If hProcess = 0 Then Exit Function
        Dim sLenth As Long
        Dim BaseAddress As Long
        sLenth = LenB(lpImagePath) + 1 + 26
        BaseAddress = VirtualAllocEx(hProcess, ByVal 0&, ByVal sLenth, MEM_COMMIT, PAGE_READWRITE)
        If BaseAddress = 0 Then Exit Function
        WriteProcessMemoryAPI hProcess, ByVal BaseAddress + 0, ByVal VarPtr(&H30058B64), 4, False
        WriteProcessMemoryAPI hProcess, ByVal BaseAddress + 4, ByVal VarPtr(&H8B000000), 4, False
        WriteProcessMemoryAPI hProcess, ByVal BaseAddress + 8, ByVal VarPtr(&HC0831040), 4, False
        WriteProcessMemoryAPI hProcess, ByVal BaseAddress + 12, ByVal VarPtr(&H245C8B3C), 4, False
        WriteProcessMemoryAPI hProcess, ByVal BaseAddress + 16, ByVal VarPtr(&H89188904), 4, False
        WriteProcessMemoryAPI hProcess, ByVal BaseAddress + 20, ByVal VarPtr(&HC2042444), 4, False
        WriteProcessMemoryAPI hProcess, ByVal BaseAddress + 24, ByVal VarPtr(&H10), 2, False
        WriteProcessMemoryAPI hProcess, ByVal BaseAddress + 26, ByVal StrPtr(lpImagePath), sLenth, False
        CloseHandleAPI hProcess
        CallWindowProc BaseAddress, BaseAddress + 26, 0, 0, 0
        BeUsed = True
        Crack_Hs = BaseAddress
    End If
End Function




