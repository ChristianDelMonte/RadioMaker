Attribute VB_Name = "modPriority"
'Global constants
Public Const PROCESS_PRIORITY_IDLE = 4
Public Const PROCESS_PRIORITY_NORMAL = 8
Public Const PROCESS_PRIORITY_HIGH = 13
Public Const PROCESS_PRIORITY_REALTIME = 24

Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100

Const HSHELL_ACTIVATESHELLWINDOW = 3
Const HSHELL_WINDOWCREATED = 1
Const HSHELL_WINDOWDESTROYED = 2
Const HSHELL_WINDOWACTIVATED = 4
Const HSHELL_GETMINRECT = 5
Const HSHELL_REDRAW = 6
Const HSHELL_TASKMAN = 7
Const HSHELL_LANGUAGE = 8
Const HSHELL_ACCESSIBILITYSTATE = 11
Const LOCALE_SENGLANGUAGE As Long = &H1001

'Public Constants
Public Const GWL_WNDPROC = (-4)
Public Const RSH_DEREGISTER = 0
Public Const RSH_REGISTER = 1
Public Const RSH_REGISTER_PROGMAN = 2
Public Const RSH_REGISTER_TASKMAN = 3
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_TERMINATE = &H1&
Public Const PROCESS_CREATE_THREAD = &H2&
Public Const PROCESS_VM_OPERATION = &H8&
Public Const PROCESS_VM_READ = &H10&
Public Const PROCESS_VM_WRITE = &H206
Public Const PROCESS_DUP_HANDLE = &H40&
Public Const PROCESS_CREATE_PROCESS = &H80&
Public Const PROCESS_SET_QUOTA = &H100&
Public Const PROCESS_SET_INFORMATION = &H200&
Public Const PROCESS_QUERY_INFORMATION = &H400&
Public Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

'Type Definitions
Type ProcessEntry
    dwSize As Long
    peUsage As Long
    peProcessID As Long
    peDefaultHeapID As Long
    peModuleID As Long
    peThreads As Long
    peParentProcessID As Long
    pePriority As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

'Local Variables
Dim hnd                             As Long
Dim lRet                            As Long
Dim lExitCode                       As Long
Dim lPriority                       As Long
Dim exePriority                     As Long

'Public Variables
Public OldProc                      As Long
Public uRegMsg                      As Long

'API Declarations
Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal dwIdProc As Long) As Long
Declare Function Process32First Lib "kernel32" (ByVal hndl As Long, ByRef pstru As ProcessEntry) As Boolean
Declare Function Process32Next Lib "kernel32" (ByVal hndl As Long, ByRef pstru As ProcessEntry) As Boolean
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hnd As Long) As Boolean
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function RegisterShellHook Lib "shell32" Alias "#181" (ByVal hwnd As Long, ByVal nAction As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ptWord As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0


Sub SetPriority(pid As Long, priorityClass As Long)

hnd = OpenProcess(PROCESS_SET_INFORMATION, 0, pid)
lRet = SetPriorityClass(hnd, priorityClass)
lRet = CloseHandle(hnd)

End Sub

Public Function GetProcessId(hwnd As Long) As Long

Dim ProcessID As Long
GetWindowThreadProcessId hwnd, ProcessID
GetProcessId = ProcessID
    
End Function

Function GetProcessEXE(pid As Long) As String

Dim bRet            As Boolean
Dim lSnapShot       As Long
Dim tmpPE           As ProcessEntry
Dim tmpProcName     As String
Dim tmpPriority     As String

    lSnapShot = CreateToolhelp32Snapshot(&H2, 0)
    tmpPE.dwSize = Len(tmpPE)
    bRet = Process32First(lSnapShot, tmpPE)
    
    Do Until bRet = False
        If tmpPE.peProcessID = pid Then
            tmpProcName = LCase(Mid(tmpPE.szExeFile, _
                InStrRev(tmpPE.szExeFile, "\", Len(tmpPE.szExeFile)) + 1, _
                Len(tmpPE.szExeFile) - InStrRev(tmpPE.szExeFile, "\", 1)))
            GetProcessEXE = Left(tmpProcName, InStr(1, tmpProcName, Chr(0)) - 1)
        End If
    bRet = Process32Next(lSnapShot, tmpPE)
    Loop
    
    tmpProcName = LCase(Mid(tmpPE.szExeFile, _
                            InStrRev(tmpPE.szExeFile, "\", Len(tmpPE.szExeFile)) + 1, _
                            Len(tmpPE.szExeFile) - InStrRev(tmpPE.szExeFile, "\", 1)))
    GetProcessEXE = Left(tmpProcName, InStr(1, tmpProcName, Chr(0)) - 1)
    exePriority = tmpPE.pePriority
    
    bRet = CloseHandle(lSnapShot)

End Function


