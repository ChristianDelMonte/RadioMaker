VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////
'*
'*  // DLL module class register / unregister for Vb.6+ //
'*  This module is for Radio Maker only
'*  Code by: Herman Liu
'*
'*  Copyright (c) 1987-2002 Only development Inc.
'*
'////////////////////////////////////////////////////////
Option Explicit

Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpdata As Any) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
'Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long,ByVal dwExitCode As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public ConName As String
Public mCompanyName As String
Public mProductVersion As String
Public RegFlag As Boolean
Public UnregFlag As Boolean
Public mresult
Public gcdg As Object
Public Dropped As Boolean
Public ReX As Integer
Public DoClear As Boolean
Public NotQueried As Boolean

Public Function PlugInExists(WFPlugInName As String) As Boolean

Dim Result As Boolean

Result = FileExist(App.Path & AppPlugInDir & WFPlugInName)
If Result = True Then
    PlugInExists = True
Else
    PlugInExists = False
End If

End Function

Public Sub DispProdVersion(inFile As String)
    If Not GetFileInfo(inFile) Then
        MsgBox = "(No Product Version available for this file)"
    Else
        MsgBox = "Company Name:  " & mCompanyName & vbCrLf & "Product Version:  " & mProductVersion
    End If
End Sub

'//////////////////////////////////////////////////////
'* Function to register / unregister the DLL
'/////////////////////////////////////////////////////
Public Function RegUNRegLib(ByVal inFileSpec As String, OptionalID As Long) As Boolean

On Error Resume Next

Dim lLib As Long                 ' Store handle of the control library
Dim lpDLLEntryPoint As Long      ' Store the address of function called
Dim lpThreadID As Long           ' Pointer that receives the thread identifier
Dim lpExitCode As Long           ' Exit code of GetExitCodeThread
Dim mThread
    
' Load the control DLL, i. e. map the specified DLL file into the
' address space of the calling process
lLib = LoadLibrary(inFileSpec)
If lLib = 0 Then
    ' e.g. file not exists or not a valid DLL file
    RegUNRegLib = False
    Exit Function
End If
    
' Find and store the DLL entry point, i.e. obtain the address of the
' “DllRegisterServer” or "DllUnregisterServer" function (to register
' or deregister the server’s components in the registry).
Select Case OptionalID
    Case 1  'REGISTER DLL
        lpDLLEntryPoint = GetProcAddress(lLib, "DllRegisterServer")
    Case 2  'UNREGISTER DLL
        lpDLLEntryPoint = GetProcAddress(lLib, "DllUnregisterServer")
End Select

If lpDLLEntryPoint = vbNull Then
    GoTo earlyExit1
End If
    
Screen.MousePointer = vbHourglass
    
' Create a thread to execute within the virtual address space of the calling process
mThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
If mThread = 0 Then
    GoTo earlyExit1
End If
    
' Use WaitForSingleObject to check the return state (i) when the specified object
' is in the signaled state or (ii) when the time-out interval elapses.  This
' function can be used to test Process and Thread.
mresult = WaitForSingleObject(mThread, 10000)
If mresult <> 0 Then
    GoTo earlyExit2
End If
    
' We don't call the dangerous TerminateThread(); after the last handle
' to an object is closed, the object is removed from the system.
CloseHandle mThread
FreeLibrary lLib
    
Screen.MousePointer = vbDefault

RegUNRegLib = True
Exit Function
    
    
earlyExit1: '-----------------------------------------------------------
Screen.MousePointer = vbDefault
' Decrements the reference count of loaded DLL module before leaving
FreeLibrary lLib
RegUNRegLib = False
Exit Function
    
earlyExit2: '-----------------------------------------------------------
Screen.MousePointer = vbDefault
FreeLibrary lLib
' Terminate the thread to free up resources that are used by the thread
' NB Calling ExitThread for an application's primary thread will cause
' the application to terminate
lpExitCode = GetExitCodeThread(mThread, lpExitCode)
ExitThread lpExitCode
RegUNRegLib = False
End Function

Public Function GetFileInfo(inFileSpec As String) As Boolean

On Error Resume Next
    
Dim lInfoSize As Long
Dim lpHandle As Long
Dim strFileInfoString As String
Dim i As Integer
    
GetFileInfo = False                                ' Assume
    
' GetFileVersionInfoSize determines if system can obtain version info
' about the specified file.  If yes, it returns its size in bytes and
' a handle to the data.
lpHandle = 0
lInfoSize = GetFileVersionInfoSize(inFileSpec, lpHandle)
If lInfoSize = 0 Then
    Exit Function
End If

' We pass the file name, size(ignored), size of buffer and the buffer of
' version info to GetFileVersionInfo, which will fill the buffer with
' version info about the file. (Modified here).
strFileInfoString = String(lInfoSize, 0)
mresult = GetFileVersionInfo(ByVal inFileSpec, 0&, ByVal lInfoSize, ByVal strFileInfoString)
If mresult = 0 Then
    Exit Function
End If

     ' We now have a block of version data, in an unreadable format though. If you
     ' wish, you may check the existence of "StringFileInfo" with InStr function.
     ' Normally we must call VerQueryValue to read selected pieces of data of the
     ' above, with arguments such as "\VarFileInfo\Translation" or "\StringFileInfo
     ' \lang-codepage\string-name" where lang-codepage is a code which has yet to be
     ' obtained from first 2 words(high-low) returned by "\VarFileInfo\Translation"
     ' from the strFileInfoString (and padded to fixed 8-digit), and string-name is
     ' one of predefined string names such as "CompanyName" & "FileDescription", etc.
     ' However, the following simple alternative is OK for our purpose.

     mCompanyName = ""
     mProductVersion = ""
     i = InStr(strFileInfoString, "CompanyName")
     If i > 0 Then
         i = i + 12
         mCompanyName = Mid$(strFileInfoString, i, 21)
     End If
     i = InStr(strFileInfoString, "FileDescription")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "FileVersion")
     If i > 0 Then
         i = i + 12
     End If
     i = InStr(strFileInfoString, "InternalName")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "LegalCopyright")
     If i > 0 Then
         i = i + 16
     End If
     i = InStr(strFileInfoString, "OriginalFilename")
     If i > 0 Then
         i = i + 20
     End If
     i = InStr(strFileInfoString, "ProductName")
     If i > 0 Then
         i = i + 12
     End If
     i = InStr(strFileInfoString, "ProductVersion")
     If i > 0 Then
         i = i + 16
         mProductVersion = Mid$(strFileInfoString, i)
     End If

     If Trim(mProductVersion) <> "" Then
         GetFileInfo = True
     End If
     
End Function


