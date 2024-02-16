Attribute VB_Name = "FilesMod"
'*==================================================*
'*This module contains common functions dealing with*
'* files and directory/file name strings            *
'****************************************************

Public Declare Function GetWindowsDirectory Lib "KERNEL32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "KERNEL32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Option Explicit

'*******************************************************
'* Procedure Name: IsEnglishWin                        *
'*-----------------------------------------------------*
'* Created: 01-01-02   By: Only development inc        *
'* Modified:           By:                             *
'*=====================================================*
'*This function will return the language of windows dir*
'*menu/desktop/startmenu                               *
'*******************************************************

Function IsEnglishWin() As Boolean

Dim mWinDir As String
Dim mStartMP As String
Dim mProgramsPA As String

mWinDir = sGetWinDir
mStartMP = mWinDir & "\Start Menu"
mProgramsPA = mStartMP & "\Programs"
    
On Error GoTo er
Open mProgramsPA & "\dr.d" For Output As #55
Write #55, "Ok"
Close #55
IsEnglishWin = True
Kill mProgramsPA & "\dr.d"
Exit Function

er:
IsEnglishWin = False

End Function

'*******************************************************
'* Procedure Name: sGetWinDir                          *
'*-----------------------------------------------------*
'* Created: 8/29/94   By: David McCarter               *
'* Modified: 01-01-02 By: Only Development inc         *
'*=====================================================*
'*This function will return the Windows directory as a *
'*string.                                              *
'*******************************************************
Function sGetWinDir() As String
Dim sPath As String

    sPath = String(255, Chr$(0))

    sGetWinDir = Left$(sPath, GetWindowsDirectory(sPath, Len(sPath)))

End Function

'*******************************************************
'* Procedure Name: sGetWinSysDir                       *
'*-----------------------------------------------------*
'* Created: 10/30/94  By: David McCarter               *
'* Modified: 01-01-02 By: Only Development inc         *
'*=====================================================*
'* Returns the system directory that Windows is        *
'* currently using.                                    *
'*                                                     *
'*                                                     *
'*                                                     *
'*******************************************************
Function sGetWinSysDir() As String
Dim sTempBuff As String
Dim R As Integer

    sTempBuff = Space(144)

    R = GetSystemDirectory(sTempBuff, 144)
    sGetWinSysDir = Trim$(sTempBuff)

End Function
