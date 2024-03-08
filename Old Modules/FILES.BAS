'****************************************************
'* FILES.BAS Version 2.0 Date: 02/01/95             *
'* VB Tips & Tricks                                 *
'* 8430-D Summerdale Road San Diego CA 92126-5415   *
'* Compuserve: 74227,1557                           *
'* America On-Line: DPMCS                           *
'* InterNet: DPMCS@AOL.COM                          *
'*==================================================*
'*This module contains common functions dealing with*
'* files and directory/file name strings            *
'****************************************************
Option Explicit

'*******************************************************
'* Procedure Name: FileExists                          *
'*-----------------------------------------------------*
'* Created: 8/29/94   By: David McCarter               *
'* Modified:          By:                              *
'*=====================================================*
'*This function will check to make sure that a file    *
'*exists.It will return True if the file was found and *
'*False if it was not found.                           *
'*Example: If Not FileExists("autoexec.bat") Then...   *
'*******************************************************
Function iFileExists (ByVal sFileName As String) As Integer
Dim I As Integer
On Error Resume Next

    I = Len(Dir$(sFileName))
    
    If Err Or I = 0 Then
        iFileExists = False
        Else
            iFileExists = True
    End If

End Function

'*******************************************************
'* Procedure Name: ValDir                              *
'*-----------------------------------------------------*
'* Created: 8/29/94   By: David McCarter               *
'* Modified:          By:                              *
'*=====================================================*
'*This function is used by MakeDir to validate if a    *
'*directory already exists.                            *
'*******************************************************
Function iValDir (ByVal sDirIn As String) As Integer
Dim iCheck As String, iErrResult As Integer

On Local Error GoTo ValDirError

    iCheck = Dir$(sDirIn)

    If iErrResult = 76 Then
        iValDir = False
        Else
            iValDir = True
    End If

    Exit Function

ValDirError:

    Select Case Err
        Case Is = 76
            iErrResult = Err
            Resume Next

        Case Else

    End Select

End Function

'*******************************************************
'* Procedure Name: MakeDir                             *
'*-----------------------------------------------------*
'* Created: 8/29/94   By: David McCarter               *
'* Modified:          By:                              *
'*=====================================================*
'*This function will create a directory even if the    *
'*underlying directories do not exist.                 *
'*Usage: MakeDir "c:\temp\demo"                        *
'*This procedue also uses the ValDir to find if the    *
'*directory already exists.                            *
'*******************************************************
Sub MakeDir (sDirName As String)
Dim iMouseState As Integer
Dim iNewLen As Integer
Dim iDirLen As Integer

    'Get Mouse State
    iMouseState = Screen.MousePointer

    'Change Mouse To Hour Glass
    Screen.MousePointer = 11

    'Set Start Length To Search For [\]
    iNewLen = 4

    'Add [\] To Directory Name If Not There
    If Right$(sDirName, 1) <> "\" Then sDirName = sDirName + "\"

    'Create Nested Directory
    While Not iValDir(sDirName)
        iDirLen = InStr(iNewLen, sDirName, "\")
        
        If Not iValDir(Left$(sDirName, iDirLen)) Then
            MkDir Left$(sDirName, iDirLen - 1)
        End If
        
        iNewLen = iDirLen + 1
    Wend

    'Leave The Mouse The Way You Found It
    Screen.MousePointer = iMouseState

End Sub

