Attribute VB_Name = "DirTree"
'********************* RM100 *********************
'       RADIO MAKER MINI EXPLORER MODULE
'COPYRIGHT (C) 1987-2002 ONLY development inc.
'Christian A. Del Monte
'*************************************************

Option Explicit

'  API's para informacion de unidades
Public Const MAX_PATH = 255
Public Const WM_SETREDRAW = &HB
Private Const ERROR_NO_MORE_FILES = 18&
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_ICON = &H100
Private Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Private Const SHGFI_LARGEICON = &H0                      '  get large icon
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const ILD_TRANSPARENT = &H1

Private Type SHFILEINFO 'Estructura usada por SHGetFileInfo
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Private Declare Function OSGetLongPathName Lib "STKIT432.DLL" Alias "GetLongPathName" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal Flags&) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public Const DRIVE_PARTITION = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpszSubKey$, dwOptions&, ByVal samDesired&, lpHKey&)
Declare Function RegQueryInfoKey& Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey&, ByVal lpClass$, lpcbClass&, ByVal lpReserved&, lpcSubKeys&, lpcbMaxSubKeyLen&, lpcbMaxClassLen&, lpcValues&, lpcbMaxValueNameLen&, lpcbMaxValueLen&, lpcbSecurityDescriptor&, lpftLastWriteTime As FILETIME)
Declare Function RegEnumValue& Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey&, ByVal dwIndex&, ByVal lpName$, lpcbName&, ByVal lpReserved&, lpdwType&, lpValue As Any, lpcbValue&)
Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)

Private Const ERROR_SUCCESS = 0&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY

Public Type SHITEMID    'Browse Dialog
   cb             As Long
   abID           As Byte
End Type

Public Type ITEMIDLIST  'Browse Dialog
   mkid           As SHITEMID
End Type

Public Type BROWSEINFO  'Browse Dialog
   hOwner         As Long
   pidlRoot       As Long
   pszDisplayName As String
   lpszTitle      As String
   ulFlags        As Long
   lpfn           As Long
   lParam         As Long
   iImage         As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1 'Browse Dialog
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Cancelled As Boolean
'Opens Browse dialog
Public Function BrowseForFolder(Optional Title As String) As String
   
   Dim bi As BROWSEINFO
   Dim pidl As Long
   Dim nRet As Long
   Dim szPath As String
   
   szPath = Space$(512)
   
   bi.hOwner = 0&
   bi.pidlRoot = 0&
   
   bi.lpszTitle = IIf(Title = "", "Directory", Title)
   bi.ulFlags = BIF_RETURNONLYFSDIRS
   
   'Display the dialog and get the selected path
   pidl& = SHBrowseForFolder(bi)
   SHGetPathFromIDList ByVal pidl&, ByVal szPath
   
   'Return value
   BrowseForFolder = Trim$(szPath)
   
End Function

Public Function ExtractIcon(picImage As PictureBox)

Dim himl As Long
Dim lpzxExeName As String
Dim lRet As Long
Dim sWinPath As String
Dim shinfo As SHFILEINFO

sWinPath = GetWinPath()

lpzxExeName = sWinPath & "\explorer.exe"
himl = SHGetFileInfo(lpzxExeName, 0&, shinfo, Len(shinfo), SHGFI_SYSICONINDEX Or SHGFI_SMALLICON)

If himl <> 0 Then
    picImage.AutoRedraw = True
    lRet = ImageList_Draw(himl, shinfo.iIcon, picImage.hDC, 0, 0, ILD_TRANSPARENT)
    picImage.Refresh
    ExtractIcon = True
End If

End Function

Public Function GetDriveName(ByVal sDrive As String) As String

Dim sVolBuf As String, sSysName As String
Dim lSerialNum As Long, lSysFlags As Long, lComponentLength As Long
Dim lRet As Long

sVolBuf = String$(256, 0)
sSysName = String$(256, 0)
lRet = GetVolumeInformation(sDrive, sVolBuf, MAX_PATH, lSerialNum, lComponentLength, lSysFlags, sSysName, MAX_PATH)

If lRet > 0 Then
    sVolBuf = StripTerminator(sVolBuf)
    GetDriveName = StrConv(sVolBuf, vbProperCase)
End If

End Function

Public Function GetComputerName() As String

Dim sValue As String

sValue = "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
sValue = RegEnumValues(HKEY_CLASSES_ROOT, sValue)

If Len(sValue) > 0 Then
    GetComputerName = sValue
Else
    GetComputerName = "My Computer"
End If

End Function

Function GetWinPath() As String

Dim buffer As String, lRet As Long

buffer = Space(MAX_PATH)
lRet = GetWindowsDirectory(buffer, Len(buffer))
GetWinPath = Left(buffer, lRet)

End Function

Public Function RegEnumValues(lMainKey As Long, sSubKey As String) As String

Dim lRtn As Long
Dim hKey As Long
Dim lLenValueName As Long
Dim lLenValue As Long
Dim sRegEntry As String
Dim sRegValue As String
Dim lDataType As Long
Dim sClassName As String
Dim lClassLen As Long
Dim lSubKeys As Long
Dim lMaxSubKey As Long
Dim lMaxClass As Long
Dim lValues As Long
Dim lMaxValueName As Long
Dim lMaxValueData As Long
Dim lSecurityDesc As Long
Dim strucLastWriteTime As FILETIME

lRtn = RegOpenKeyEx(lMainKey, sSubKey, 0&, KEY_READ, hKey)

If lRtn <> ERROR_SUCCESS Then Exit Function

sClassName = Space$(255)
lClassLen = CLng(Len(sClassName))
lRtn = RegQueryInfoKey(hKey, sClassName, lClassLen, 0&, lSubKeys, lMaxSubKey, lMaxClass, lValues, lMaxValueName, lMaxValueData, lSecurityDesc, strucLastWriteTime)


sRegEntry = Space$(lMaxValueName + 1)
lLenValueName = CLng(Len(sRegEntry))
sRegValue = Space$(lMaxValueData + 1)
lLenValue = CLng(Len(sRegValue))

lRtn = RegEnumValue(hKey, 0&, sRegEntry, lLenValueName, 0&, lDataType, ByVal sRegValue, lLenValue)
    
If lRtn = ERROR_SUCCESS Then
    RegEnumValues = Mid$(sRegValue, 1, lLenValue)
End If

lRtn = RegCloseKey(hKey)

End Function

Function HasSubDirs(ByVal sStartDir As String) As String

'Esta funcion chequea los directorios en busca de sub-directorios

Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long, lRet As Long
Dim sTemp As String

On Error Resume Next

If sStartDir = "" Then Exit Function

If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"

sStartDir = sStartDir & "*.*"
lFileHdl = FindFirstFile(sStartDir, lpFindFileData)

If lFileHdl <> 0 Then
    Do Until lRet = ERROR_NO_MORE_FILES
        If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 16 Then
            sTemp = StripTerminator(lpFindFileData.cFileName)
            If sTemp <> "." And sTemp <> ".." Then
                HasSubDirs = StrConv(sTemp, vbProperCase)
                Exit Do
            End If
        End If
        lRet = FindNextFile(lFileHdl, lpFindFileData)
        If lRet = 0 Then Exit Function
    Loop
End If

lRet = FindClose(lFileHdl)

End Function

Sub ListDirs(tvwTree As ComctlLib.TreeView, ByVal sStartDir As String)

'Esta funcion busca todos los subdirectorios dentro del directorio dado para
'a su vez buscar mas sub directorios dentro de los mismos.

Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long, nodX As ComctlLib.Node
Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer

On Error Resume Next
If sStartDir = "" Then Exit Sub

If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"

sStartDir = sStartDir & "*.*"
lFileHdl = FindFirstFile(sStartDir, lpFindFileData)

If lFileHdl <> 0 Then
    Do Until lRet = ERROR_NO_MORE_FILES
        If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = vbDirectory Then
            sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
            If sTemp <> "." And sTemp <> ".." Then
                Set nodX = tvwTree.Nodes.Add(Left$(sStartDir, Len(sStartDir) - 4), tvwChild, Left$(sStartDir, Len(sStartDir) - 4) & "\" & sTemp, sTemp, "closed")
                nodX.SelectedImage = "open"
                sTemp2 = HasSubDirs(Left$(sStartDir, Len(sStartDir) - 4) & "\" & sTemp)
                If Len(sTemp2) > 0 Then
                    Set nodX = tvwTree.Nodes.Add(Left$(sStartDir, Len(sStartDir) - 4) & "\" & sTemp, tvwChild, Left$(sStartDir, Len(sStartDir) - 4) & "\" & sTemp & "\" & sTemp2, sTemp2, "closed")
                    nodX.SelectedImage = "open"
                End If
            End If
        End If
        lRet = FindNextFile(lFileHdl, lpFindFileData)
        If lRet = 0 Then Exit Do
    Loop
End If

lRet = FindClose(lFileHdl)
End Sub

Public Function ListDrives(frm As Form, tvwTree As ComctlLib.TreeView, imgList As ComctlLib.ImageList, picImage As PictureBox) As String

Dim sTemp As String, sTemp2 As String
Dim lRet As Long
Dim iNullSpot As Integer
Dim nodX As ComctlLib.Node, lDrive As Long
Dim sSubDir As String
Dim sCompName As String
Dim sSelectDrive As String
Dim sDrvPic As String
Dim imgX As ComctlLib.ListImage
Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long

Screen.MousePointer = 13

If ExtractIcon(picImage) Then
    Set imgX = imgList.ListImages.Add(, "desktop", picImage.Image)
    sDrvPic = "desktop"
Else
    sDrvPic = "open"
End If

tvwTree.ImageList = imgList
sCompName = GetComputerName()
Set nodX = tvwTree.Nodes.Add(, , "desk", "Desktop", "desk")
Set nodX = tvwTree.Nodes.Add("desk", tvwChild, "root", sCompName, sDrvPic)

sTemp = String$(2048, 0)
Call GetLogicalDriveStrings(2047, sTemp)
Do
    iNullSpot = InStr(sTemp, Chr$(0))
    
    If iNullSpot > 1 Then
        sTemp2 = UCase$(Left$(sTemp, iNullSpot - 2))
        lDrive = GetDriveType(sTemp2)
        sSubDir = ""
        
        Select Case lDrive
            Case DRIVE_FIXED, DRIVE_PARTITION
                sDrvPic = "hard"
                sSubDir = HasSubDirs(sTemp2)
                sTemp2 = GetDriveName((sTemp2 & "\")) & " (" & sTemp2 & ")"
                If sSelectDrive = "" Then sSelectDrive = sTemp2
            Case DRIVE_CDROM
                sDrvPic = "cdrom"
                sSubDir = HasSubDirs(sTemp2)
                sTemp2 = GetDriveName((sTemp2 & "\")) & " (" & sTemp2 & ")"
            Case DRIVE_REMOTE
                sDrvPic = "net"
                sSubDir = HasSubDirs(sTemp2)
                sTemp2 = GetDriveName((sTemp2 & "\")) & " (" & sTemp2 & ")"
            Case DRIVE_REMOVABLE
                sDrvPic = "floppy"
                sTemp2 = "Floppy (" & sTemp2 & ")"
            Case Else
                sTemp2 = "(" & sTemp2 & ")"
        End Select
        
        Set nodX = tvwTree.Nodes.Add("root", tvwChild, UCase$(Left$(sTemp, iNullSpot - 2)), sTemp2, sDrvPic)
        
        If Len(sSubDir) > 0 Then
            Set nodX = tvwTree.Nodes.Add(UCase$(Left$(sTemp, iNullSpot - 2)), tvwChild, UCase$(Left$(sTemp, iNullSpot - 2)) & "\" & sSubDir, sSubDir, "closed")
            nodX.SelectedImage = "open"
        End If
        
        sTemp = Mid$(sTemp, iNullSpot + 1)
    End If

Loop Until iNullSpot <= 1
 
sSubDir = GetWinPath()
sSubDir = StrConv(Left$(sSubDir, 3), vbProperCase) & StrConv(Right$(sSubDir, Len(sSubDir) - 3), vbProperCase) & "\Desktop\*.*"
lFileHdl = FindFirstFile(sSubDir, lpFindFileData)

If lFileHdl <> 0 Then
    Do Until lRet = ERROR_NO_MORE_FILES
        If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 16 Then
            sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
            If sTemp <> "." And sTemp <> ".." Then
                Set nodX = tvwTree.Nodes.Add("desk", tvwChild, Left$(sSubDir, Len(sSubDir) - 3) & sTemp, sTemp, "closed")
            End If
        End If
        lRet = FindNextFile(lFileHdl, lpFindFileData)
        If lRet = 0 Then Exit Do
    Loop
End If

tvwTree.Nodes(4).EnsureVisible

Screen.MousePointer = 0

ListDrives = sSelectDrive

End Function

Function StripTerminator(ByVal strString As String) As String
    
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
    
End Function
