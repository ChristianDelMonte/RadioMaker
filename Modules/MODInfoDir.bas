Attribute VB_Name = "MODInfoDir"
'/////////////////////////////////////////////////////////
'
'       (c) Francisco Bonet, Julio de 1997
'
'Programa de prueba y evaluación de las funcions Infodir e
'GETInfoSubDir para obtener el nº de subdirectorios totales,
'nº de archivos, ...
'
'/////////////////////////////////////////////////////////

Private Const MAX_PATH = 64
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

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

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFind As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Type INFODir
    PDirNum As Long
    PFilesNum As Long
    PTotalSize As String * 255
End Type

Public Type INFOSubDir
    SSubDirNum As Long
    SFilesNum As Long
    STotalSize As String * 255
End Type

Private TotSize As Long
Private NumSubdirs As Long
Private NumArxius As Long

Public Function GETInfoDir(WPath As String, WExt As String) As INFODir
    
    Dim atribarx As Long, TotSize As Long
    Dim valor1 As Long, valor2 As Long
    Dim InfoTd As WIN32_FIND_DATA
    Dim NomArxiu As String
    
    On Error Resume Next
    If Right(WPath, 1) <> "\" Then WPath = WPath & "\"
    
    TotSize = 0
    NumSubdirs = 0
    NumArxius = 0
    valor1 = 0
    valor2 = True
    valor1 = FindFirstFile(WPath & WExt, InfoTd)

    Do
        NomArxiu = InfoTd.cFileName
        atribarx = InfoTd.dwFileAttributes
        If Left(NomArxiu, 1) <> "." Then
            If atribarx And FILE_ATTRIBUTE_DIRECTORY Then
                NumSubdirs = NumSubdirs + 1
            Else
                NumArxius = NumArxius + 1
                TotSize = TotSize + InfoTd.nFileSizeLow
            End If
        End If
        valor2 = FindNextFile(valor1, InfoTd)
    Loop Until valor2 = 0
    
    FindClose (valor1)
    'Label2 = "Directorios principales: " & NumSubdirs
    GETInfoDir.PDirNum = NumSubdirs
    DoEvents
    
    'Label3 = "Archivos: " & NumArxius
    GETInfoDir.PFilesNum = NumArxius
    DoEvents
    
    If Int(TotSize) > 1023 Then
        'Label4 = "Bytes: " & Format(Int(TotSize / 1024), "###,###") & " KB"
        GETInfoDir.PTotalSize = "Bytes: " & Format(Int(TotSize / 1024), "###,###") & " KB"
    Else
        'Label4 = "Bytes: " & TotSize & " bytes"
        GETInfoDir.PTotalSize = "Bytes: " & TotSize & " bytes"
    End If
    
    'Label5.Caption = "Un momento por favor ..."
    DoEvents
    
    TotSize = 0
    NumSubdirs = 0
    NumArxius = 0
    
End Function

Public Function GETInfoSubDir(WPath As String, WExt As String) As INFOSubDir
    
    Dim valor1 As Long, valor2 As Long, atribarx As Long
    Dim inull As Integer
    Dim NomArxiu As String, NomSdir As String, NouPath As String, NouCalcul As String
    Dim InfoTd As WIN32_FIND_DATA
    
    On Error Resume Next
    If Right$(WPath, 1) <> "\" Then WPath = WPath & "\"
    valor1 = 0
    valor2 = 1
    valor1 = FindFirstFile(WPath & WExt, InfoTd)
    
    Do
        NomArxiu = RTrim$(InfoTd.cFileName)
        atribarx = InfoTd.dwFileAttributes
        If Left(NomArxiu, 1) <> "." Then
            If atribarx And FILE_ATTRIBUTE_DIRECTORY Then
                NomSdir = NomSdir & WPath & NomArxiu
                NumSubdirs = NumSubdirs + 1
            Else
                NumArxius = NumArxius + 1
                TotSize = TotSize + InfoTd.nFileSizeLow
            End If
        End If
        InfoTd.cFileName = ""
        valor2 = FindNextFile(valor1, InfoTd)
        DoEvents
    Loop Until valor2 = 0
    
    FindClose (valor1)
    
    'RECURSIVIDAD
    Do Until NomSdir = ""
       inull = InStr(NomSdir, vbNullChar)
       If inull Then
           NouPath = Left$(NomSdir, inull - 1)
       End If
       NomSdir = Right$(NomSdir, Len(NomSdir) - inull%)
       'NouCalcul = GETInfoSubDir(NouPath, WExt)
       DoEvents
    Loop
    
    'Label6 = "Archivos: " & NumArxius
    GETInfoSubDir.SFilesNum = NumArxius
    DoEvents
    
    'Label5 = "Subdirectorios: " & NumSubdirs
    GETInfoSubDir.SSubDirNum = NumSubdirs
    DoEvents
    
    If Int(TotSize) >= 1024 Then
        'Label7 = "Bytes: " & Format(Int(TotSize / 1024), "###,###") & " KB"
        GETInfoSubDir.STotalSize = "Bytes: " & Format(Int(TotSize / 1024), "###,###") & " KB"
    Else
        'Label7 = "Bytes: " & TotSize & " bytes"
        GETInfoSubDir.STotalSize = "Bytes: " & TotSize & " bytes"
    End If
    
    DoEvents
    
End Function
