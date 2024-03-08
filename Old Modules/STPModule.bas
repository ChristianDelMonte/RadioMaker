Attribute VB_Name = "STPModule"
'////////////////////////////////////////////
'*
'*  Setup Managger module.
'*  Code by: Only development software inc.
'*           Copyright (c) 1987-2002.
'*  Christian A. Del Monte
'////////////////////////////////////////////

Option Explicit

Public Const APPStpIniFile = "\Setup.ini"
Public Const APPStpFile = "\Setup.dat"

'---
Public Const DefPass = "STPMaker1"
'---

'--------------------------------------------
'Constantes para manejo de directorios
'--------------------------------------------
Public Const STP_None = 0
Public Const STP_AppDir = 1
Public Const STP_WinDir = 2
Public Const STP_WinSysDir = 3
Public Const STP_AppSubDir = 4
'--------------------------------------------

Public Type StpFile
    Id As Integer
    FileName As String * 255         'nombre de archivo
    CABFileName As String * 255      'cab donde se encuentra
    Destination As Integer           'directorio de destino STP_XXX
    DestNum As Integer               'the number of AppPSubDir if is
End Type                             'an STP_AppSubDir (1 to 9) 0=none

Public Type AppSubDirs
    Dr(0 To 8) As String * 255       'directorio 1 de destino
End Type

Public Type StpIniFile
    Id As Integer
    AppTitle As String * 100            'titulo de la aplicacion
    APPEXEName As String * 100          'nombre de archivo EXE
    APPEXEDesc As String * 100          'descripcion para EXEName (in start menu)
    APPReadmeName As String * 100       'nombre de archivo readme
    APPReadmeDesc As String * 100       'descripcion para readme (in start menu)
    AppVersion As String * 50           'version
    AppCompany As String * 100          'nombre de la compania
    AppComment As String * 255          'comentarios opcionales
    AppDefDir As String * 255           'directorio maestro (por defecto)
    AppDefSubDir As AppSubDirs          'subdirectorios dentro del maestro
    FrmTitle As String * 100            'titulo dentro del formulario
End Type

Public IniData As StpIniFile
Public AppData As StpFile

Public Function StripFileFromExt(WFileName As String) As String

'******************************************************
'funcion para extraer el nombre de un archivo solamente
'sin la extensión y que puede incluir un path.
'******************************************************

Dim FinLn As Long, GeSTR As String

WFileName = Trim(WFileName)

FinLn = InStr(1, WFileName, ".", vbTextCompare)
GeSTR = Left$(WFileName, FinLn - 1)
StripFileFromExt = GeSTR

End Function

Public Function DenByte(WByte() As Byte, WLen As Long) As String

'funcion para reconvertir una cadena de bytes en codigos
'ASCCI y posteriormente en texto
'///////////////////////////////////////////////////////

Dim i As Integer, char As String, Num As Long, Str As String

For i = 1 To WLen
    Num = WByte(i)
    char = Chr$(Num)
    Str = Str & char
Next i

DenByte = Trim(Str)

End Function

Public Function EnByte(WStr As String, RtnByte() As Byte, RtnSize As Long)

'Funcion para convertir una cadena de texto en codigos
'ASCCI y posteriormente en bytes
'/////////////////////////////////////////////////////

Dim ln As Long, i As Integer, char As String, rst As Integer

ln = Len(Trim(WStr))
ReDim RtnByte(1 To ln) As Byte

For i = 1 To ln
    char = Mid$(WStr, i, 1)
    rst = Asc(char)
    RtnByte(i) = rst
    RtnSize = ln
Next i

End Function

Public Function StripDirFromFile(WFileName As String) As String

'******************************************************
'funcion para extraer el path "solo" de un path que
'se encuentra completo (incluyendo nombre de archivo)
'******************************************************

Dim FileLn As Long
Dim FindLn As Long
Dim d As Integer
Dim GetSTR As String

WFileName = Trim(WFileName)
FileLn = Len(WFileName)

For d = FileLn To 1 Step -1
    GetSTR = Mid$(WFileName, d, 1)
    If GetSTR = "\" Then
        FindLn = d - 1
        StripDirFromFile = Left$(WFileName, FindLn)
        Exit For
        Exit Function
    End If
Next d

End Function

Public Function StripExtFromFile(WFileName As String) As String

'******************************************************
'funcion para extraer la extension "sola" de un nombre
'de archivo completo que puede incluir un path.
'******************************************************

Dim FilLn As Long, FinLn As Long

WFileName = Trim(WFileName)
FilLn = Len(WFileName)

FinLn = InStr(1, WFileName, ".", vbTextCompare)
FinLn = FilLn - FinLn
StripExtFromFile = Right$(WFileName, FinLn)

End Function

Public Function ReadAppFile(WPath As String, WRegId As Integer) As StpFile

'funcion para extraer los datos del archivo "setup.dat"
'//////////////////////////////////////////////////////

On Error GoTo FileOpenErr
'lets open the tanda file for read the data
Open WPath & APPStpFile For Random As #12 Len = Len(AppData)
'get the file data
Get #12, WRegId, AppData
Close #12

ReadAppFile = AppData
Exit Function

FileOpenErr:
    Close #12
    RStpError Err.Number
    
End Function

Public Function StripFileFromDir(WFileName As String) As String

'******************************************************
'funcion para extraer el nombre de un archivo "solo"
'que se encuentra dentro de un path completo.
'******************************************************

Dim FLn As Long
Dim FiLn As Long
Dim z As Integer
Dim GSTR As String

WFileName = Trim(WFileName)
FLn = Len(WFileName)

For z = FLn To 1 Step -1
    GSTR = Mid$(WFileName, z, 1)
    If GSTR = "\" Then
        FiLn = (FLn - z)
        StripFileFromDir = Right$(WFileName, FiLn)
        Exit For
        Exit Function
    End If
Next z

End Function

Public Function WriteAppFile(WData As StpFile, WPath As String)

'******************************************************
'funcion para guardar los datos del archivo "setup.dat"
'******************************************************

Dim NumReg As Integer
Dim RPath As String

NumReg = WData.Id

'chequeamos el path
If Trim(WPath) = "" Or Trim(WPath) = " " Then
    RPath = App.Path
Else
    RPath = WPath
End If

'encriptamos los datos antes de guardar
WData.CABFileName = Encriptar(DefPass, WData.CABFileName)
WData.FileName = Encriptar(DefPass, WData.FileName)

'guardamos los datos en el archivo de datos
On Error GoTo Err
Open RPath & APPStpFile For Random As #43 Len = Len(AppData)
Put #43, NumReg, WData
Close #43
Exit Function

Err:
Close #43
RStpError Err.Number    'raise the error

End Function

Public Function GetAppFileLastReg(WPath As String) As Integer

'function para extraer el numero de ultimo registro del archivo "setup.dat"
'//////////////////////////////////////////////////////////////////////////

'cargamos el archivo de datos
On Error GoTo Err
Open WPath & APPStpFile For Random As #10 Len = Len(AppData)

'get the file last record
GetAppFileLastReg = LOF(10) \ Len(AppData)

Close #10
Exit Function

Err:
Close #10
'RStpError Err.Number    'raise the error
GetAppFileLastReg = 0

End Function

Public Function ReadIniFile(WPath As String, WRegId As Integer) As StpIniFile

'function para extraer los datos del archivo "setup.ini"
'///////////////////////////////////////////////////////

On Error GoTo FileOpenErr
'lets open the tanda file for read the data
Open WPath & APPStpIniFile For Random As #14 Len = Len(IniData)
'get the file data
Get #14, WRegId, IniData
Close #14

ReadIniFile = IniData
Exit Function

FileOpenErr:
    Close #14
    RStpError Err.Number

End Function

Public Function WriteIniFile(WData As StpIniFile, WPath As String)

'******************************************************
'funcion para guardar los datos del archivo "setup.ini"
'******************************************************

Dim NumReg As Integer
Dim RPath As String
Dim i As Integer

NumReg = WData.Id

'chequeamos el path
If Trim(WPath) = "" Or Trim(WPath) = " " Then
    RPath = App.Path
Else
    RPath = WPath
End If

'encriptamos los datos antes de guardar
WData.AppComment = Encriptar(DefPass, WData.AppComment)
WData.AppCompany = Encriptar(DefPass, WData.AppCompany)
WData.AppDefDir = Encriptar(DefPass, WData.AppDefDir)
For i = 0 To 8
    WData.AppDefSubDir.Dr(i) = Encriptar(DefPass, WData.AppDefSubDir.Dr(i))
Next i
WData.AppTitle = Encriptar(DefPass, WData.AppTitle)
WData.AppVersion = Encriptar(DefPass, WData.AppVersion)
WData.FrmTitle = Encriptar(DefPass, WData.FrmTitle)
WData.APPEXEName = Encriptar(DefPass, WData.APPEXEName)
WData.APPEXEDesc = Encriptar(DefPass, WData.APPEXEDesc)
WData.APPReadmeName = Encriptar(DefPass, WData.APPReadmeName)
WData.APPReadmeDesc = Encriptar(DefPass, WData.APPReadmeDesc)

'guardamos los datos en el archivo ini
On Error GoTo Err
Open RPath & APPStpIniFile For Random As #44 Len = Len(IniData)
Put #44, NumReg, WData
Close #44
Exit Function

Err:
Close #44
RStpError Err.Number    'raise the error

End Function

Public Function GetIniFileLastReg(WPath As String) As Integer

'funcion para extraer el numero del ultimo registro del archivo "setup.ini"
'//////////////////////////////////////////////////////////////////////////

'cargamos los datos del archivo ini
On Error GoTo Err
Open WPath & APPStpIniFile For Random As #11 Len = Len(IniData)

'get the file last record
GetIniFileLastReg = LOF(10) \ Len(IniData)

Close #11
Exit Function

Err:
Close #11
'RStpError Err.Number    'raise the error
GetIniFileLastReg = 0

End Function

Public Function RStpError(ErrNum As Long)

'******************************************************
'funcion para realzar errores dentro del STPModule
'******************************************************

MsgBox Error$(ErrNum), vbCritical, App.ProductName & " ERROR!"

End Function
