Attribute VB_Name = "FileManagger"
'////////////////////////////////////////////////////
'*
'*  // FILE managger module for Vb.6+ //
'*  ** module for Radiomaker 1.0 only **
'*  Copyright (c) 1987-2002 Only development Inc.
'*  Christian A. Del Monte
'///////////////////////////////////////////////////

Option Explicit

'DEFINICION DE CONSTANTES GLOBALES DEL SISTEMA
'----------------------------------------------------------------------------------
'**************** DIRECTORIOS DE TRABAJO ******************************************
Public Const AppConfigDir = "\Config"           'directorio de la configuracion
Public Const AppPHDir = "\PH"                   'directorio de Programacion Horaria
Public Const AppTandaDir = "\Tandas"            'directorio de Tandas
Public Const AppDataDir = "\Data"               'directorio de datos del programa
Public Const AppReportDir = "\Report"           'directorio de los reportes de aire
Public Const AppEstDir = "\EstData"             'directorio de las estaciones
Public Const AppProgDir = "\Prog"               'directorio de la programacion
Public Const AppAutoGenDir = "\AutoGen"         'directorio del generador de Tandas Autom.
Public Const AppTopTenDataDir = "\TopTenData"   'directorio de los TopTen data
Public Const AppPlugInDir = "\Plugins"          'directorio de los plug-ins
Public Const AppUpdateDir = "\LiveUpdate"       'directorio de actualizaciones
Public Const AppDefaultMusicPath = "\Audio"     'directorio de audio (por defecto)

'----------------------------------------------------------------------------------
'**************** ARCHIVOS DE TRABAJO ********************************************
Public Const AppConfigFile = "\Config.rmd"              'archivo de configuracion
Public Const AppStateFile = "\Config.rms"               'archivo de conf. de estado
Public Const AppInitFile = "\Status.rmi"                'archivo de inicializacion
Public Const AppErrFile = "\Reporte.rme"                'archivo de rep. de errores
Public Const AppUsrPwrFile = "\UsrData.rmp"             'password file del usuario
Public Const AppGeneralFile = "\Rm100.cfg"              'NEW config file v.3.0
Public Const AppReservedFile = "\Cam.cam"               'reserved personal file
Public Const AppUpdateVerFile = "\VerChk.dat"           'archivo de version de actualizacion

'archivos internos
Private Const AppEst1VizFileH = "\Est1ctrH.rmd"
Private Const AppEst1VizFileV = "\Est1ctrV.rmd"
Private Const AppEst2VizFileH = "\Est2ctrH.rmd"
Private Const AppEst2VizFileV = "\Est2ctrV.rmd"

Public Const AppCUEFileExt = ".rmc"       'extensión de archivos CUE
Public Const AppPHFileExt = ".ph1"        'extension de archivos PH
Public Const AppTndFileExt = ".tnd"       'extension de archivos Tanda
Public Const AppEstFileExt = ".est"       'extension de archivos EST1 y 2
Public Const AppPrgFileExt = ".prg"       'extension de archivos de Programacion

'----------------------------------------------------------------------------------
'*************** Constantes de passwords de encriptación ****************************
Private Const CipherPass = "RM100v1a"              'password de archivo Config
Private Const StatePassW = "PutState"              'Password de archivo State
Private Const UsrPass = "RM100Usr"                 'password de archivo UsrPass
Private Const CUEFilePass = "CUEFile"              'password de archivo CUE

'----------------------------------------------------------------------------------
'*************** Constantes para cabecera de archivos *******************************
Private Const StFHeader = "RM100StateFile"      '=14 cabecera de archivo state
Private Const CfgFHeader = "RM100ConfigFile"    '=15 cabecera de archivo config
Private Const CUEFHeader = "RM100CUEFile"       '=12 cabecera de archivo cue

'----------------------------------------------------------------------------------
'*** Type para el manejo de las Tandas *******************************************
Type Temas
    id As Integer                 'identificador o numero de registro
    Name As String * 255          'nombre del tema o nombre del bloque
                                  '                      ((si es un bloque debe comienzar con BLOCK:xx))
    FNType As String * 10         'tipo de archivo (stream or music)
    Direccion As String * 255     'path del tema
    Duracion As String * 8        'duracion del tema '00:00:00
                                  '                      ((o duracion del bloque si es BLOCK:))
    Hora As String * 8            'hora de lanzamiento del tema '00:00:00
                                  '                      ((u hora de lanzamiento del bloque si es BLOCK:))
    NameX As String * 255         'nombre del tema mixado
                                  '                      ((= BLOCK: si es un bloque))
    FNTypeX As String * 10        'tipo de archivo (stream or music)
                                  '                      ((=vacio si es BLOCK:))
    DireccionX As String * 255    'path del tema de mixado
                                  '                      ((=vacio si es BLOCK:))
    DuracionX As String * 8       'duracion del mixado '00:00:00
                                  '                      ((=vacio si es BLOCK:))
    HoraX As String * 5           'hora de lanzamiento del mixado '00:00
                                  '                      ((u hora predeterminada de lanz. si es BLOCK:))
End Type

'----------------------------------------------------------------------------------
'*** Type para el manejo de las estaciones ***************************************
Type EstFileData
    id As Integer                   'identificador
    Control As String * 3           'identificacion del control
    CCaption As String * 100        'caption del control
    FName As String * 255           'name y path del archivo
    FType As String * 10            'tipo de archivo (stream or music)
    FDuracion As String * 8         'duracion del archivo de audio 00:00:00
End Type

'----------------------------------------------------------------------------------
'*** Type para el manejo de Lock y Serial number de Radio Maker 1.0 **************
Type Record
    id As Integer
    Data1 As String * 40            'lock num
    Data2 As String * 40            'serial num
End Type

'----------------------------------------------------------------------------------
'*** Type para el manejo de la Programacion de Tandas ****************************
Type PrgRecord
    id As Integer
    TndFileName As String * 255     'nombre y path de la tanda
    TndFileCaption As String * 255  'caption del control o nombre de la tanda
    TndDuracion As String * 8       'duracion de la Tanda = 00:00:00 to 24:00:00 Hs
End Type

'----------------------------------------------------------------------------------
'*** Type maestro de opciones de configuracion ***********************************
Type ConfigRecord
    id As Integer
    ConfigHeader As String * 15
    'panel de configuracion GENERAL
    Gen_AutoTAG As Integer              'extraer TAG info?              1=si    0=no
    Gen_AutoName As Integer             'Extraer Music Name?            1=si    0=no
    Gen_ActiveReport As Integer         'Generar reporte?               1=si    0=no
    Gen_ReportEst As Integer            'reporte de estacion 01y02      1=si    0=no
    Gen_ReportTnd As Integer            'reporte de tanda 01y02         1=si    0=no
    Gen_ReportAll As Integer            'reportar todo                  1=si    0=no
    Gen_ReportProg As String * 255      'programa editor de reporte
    Gen_EditProg As String * 255        'programa editor de audio
    Gen_GrabProg As String * 255        'programa grabador de audio
    'panel de configuracion AUDIO
    Aud_Type As Integer                 '1=8bits        2=16bits
    Aud_Cual As Integer                 '1=Mono         2=Stereo
    Aud_Mode As Integer                 '1=Normal       2=A3d    3=3d    4=Ogg
    Aud_Mod_Type As Integer             '1=Ramp Normal  2=Ramp Sensitive
    Aud_Mod_Cual As Integer             '0=None         1=Surround
    Aud_Mod_Mode As Integer             '1=as FT2       2=as PT2
    Aud_Disp_Time As Integer            '1=normal       2=rest
    Aud_Disp_Wave As Integer            '1=normal       2=rest
    Aud_Disp_Samp As Integer            '1=normal       2=rest
    Aud_Show_MiniRM As Integer          '1=si       0=no
    Aud_Show_FTT As Integer             '1=si       0=no
    Aud_Show_SCOPE As Integer           '1=si       0=no
    'panel de configuracion de DIRECTORIOS
    Dir_Tem As String * 255             'directorio temas
    Dir_Com As String * 255             'directorio comerciales
    Dir_Inst As String * 255            'directorio institucionales
    Dir_Hor As String * 255             'directorio horario
    'panel de configuracion de SEGURIDAD
    Sec_Type As Integer                 '1=none     2=password      3=denegar
    Sec_Est12_1 As Integer              '1=open/save/new         0=none
    Sec_Est12_2 As Integer              '1=play/stop/pause       0=none
    Sec_Est12_3 As Integer              '1=drag/del/change       0=none
    Sec_Tnd12_1 As Integer              '1=open/save/new         0=none
    Sec_Tnd12_2 As Integer              '1=play/stop/pause       0=none
    Sec_Tnd12_3 As Integer              '1=drag/del/change       0=none
    Sec_Prg_1 As Integer                '1=open/save/new         0=none
    Sec_Prg_2 As Integer                '1=play/stop/pause       0=none
    Sec_Prg_3 As Integer                '1=drag/del/change       0=none
    Sec_Esp_1 As Integer                '1=in config             0=none
    Sec_Esp_2 As Integer                '1=in run rm100          0=none
    Sec_Esp_3 As Integer                '1=in change options     0=none
    Sec_Esp_4 As Integer                '1=in exit rm100         0=none
    Sec_Esp_5 As Integer                '1=in run plug-in        0=none
End Type

'----------------------------------------------------------------------------------
'*** Type para archivos de opciones CUE ******************************************
Type CUERecord
    id As Integer
    CueHeader As String * 12
    DisplayCUEStartTime As String * 8   'tiempo de inicio       (hora,min,seg)
    DisplayCUEEndTime As String * 8     'tiempo de fin          (hora,min,seg)
    DisplayCUEStartMark As String * 4   'inicio de la marca     (seg)
    DisplayCUELengthMark As String * 4  'longitud de la marca   (seg)
    CUEStartByte As String * 30         'tiempo de inicio       (byte)
    CUEEndByte As String * 30           'tiempo de finalizacion (byte)
    'EQ presets
    EQValue(0 To 10) As Integer
End Type

'----------------------------------------------------------------------------------
'*** Type para archivo de estado del programa ************************************
Type StateRecord
    id As Integer
    StateHeader As String * 14
    LastTndFile As String * 255         'Ultima tanda utilizada
    LastPrgFile As String * 255         'Ultima programacion utilizada
    LastEst1File As String * 255        'Ultimo archivo de estacion 1 utilizado
    LastEst2File As String * 255        'Ultimo archivo de estacion 2 utilizado
    LastWinOrder As String * 8          'Ultimo orden de ventanas al cerrar
End Type

'----------------------------------------------------------------------------------
'*** Type para archivos de Programacion horaria. *********************************
Type PHRecord
    id As Integer
    filename As String * 255            'Nombre del archivo de audio a procesar
    FileLounch As String * 5            'Hora de lanzamiento del mismo
End Type

'----------------------------------------------------------------------------------
'*** Dimensiones de trabajo ******************************************************
Public NumReg As Long
Public EstData As EstFileData           'registros de Archivos de Estacion
Public TndData As Temas                 'registros de archivos de Tandas
Public PrgData As PrgRecord             'registros de Programacion de Tandas
Public ConfigData As ConfigRecord       'registros de Configuracion
Public StateData As StateRecord         'registros de estado del programa
Public LockData As Record               'Registros de Registro del Sistema
Public CUEData As CUERecord             'Registros de tiempo para proceso CUE
Public PHData As PHRecord               'Registros de tiempo de Programacion Horaria

Public Function StripFileFromDir(WFileName As String) As String

'******************************************************
'funcion para extraer el nombre de un archivo "solo"
'que se encuentra dentro de un path completo.
'******************************************************

Dim FLn As Long
Dim FiLn As Long
Dim Z As Integer
Dim GSTR As String

On Error GoTo StripErr
WFileName = Trim(WFileName)
FLn = Len(WFileName)

For Z = FLn To 1 Step -1
    GSTR = Mid$(WFileName, Z, 1)
    If GSTR = "\" Then
        FiLn = (FLn - Z)
        StripFileFromDir = Right$(WFileName, FiLn)
        Exit For
        Exit Function
    Else
        StripFileFromDir = WFileName
    End If
Next Z
Exit Function

StripErr:
StripFileFromDir = "NoFiles..."
End Function

Public Function StripExtFromFile(WFileName As String) As String

'******************************************************
'funcion para extraer la extension "sola" de un nombre
'de archivo completo que puede incluir un path.
'******************************************************

Dim FilLn As Long, FinLn As Long

On Error GoTo StripErr
WFileName = Trim(WFileName)
'FilLn = Len(WFileName)

'FinLn = InStr(1, WFileName, ".", vbTextCompare)
'FinLn = FilLn - FinLn
StripExtFromFile = Right$(WFileName, 3) ' FinLn)
Exit Function

StripErr:
StripExtFromFile = "NoFiles..."
End Function

Public Function StripDirFromFile(WFileName As String) As String

'******************************************************
'funcion para extraer el path "solo" de un path que
'se encuentra completo (incluyendo nombre de archivo)
'******************************************************

Dim FileLn As Long
Dim FindLn As Long
Dim D As Integer
Dim GetSTR As String

On Error GoTo StripErr
WFileName = Trim(WFileName)
FileLn = Len(WFileName)

For D = FileLn To 1 Step -1
    GetSTR = Mid$(WFileName, D, 1)
    If GetSTR = "\" Then
        FindLn = D - 1
        StripDirFromFile = Left$(WFileName, FindLn)
        Exit For
        Exit Function
    Else
        StripDirFromFile = WFileName
    End If
Next D
Exit Function

StripErr:
StripDirFromFile = "NoFiles..."
End Function

Public Function StripFileFromExt(WFileName As String) As String

'******************************************************
'funcion para extraer el nombre de un archivo solamente
'sin la extensión y que puede incluir un path.
'devuelve NotOk si hay algun error!!!.
'******************************************************

Dim FinLn As Long, GeSTR As String

WFileName = Trim(WFileName)

On Error GoTo StripErr
FinLn = InStr(1, WFileName, ".", vbTextCompare)
GeSTR = Left$(WFileName, FinLn - 1)
StripFileFromExt = GeSTR
Exit Function

StripErr:
StripFileFromExt = "NoFiles..."
End Function

Function FileExist(WFileName As String) As Boolean

Dim Data As String
Dim FName As String

FName = Trim(WFileName)

On Error GoTo nop
    '/// check if file exists
'Open FName For Input As #44
'Input #44, Data
'Close #44

FileExist = True
Exit Function

nop:    '/// file does not exists
Close #44
FileExist = False
End Function

Function OpenConfigFile() As ConfigRecord

NumReg = 1

'abrimos el archivo de configuracion
On Error GoTo err
Open App.path & AppConfigDir & AppConfigFile For Random As #33 Len = Len(ConfigData)
Get #33, NumReg, ConfigData
Close #33

'chequeamos la cabecera del archivo
If Trim(ConfigData.ConfigHeader) = Trim(CfgFHeader) Then
    OpenConfigFile = ConfigData
Else
    GoSub err
End If
Exit Function

'------------------------------------------
err:
'no hay config file
Close #33

End Function

Function OpenPHFile(FNombre As String) As String

Dim LastRecord As Long
Dim NumRg As Integer
Dim NewCont As Integer
Dim X As Integer

If FileExist(FNombre) = False Then
    OpenPHFile = "NotOk"
    Exit Function
End If

On Error GoTo FileOpenErr
FNombre = Trim(FNombre)
'lets open the ph file for read the data
Open FNombre For Random As #10 Len = Len(PHData)

'lets get the record len
LastRecord = LOF(10) \ Len(PHData)
'lets get out the data
For X = 1 To LastRecord
    NumRg = X
    'get the file data
    Get #10, NumRg, PHData
    'put the data in a list
    NewCont = NumRg - 1
    FrmTime.Text1(NewCont).text = Trim(PHData.filename)
    If FrmTime.Text1(NewCont).text = "" Then
        FrmTime.Text1(NewCont).BackColor = &HFFFFFF   'normal
        FrmTime.Text2(NewCont).BackColor = &HFFFFFF   'normal
    Else
        FrmTime.Text1(NewCont).BackColor = &HC0FFFF 'amarillo
        FrmTime.Text2(NewCont).BackColor = &HC0FFFF 'amarillo
    End If
    FrmTime.Text2(NewCont).text = Trim(PHData.FileLounch)
Next X
'close the ph file
Close #10
OpenPHFile = "Ok"
Exit Function

FileOpenErr:
ErrorMsg err.Number, LoadResString(340)
ErrorReporte LoadResString(340)
OpenPHFile = "NotOk"
Close #10

End Function

Function OpenPrgFile(WPrgFileName As String) As String

Dim LastRecord As Long
Dim NumRg As Integer
Dim X As Integer

If FileExist(WPrgFileName) = False Then
    OpenPrgFile = "NotOk"
    Exit Function
End If

On Error GoTo FileOpenErr
WPrgFileName = Trim(WPrgFileName)
'lets open the ph file for read the data
Open WPrgFileName For Random As #14 Len = Len(PrgData)
'lets get the record len
LastRecord = LOF(14) \ Len(PrgData)
LastRecord = LastRecord - 1
'lets get out the data
For X = 0 To LastRecord
    NumRg = X + 1
    'get the file data
    Get #14, NumRg, PrgData
    'put the data in a programacion control
    Est12Data.PF(X).Caption = Trim(PrgData.TndFileName)
    Est12Data.PC(X).Caption = Trim(PrgData.TndFileCaption)
    Est12Data.PD(X).Caption = Trim(PrgData.TndDuracion)
    Prg01.Prg1(X).Caption = Trim(PrgData.TndFileCaption)
    Prg01.Prg1(X).ToolTipText = "Duración: " & Trim(PrgData.TndDuracion)
Next X

'close the ph file
Close #14
OpenPrgFile = "Ok"
Exit Function

FileOpenErr:
ErrorMsg err.Number, LoadResString(350)
ErrorReporte LoadResString(350)
OpenPrgFile = "NotOk"
Close #14

End Function
Public Sub SaveChanges(WOrigen As String)

'guarda los cambios producidos en cualquiera de los
'modulos del programa.

Dim Msg As String, Msg0 As String, Msg3 As String, Msg4 As String
Dim Style, Title, Response

Select Case WOrigen
    Case "EST1"
        Msg0 = "El contenido de la Estación 01 ha cambiado!."
        Msg3 = " "
        Msg4 = "Desea guardar los cambios antes de salir?"
        Msg = Msg0 & " " & Msg3 & " " & Msg4
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Rm100 - información."
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            Call Est01.E1Save_Click
        End If
    
    Case "EST2"
        Msg0 = "El contenido de la Estación 02 ha cambiado!."
        Msg3 = " "
        Msg4 = "Desea guardar los cambios antes de salir?"
        Msg = Msg0 & " " & Msg3 & " " & Msg4
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Rm100 - información."
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            Call Est02.E2Save_Click
        End If
    
    Case "TANDA"
        Msg0 = "El contenido de la Tanda ha cambiado!."
        Msg3 = " "
        Msg4 = "Desea guardar los cambios antes de salir?"
        Msg = Msg0 & " " & Msg3 & " " & Msg4
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Rm100 - información."
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            Call Tanda01.T1Save_Click
        End If
    
    Case "PROGTANDA"
        Msg0 = "El contenido de la Programación de Tandas"
        Msg3 = "ha cambiado!."
        Msg4 = "Desea guardar los cambios antes de salir?"
        Msg = Msg0 & " " & Msg3 & " " & Msg4
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Rm100 - información."
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            Call Prg01.P1Save_Click
        End If
End Select

End Sub

Function SaveConfigFile(Data As ConfigRecord) As String

NumReg = 1

'abrimos el archivo de configuracion
On Error GoTo err
Open App.path & AppConfigDir & AppConfigFile For Random As #33 Len = Len(Data)

'seteamos los datos de la configuracion a guardar
Data.id = NumReg
Data.ConfigHeader = CfgFHeader

Put #33, NumReg, Data
Close #33
SaveConfigFile = "Ok"
Exit Function

'------------------------------------------
err:
SaveConfigFile = "NotOk"
Close #33

End Function

Function SavePHFile(PHFileName As String) As String

Dim OIndex, nIndex
Dim MaxIndex
Dim Z As Integer

If PHFileName = "" Or PHFileName = " " Then
    SavePHFile = "NotOk"
    Exit Function
End If

'check the file for correct extension
If LCase(StripExtFromFile(PHFileName)) = AppPHFileExt Then
    PHFileName = PHFileName
Else
    PHFileName = StripFileFromExt(PHFileName) & AppPHFileExt
End If

On Error GoTo FileSaveErr
'save the data into the ph file
Open PHFileName For Random As #10 Len = Len(PHData)

nIndex = 0
MaxIndex = 23
'contador de proceso en accion
For Z = nIndex To MaxIndex
    PHData.filename = FrmTime.Text1(Z).text
    PHData.FileLounch = FrmTime.Text2(Z).text
    NumReg = Z + 1
    Put #10, NumReg, PHData
    nIndex = Z
    If nIndex > MaxIndex Then
        Exit For
    End If
Next Z
'close the ph file
Close #10
SavePHFile = "Ok"
Exit Function

FileSaveErr:
ErrorMsg err.Number, LoadResString(346)
ErrorReporte LoadResString(346)
SavePHFile = "NotOk"
Close #10

End Function

Function OpenTandaFile(FNombre As String) As String

'dimensiones de OpenTandaFile()
Dim LastRecord As Long
Dim NewCont As Integer
Dim NumRg As Integer
Dim RTime As String
Dim OTime As String
Dim NTime As String
Dim Op1 As Double
Dim Op2 As Double
Dim Suma As Double
Dim X As Integer
Dim ONum As Integer, NNum As Integer
Dim TxtKey As String, NewKey As String
Dim TMint As Double
Dim ItmX As ListItem

If FileExist(FNombre) = False Then
    OpenTandaFile = "NotOk"
    Exit Function
End If

'On Error GoTo FileOpenErr
FNombre = Trim(FNombre)

'//// lets open the tanda file for read the data
Open FNombre For Random As #10 Len = Len(TndData)

'//// lets get the record len
LastRecord = LOF(10) \ Len(TndData)

'//// lets set the progressbas
Tanda01.Prbar1.Min = 0
Tanda01.Prbar1.Max = LastRecord
Tanda01.Prbar1.Visible = True
Tanda01.Prbar1.Value = 0

'//// lets get out the data
For X = 1 To LastRecord
    NumRg = X
    'get the file data
    Get #10, NumRg, TndData
    ONum = Tanda01.T1View.ListItems.count
    NNum = ONum + 1
    TxtKey = "r"
    NewKey = TxtKey & NNum
    'put the data in a list
    Set ItmX = Tanda01.T1View.ListItems.Add(NNum, NewKey, Trim(TndData.Direccion))
    ItmX.SubItems(1) = Trim(TndData.FNType)
    ItmX.SubItems(2) = Trim(TndData.Name)
    ItmX.SubItems(3) = Trim(TndData.Duracion)
        OTime = Trim(Tanda01.Ltime.Caption)
        NTime = Trim(TndData.Duracion)
        Op1 = ConvMinToSec(OTime)
        Op2 = ConvMinToSec(NTime)
        TMint = CDbl(Trim(Tanda01.Intr.text))
        Suma = Op1 + Op2
        Suma = (Suma - TMint) + 1
        RTime = ConvSecToMin(Suma)
        SetSumTime RTime, 1
        Tanda01.Ltime.Caption = RTime
    ItmX.SubItems(4) = Trim(TndData.Hora)
    ItmX.SubItems(5) = Trim(TndData.DireccionX)
    ItmX.SubItems(6) = Trim(TndData.FNTypeX)
    ItmX.SubItems(7) = Trim(TndData.NameX)
    ItmX.SubItems(8) = Trim(TndData.DuracionX)
    ItmX.SubItems(9) = Trim(TndData.HoraX)
    Tanda01.Prbar1.Value = X
Next X

'//// close the tanda file
Close #10
OpenTandaFile = "Ok"
Tanda01.Prbar1.Visible = False
Exit Function

FileOpenErr:
ErrorMsg err.Number, LoadResString(341)
ErrorReporte LoadResString(341)
OpenTandaFile = "NotOk"
Close #10

End Function

Function SavePrgFile(WPrgFileName As String) As String

Dim nIndex As Integer
Dim MaxIndex As Integer
Dim Z As Integer

If WPrgFileName = "" Or WPrgFileName = " " Then
    SavePrgFile = "NotOk"
    Exit Function
End If

'check the file for correct extension
If LCase(StripExtFromFile(WPrgFileName)) = AppPrgFileExt Then
    WPrgFileName = WPrgFileName
Else
    WPrgFileName = StripFileFromExt(WPrgFileName) & AppPrgFileExt
End If

On Error GoTo FileSaveErr
'save the data into the ph file
Open WPrgFileName For Random As #15 Len = Len(PrgData)

nIndex = 0
MaxIndex = 23
'contador de proceso en accion
For Z = nIndex To MaxIndex
    PrgData.TndFileName = Est12Data.PF(Z).Caption
    PrgData.TndFileCaption = Est12Data.PC(Z).Caption
    PrgData.TndDuracion = Est12Data.PD(Z).Caption
    NumReg = Z + 1
    Put #15, NumReg, PrgData
    nIndex = Z
    If nIndex > MaxIndex Then
        Exit For
    End If
Next Z
'close the ph file
Close #15
SavePrgFile = "Ok"
Exit Function

FileSaveErr:
ErrorMsg err.Number, LoadResString(349)
ErrorReporte LoadResString(349)
SavePrgFile = "NotOk"
Close #15

End Function

Function SaveTandaFile(TandaFileName As String) As String

Dim OIndex, nIndex
Dim MaxIndex
Dim Z As Integer

If TandaFileName = "" Or TandaFileName = " " Then
    SaveTandaFile = "NotOk"
    Exit Function
End If

'check the file for correct extension
If LCase(StripExtFromFile(TandaFileName)) = AppTndFileExt Then
    TandaFileName = TandaFileName
Else
    TandaFileName = StripFileFromExt(TandaFileName) & AppTndFileExt
End If

On Error GoTo FileSaveErr
'save the data into the tanda file
Open TandaFileName For Random As #10 Len = Len(TndData)

'primero seleccionamos el primer item de la lista
Tanda01.T1View.ListItems.Item(1).Selected = True
'procesamos cada uno de los items dentro de la lista
OIndex = Tanda01.T1View.SelectedItem.index
MaxIndex = Tanda01.T1View.ListItems.count
'contador de proceso en accion
For Z = OIndex To MaxIndex
    TndData.id = Z
    TndData.Direccion = Tanda01.T1View.SelectedItem.text
    TndData.FNType = Tanda01.T1View.SelectedItem.SubItems(1)
    TndData.Name = Tanda01.T1View.SelectedItem.SubItems(2)
    TndData.Duracion = Tanda01.T1View.SelectedItem.SubItems(3)
    TndData.Hora = Tanda01.T1View.SelectedItem.SubItems(4)
    TndData.DireccionX = Tanda01.T1View.SelectedItem.SubItems(5)
    TndData.FNTypeX = Tanda01.T1View.SelectedItem.SubItems(6)
    TndData.NameX = Tanda01.T1View.SelectedItem.SubItems(7)
    TndData.DuracionX = Tanda01.T1View.SelectedItem.SubItems(8)
    TndData.HoraX = Tanda01.T1View.SelectedItem.SubItems(9)
    NumReg = Z
    'guardamos los datos a medida que examinamos los items de
    'la lista en la Tanda.
    Put #10, NumReg, TndData
    nIndex = Z + 1
    If nIndex > MaxIndex Then
        Exit For
    Else
        Tanda01.T1View.ListItems.Item(nIndex).Selected = True
    End If
Next Z
'close the tanda file
Close #10
SaveTandaFile = "Ok"
Exit Function

FileSaveErr:
ErrorMsg err.Number, LoadResString(150)
ErrorReporte LoadResString(150)
SaveTandaFile = "NotOk"
Close #10

End Function

Function GetWPos(WEst As Long, WMode As String) As String

'esta funcion extrae las posiciones width,left,top,height
'de los archivos y ordena las ventanas y sus controles
'en las estaciones 01 y 02

Dim i As Integer
Dim Data1 As Long, Data2 As Long, Data3 As Long, Data4 As Long, start As Long

Select Case WEst
    Case 1
        Select Case WMode
            Case "3x3", "Default" ' ********************************************
                On Error GoTo er
                Open App.path & AppDataDir & AppEst1VizFileV For Input As #33
                For i = 0 To 21
                    Input #33, Data1, Data2, Data3, Data4
                    Est01.E11(i).Height = Data1 'tamano vertical
                    Est01.E11(i).Left = Data2 'posicion hacia la izquierda
                    Est01.E11(i).Top = Data3  'posicion hacia arriba
                    Est01.E11(i).Width = Data4 'tamaño horizontal
                Next i
                Close #33
                'ordenamos los demas controles
                Est01.E1Play.Height = 375: Est01.E1Play.Left = 120
                Est01.E1Play.Top = 8400: Est01.E1Play.Width = 735
                Est01.E1Pause.Height = 375: Est01.E1Pause.Left = 840
                Est01.E1Pause.Top = 8400: Est01.E1Pause.Width = 615
                Est01.E1Stop.Height = 375: Est01.E1Stop.Left = 1440
                Est01.E1Stop.Top = 8400: Est01.E1Stop.Width = 735
                Est01.E1New.Height = 375: Est01.E1New.Left = 2520
                Est01.E1New.Top = 8400: Est01.E1New.Width = 375
                Est01.E1Open.Height = 375: Est01.E1Open.Left = 2880
                Est01.E1Open.Top = 8400: Est01.E1Open.Width = 375
                Est01.E1Save.Height = 375: Est01.E1Save.Left = 3240
                Est01.E1Save.Top = 8400: Est01.E1Save.Width = 375
                Est01.Pn.Height = 255: Est01.Pn.Left = 120
                Est01.Pn.Top = 960: Est01.Pn.Width = 255
                start = 360
                For i = 0 To 8
                    Est01.P11(i).Height = 255
                    Est01.P11(i).Left = start
                    Est01.P11(i).Top = 960
                    Est01.P11(i).Width = 375
                    start = Est01.P11(i).Left + 360
                Next i
                'Est01.Command1.Height = 255: Est01.Command1.Left = 120
                'Est01.Command1.Top = 1320: Est01.Command1.Width = 375
                'Est01.Command1.Enabled = False
                'Est01.Command2.Height = 255: Est01.Command2.Left = 3120
                'Est01.Command2.top = 1320: Est01.Command2.Width = 495
                'Est01.E1Shape.Height = 495: Est01.E1Shape.Left = 60
                'Est01.E1Shape.Top = 8340: Est01.E1Shape.Width = 3615
                GetWPos = "Ok"
                
            Case "4x4h" '*******************************************************
                On Error GoTo er
                Open App.path & AppDataDir & AppEst1VizFileH For Input As #33
                For i = 0 To 21
                    Input #33, Data1, Data2, Data3, Data4
                    Est01.E11(i).Height = Data1 'tamano vertical
                    Est01.E11(i).Left = Data2 'posicion hacia la izquierda
                    Est01.E11(i).Top = Data3  'posicion hacia arriba
                    Est01.E11(i).Width = Data4 'tamaño horizontal
                Next i
                Close #33
                'ordenamos los demas controles
                Est01.E1Play.Height = 375: Est01.E1Play.Left = 3960
                Est01.E1Play.Top = 3780: Est01.E1Play.Width = 735
                Est01.E1Pause.Height = 375: Est01.E1Pause.Left = 4680
                Est01.E1Pause.Top = 3780: Est01.E1Pause.Width = 615
                Est01.E1Stop.Height = 375: Est01.E1Stop.Left = 5280
                Est01.E1Stop.Top = 3780: Est01.E1Stop.Width = 735
                Est01.E1New.Height = 375: Est01.E1New.Left = 6300
                Est01.E1New.Top = 3780: Est01.E1New.Width = 375
                Est01.E1Open.Height = 375: Est01.E1Open.Left = 6660
                Est01.E1Open.Top = 3780: Est01.E1Open.Width = 375
                Est01.E1Save.Height = 375: Est01.E1Save.Left = 7020
                Est01.E1Save.Top = 3780: Est01.E1Save.Width = 375
                Est01.Pn.Height = 255: Est01.Pn.Left = 3900
                Est01.Pn.Top = 600: Est01.Pn.Width = 315
                start = 4200
                For i = 0 To 8
                    Est01.P11(i).Height = 255
                    Est01.P11(i).Left = start
                    Est01.P11(i).Top = 600
                    Est01.P11(i).Width = 375
                    start = Est01.P11(i).Left + 360
                Next i
                'Est01.Command1.Height = 255: Est01.Command1.Left = 3900
                'Est01.Command1.Top = 180: Est01.Command1.Width = 375
                'Est01.Command1.Enabled = True
                'Est01.Command2.Height = 255: Est01.Command2.Left = 6960
                'Est01.Command2.top = 180: Est01.Command2.Width = 495
                'Est01.E1Shape.Height = 495: Est01.E1Shape.Left = 3840
                'Est01.E1Shape.Top = 3720: Est01.E1Shape.Width = 3675
                GetWPos = "Ok"
                
            Case "4x4v" '*******************************************************
                On Error GoTo er
                Open App.path & AppDataDir & AppEst1VizFileH For Input As #33
                For i = 0 To 21
                    Input #33, Data1, Data2, Data3, Data4
                    Est01.E11(i).Height = Data1 'tamano vertical
                    Est01.E11(i).Left = Data2 'posicion hacia la izquierda
                    Est01.E11(i).Top = Data3  'posicion hacia arriba
                    Est01.E11(i).Width = Data4 'tamaño horizontal
                Next i
                Close #33
                'ordenamos los demas controles
                Est01.E1Play.Height = 375: Est01.E1Play.Left = 3960
                Est01.E1Play.Top = 3780: Est01.E1Play.Width = 735
                Est01.E1Pause.Height = 375: Est01.E1Pause.Left = 4680
                Est01.E1Pause.Top = 3780: Est01.E1Pause.Width = 615
                Est01.E1Stop.Height = 375: Est01.E1Stop.Left = 5280
                Est01.E1Stop.Top = 3780: Est01.E1Stop.Width = 735
                Est01.E1New.Height = 375: Est01.E1New.Left = 6300
                Est01.E1New.Top = 3780: Est01.E1New.Width = 375
                Est01.E1Open.Height = 375: Est01.E1Open.Left = 6660
                Est01.E1Open.Top = 3780: Est01.E1Open.Width = 375
                Est01.E1Save.Height = 375: Est01.E1Save.Left = 7020
                Est01.E1Save.Top = 3780: Est01.E1Save.Width = 375
                Est01.Pn.Height = 255: Est01.Pn.Left = 3900
                Est01.Pn.Top = 600: Est01.Pn.Width = 315
                start = 4200
                For i = 0 To 8
                    Est01.P11(i).Height = 255
                    Est01.P11(i).Left = start
                    Est01.P11(i).Top = 600
                    Est01.P11(i).Width = 375
                    start = Est01.P11(i).Left + 360
                Next i
                'Est01.Command1.Height = 255: Est01.Command1.Left = 3900
                'Est01.Command1.Top = 180: Est01.Command1.Width = 375
                'Est01.Command1.Enabled = True
                'Est01.Command2.Height = 255: Est01.Command2.Left = 6960
                'Est01.Command2.top = 180: Est01.Command2.Width = 495
                'Est01.E1Shape.Height = 495: Est01.E1Shape.Left = 3840
                'Est01.E1Shape.Top = 3720: Est01.E1Shape.Width = 3675
                GetWPos = "Ok"
                
        End Select
    Case 2 '*******************************************************************
        Select Case WMode
            Case "3x3", "Default" '********************************************
                On Error GoTo er
                Open App.path & AppDataDir & AppEst2VizFileV For Input As #33
                For i = 0 To 21
                    Input #33, Data1, Data2, Data3, Data4
                    Est02.E21(i).Height = Data1 'tamano vertical
                    Est02.E21(i).Left = Data2 'posicion hacia la izquierda
                    Est02.E21(i).Top = Data3  'posicion hacia arriba
                    Est02.E21(i).Width = Data4 'tamaño horizontal
                Next i
                Close #33
                'ordenamos los demas controles
                Est02.E2Play.Height = 375: Est02.E2Play.Left = 120
                Est02.E2Play.Top = 8400: Est02.E2Play.Width = 735
                Est02.E2Pause.Height = 375: Est02.E2Pause.Left = 840
                Est02.E2Pause.Top = 8400: Est02.E2Pause.Width = 615
                Est02.E2Stop.Height = 375: Est02.E2Stop.Left = 1440
                Est02.E2Stop.Top = 8400: Est02.E2Stop.Width = 735
                Est02.E2New.Height = 375: Est02.E2New.Left = 2520
                Est02.E2New.Top = 8400: Est02.E2New.Width = 375
                Est02.E2Open.Height = 375: Est02.E2Open.Left = 2880
                Est02.E2Open.Top = 8400: Est02.E2Open.Width = 375
                Est02.E2Save.Height = 375: Est02.E2Save.Left = 3240
                Est02.E2Save.Top = 8400: Est02.E2Save.Width = 375
                Est02.Pn.Height = 255: Est02.Pn.Left = 120
                Est02.Pn.Top = 960: Est02.Pn.Width = 255
                start = 360
                For i = 0 To 8
                    Est02.P21(i).Height = 255
                    Est02.P21(i).Left = start
                    Est02.P21(i).Top = 960
                    Est02.P21(i).Width = 375
                    start = Est02.P21(i).Left + 360
                Next i
                Est02.Command1.Height = 255: Est02.Command1.Left = 120
                Est02.Command1.Top = 1320: Est02.Command1.Width = 375
                Est02.Command1.Enabled = False
                'Est02.Command2.Height = 255: Est02.Command2.Left = 3120
                'Est02.Command2.top = 1320: Est02.Command2.Width = 495
                'Est02.E2Shape.Height = 495: Est02.E2Shape.Left = 60
                'Est02.E2Shape.top = 8340: Est02.E2Shape.Width = 3615
                GetWPos = "Ok"
                
            Case "4x4h" '*******************************************************
                On Error GoTo er
                Open App.path & AppDataDir & AppEst2VizFileH For Input As #33
                For i = 0 To 21
                    Input #33, Data1, Data2, Data3, Data4
                    Est02.E21(i).Height = Data1 'tamano vertical
                    Est02.E21(i).Left = Data2 'posicion hacia la izquierda
                    Est02.E21(i).Top = Data3  'posicion hacia arriba
                    Est02.E21(i).Width = Data4 'tamaño horizontal
                Next i
                Close #33
                'ordenamos los demas controles
                Est02.E2Play.Height = 375: Est02.E2Play.Left = 3960
                Est02.E2Play.Top = 3780: Est02.E2Play.Width = 735
                Est02.E2Pause.Height = 375: Est02.E2Pause.Left = 4680
                Est02.E2Pause.Top = 3780: Est02.E2Pause.Width = 615
                Est02.E2Stop.Height = 375: Est02.E2Stop.Left = 5280
                Est02.E2Stop.Top = 3780: Est02.E2Stop.Width = 735
                Est02.E2New.Height = 375: Est02.E2New.Left = 6300
                Est02.E2New.Top = 3780: Est02.E2New.Width = 375
                Est02.E2Open.Height = 375: Est02.E2Open.Left = 6660
                Est02.E2Open.Top = 3780: Est02.E2Open.Width = 375
                Est02.E2Save.Height = 375: Est02.E2Save.Left = 7020
                Est02.E2Save.Top = 3780: Est02.E2Save.Width = 375
                Est02.Pn.Height = 255: Est02.Pn.Left = 3900
                Est02.Pn.Top = 600: Est02.Pn.Width = 315
                start = 4200
                For i = 0 To 8
                    Est02.P21(i).Height = 255
                    Est02.P21(i).Left = start
                    Est02.P21(i).Top = 600
                    Est02.P21(i).Width = 375
                    start = Est02.P21(i).Left + 360
                Next i
                Est02.Command1.Height = 255: Est02.Command1.Left = 3900
                Est02.Command1.Top = 180: Est02.Command1.Width = 375
                Est02.Command1.Enabled = True
                'Est02.Command2.Height = 255: Est02.Command2.Left = 6960
                'Est02.Command2.top = 180: Est02.Command2.Width = 495
                'Est02.E2Shape.Height = 495: Est02.E2Shape.Left = 3840
                'Est02.E2Shape.top = 3720: Est02.E2Shape.Width = 3675
                GetWPos = "Ok"
                
            Case "4x4v" '*******************************************************
                On Error GoTo er
                Open App.path & AppDataDir & AppEst2VizFileH For Input As #33
                For i = 0 To 21
                    Input #33, Data1, Data2, Data3, Data4
                    Est02.E21(i).Height = Data1 'tamano vertical
                    Est02.E21(i).Left = Data2 'posicion hacia la izquierda
                    Est02.E21(i).Top = Data3  'posicion hacia arriba
                    Est02.E21(i).Width = Data4 'tamaño horizontal
                Next i
                Close #33
                'ordenamos los demas controles
                Est02.E2Play.Height = 375: Est02.E2Play.Left = 3960
                Est02.E2Play.Top = 3780: Est02.E2Play.Width = 735
                Est02.E2Pause.Height = 375: Est02.E2Pause.Left = 4680
                Est02.E2Pause.Top = 3780: Est02.E2Pause.Width = 615
                Est02.E2Stop.Height = 375: Est02.E2Stop.Left = 5280
                Est02.E2Stop.Top = 3780: Est02.E2Stop.Width = 735
                Est02.E2New.Height = 375: Est02.E2New.Left = 6300
                Est02.E2New.Top = 3780: Est02.E2New.Width = 375
                Est02.E2Open.Height = 375: Est02.E2Open.Left = 6660
                Est02.E2Open.Top = 3780: Est02.E2Open.Width = 375
                Est02.E2Save.Height = 375: Est02.E2Save.Left = 7020
                Est02.E2Save.Top = 3780: Est02.E2Save.Width = 375
                Est02.Pn.Height = 255: Est02.Pn.Left = 3900
                Est02.Pn.Top = 600: Est02.Pn.Width = 315
                start = 4200
                For i = 0 To 8
                    Est02.P21(i).Height = 255
                    Est02.P21(i).Left = start
                    Est02.P21(i).Top = 600
                    Est02.P21(i).Width = 375
                    start = Est02.P21(i).Left + 360
                Next i
                Est02.Command1.Height = 255: Est02.Command1.Left = 3900
                Est02.Command1.Top = 180: Est02.Command1.Width = 375
                Est02.Command1.Enabled = True
                'Est02.Command2.Height = 255: Est02.Command2.Left = 6960
                'Est02.Command2.top = 180: Est02.Command2.Width = 495
                'Est02.E2Shape.Height = 495: Est02.E2Shape.Left = 3840
                'Est02.E2Shape.top = 3720: Est02.E2Shape.Width = 3675
                GetWPos = "Ok"
                
        End Select
    Case Else '****************************************************************
        GetWPos = "NotOk"
End Select
Exit Function

er:
Close #33
ErrorMsg err.Number, LoadResString(150) & " " & LoadResString(999)
ErrorReporte LoadResString(150)
GetWPos = "NotOk"
Call TopMenu.EndApp

End Function

Sub PutState()

'Sub para guardar la informacion del estado en el que se
'encontraba el programa la ultima vez que se lo utilizo.

Dim UltimaTanda As String * 255
Dim UltimaEst1 As String * 255
Dim UltimaEst2 As String * 255
Dim UltimaProg As String * 255
Dim UltimaWin As String * 8
Dim DestinoTanda As String
Dim DestinoEst1 As String
Dim DestinoEst2 As String
Dim DestinoProg As String
Dim DestinoWin As String
Dim LngTxt As Long
Dim EstNum As Long
Dim OrigenTanda As String
Dim OrigenEst1 As String
Dim OrigenEst2 As String
Dim OrigenProg As String
Dim OrigenWin As String
Dim CaBeCera As String
Dim ConvertTx As String
Dim NumReg As Long

On Error GoTo FileSaveErr
Open App.path & AppConfigDir & AppStateFile For Random As #24 Len = Len(StateData)

NumReg = 1

'chequeamos
If TopMenu.View3x3.Checked = True Or TopMenu.ViewDefault.Checked = True Then
    UltimaWin = "Default"
Else
    If TopMenu.View4x4h.Checked = True Then
        UltimaWin = "4x4h"
    Else
        If TopMenu.View4x4v.Checked = True Then
            UltimaWin = "4x4v"
        Else
            UltimaWin = "Default"
        End If
    End If
End If

'ultima ESTACION 01 utilizada
ConvertTx = Trim(Est01.Fn.Caption)
If ConvertTx = "" Or ConvertTx = " " Then
    UltimaEst1 = ""
Else
    UltimaEst1 = ConvertTx
End If

'ultima ESTACION 02 utilizada
ConvertTx = Trim(Est02.Fn.Caption)
If ConvertTx = "" Or ConvertTx = " " Then
    UltimaEst2 = ""
Else
    UltimaEst2 = ConvertTx
End If

'ultima TANDA utilizada
ConvertTx = Trim(Tanda01.Fn.Caption)
If ConvertTx = "" Or ConvertTx = " " Then
    UltimaTanda = ""
Else
    UltimaTanda = ConvertTx
End If

'ultima PROGTANDA utilizada
ConvertTx = Trim(Prg01.Fn.Caption)
If ConvertTx = "" Or ConvertTx = " " Then
    UltimaProg = ""
Else
    UltimaProg = ConvertTx
End If

'-------------------------------------------
'encriptamos los datos
DestinoWin = CipherData(StatePassW, UltimaWin)
DestinoEst1 = CipherData(StatePassW, UltimaEst1)
DestinoEst2 = CipherData(StatePassW, UltimaEst2)
DestinoTanda = CipherData(StatePassW, UltimaTanda)
DestinoProg = CipherData(StatePassW, UltimaProg)

StateData.StateHeader = StFHeader   'ponemos la cebecera sin encriptar
StateData.LastWinOrder = DestinoWin
StateData.LastEst1File = DestinoEst1
StateData.LastEst2File = DestinoEst2
StateData.LastTndFile = DestinoTanda
StateData.LastPrgFile = DestinoProg

Put #24, NumReg, StateData   'guardamos los registros
Close #8    'cerramos el archivo
Exit Sub

FileSaveErr:
ErrorMsg err.Number, LoadResString(342)
ErrorReporte LoadResString(342)
Close #24

End Sub

Function GetState() As String

'Sub para extraer la informacion del estado en el que se
'encontraba el programa la ultima vez que se lo utilizo.

Dim NumReg As Long, Result As String
Dim UltimaTanda As String, UltimaEst1 As String, UltimaEst2 As String
Dim UltimaProg As String, UltimaWin As String
Dim OrigenTanda As String, OrigenEst1 As String, OrigenEst2 As String
Dim OrigenProg As String, OrigenWin As String, CaBeCera As String

On Error GoTo FileOpenErr
Open App.path & AppConfigDir & AppStateFile For Random As #8 Len = Len(StateData)

NumReg = 1
Get #8, NumReg, StateData   'extraemos los registros

CaBeCera = Trim(StateData.StateHeader)    'verificamos la cabecera del archivo
If CaBeCera = StFHeader Then
    OrigenWin = StateData.LastWinOrder
    OrigenEst1 = StateData.LastEst1File
    OrigenEst2 = StateData.LastEst2File
    OrigenTanda = StateData.LastTndFile
    OrigenProg = StateData.LastPrgFile
Else
    'cabecera de archivo incorrecta
    Close #8
    GetState = "NotOk"
    Exit Function
End If

Close #8    'cerramos el archivo

'-------------------------------------------
'desencriptamos los datos
UltimaWin = DecipherData(StatePassW, OrigenWin)
UltimaEst1 = DecipherData(StatePassW, OrigenEst1)
UltimaEst2 = DecipherData(StatePassW, OrigenEst2)
UltimaTanda = DecipherData(StatePassW, OrigenTanda)
UltimaProg = DecipherData(StatePassW, OrigenProg)

'actualizamos los datos
UltimaWin = Trim(UltimaWin)
UltimaEst1 = Trim(UltimaEst1)
UltimaEst2 = Trim(UltimaEst2)
UltimaTanda = Trim(UltimaTanda)
UltimaProg = Trim(UltimaProg)

'comenzamos la actualizacion de datos
'orden y posicionamiento de las ventanas del programa
If UltimaWin = "3x3" Or UltimaWin = "Default" Then
    ShowWindow "Startup1"
Else
    If UltimaWin = "4x4h" Then
        ShowWindow "Startup2"
    Else
        If UltimaWin = "4x4v" Then
            ShowWindow "Startup3"
        Else
            ShowWindow "Startup1"
        End If
    End If
End If

'ultimo archivo de estacion 01 utilizado
If UltimaEst1 = "" Or UltimaEst1 = " " Then
    'nothing to do
Else
    Result = OpenEstFile(1, 1, UltimaEst1)
    If Result = "NotOk" Then
        Est01.Fn.Caption = ""
    Else
        Est01.Fn.Caption = UltimaEst1
    End If
End If

'ultimo archivo de estacion 02 utilizado
If UltimaEst2 = "" Or UltimaEst2 = " " Then
    'nothing to do
Else
    Result = OpenEstFile(2, 1, UltimaEst2)
    If Result = "NotOk" Then
        Est02.Fn.Caption = ""
    Else
        Est02.Fn.Caption = UltimaEst2
    End If
End If

'ultimo archivo de tanda utilizado
If UltimaTanda = "" Or UltimaTanda = " " Then
    'nothing to do
Else
    Result = OpenTandaFile(UltimaTanda)
    If Result = "NotOk" Then
        Tanda01.Fn.Caption = ""
    Else
        Tanda01.Fn.Caption = UltimaTanda
    End If
End If

'ultimo archivo de programacion de tandas utilizado
If UltimaProg = "" Or UltimaProg = " " Then
    'nothing to do
Else
    Result = OpenPrgFile(UltimaProg)
    If Result = "NotOk" Then
        Prg01.Fn.Caption = ""
    Else
        Prg01.Fn.Caption = UltimaProg
        Dim lnText As Long, NewName As String
        lnText = Len(UltimaProg)
        If lnText > 60 Then
            NewName = Left$(UltimaProg, 3) & " ... " & Right$(UltimaProg, 50)
        Else
            NewName = UltimaProg
        End If
        Prg01.LblName.Caption = NewName
        Prg01.LblName.ForeColor = &HFFFF00    'verde claro
    End If
End If

GetState = "Ok"
Exit Function

FileOpenErr:
'ErrorMsg Err.Number, LoadResString(337) ' lo quitamos para que no moleste
ErrorReporte LoadResString(337)
Close #8
GetState = "NotOk"

End Function

Sub CrearReporte(MusicName As String, filename As String)

'--- Definimos algunas variables
Dim Prueba
Dim Line01, Line02, Line03, Line04
Dim ArchNum
Dim NewNum
Dim ReportFile

'//////////////////////////////////////////////////
Init:
ArchNum = FreeFile
ReportFile = "\RMRpt" & Date$ & ".Rpt"
On Error GoTo NewReport 'Si no existe se crea uno nuevo
Open App.path & AppReportDir & ReportFile For Input As ArchNum  'Abrimos el archivo
Input #ArchNum, Prueba
Close ArchNum
EscribirReporte MusicName, filename
Exit Sub

'//////////////////////////////////////////////////
NewReport:
Close ArchNum
Close NewNum
Resume CreateNew
Exit Sub

'//////////////////////////////////////////////////
CreateNew:
NewNum = FreeFile
ReportFile = "\RMRpt" & Date$ & ".Rpt"
On Error GoTo Oupps
Open App.path & AppReportDir & ReportFile For Append As NewNum

Line01 = "    Radio Maker 1.0 - REPORTE DE REPRODUCCIONES FECHA: " & Date$
Line02 = " -----------------------------------------------------------------------"
Line03 = " " 'Aqui se colocan los datos del usuario registrado
Line04 = " " 'Aqui se colocan los demas datos del usuario registrado.

Print #NewNum, Line01
Print #NewNum, Line02
Print #NewNum, Line03
Print #NewNum, Line04
Close NewNum
GoSub Init
Exit Sub

'//////////////////////////////////////////////////
Oupps:
Close ArchNum
Close NewNum
ErrorMsg err.Number, LoadResString(335)
ErrorReporte LoadResString(335)
End Sub

Sub EscribirReporte(MusicName As String, filename As String)

Dim NewFile
Dim ReportFile

NewFile = FreeFile
ReportFile = "\RMRpt" & Date$ & ".Rpt"
If filename = "" Or filename = " " Then GoSub BadRpt
filename = LTrim(RTrim(UCase(filename)))

On Error GoTo Oupps
Open App.path & AppReportDir & ReportFile For Append As NewFile
Print #NewFile, "Nombre: " & MusicName & " - archivo: " & filename & " - Fecha: " & Date$ & " - Hora: " & time$
Close NewFile
Exit Sub

Oupps: '//////////////////////////////////////////////////
Close NewFile
ErrorMsg err.Number, LoadResString(336)
ErrorReporte LoadResString(336)
Exit Sub

BadRpt: '//////////////////////////////////////////////////
End Sub

Function GetCipherConfigData(DataIn As String) As String

Dim CipherIn As String
Dim CipherOut As String

If DataIn = " " Or DataIn = "" Then
    GetCipherConfigData = ""
    Exit Function
Else
    CipherIn = DataIn
End If

'desencriptamos
CipherOut = DecipherData(CipherPass, CipherIn)    'NEW
'Desencriptar CipherPass, CipherIn, CipherOut

GetCipherConfigData = CipherOut

End Function

Function GetUsrPassWord() As String

Dim DtIn As String
Dim DtOut As String

On Error GoTo er
Open App.path & AppConfigDir & AppUsrPwrFile For Input As #16
Input #16, DtIn
Close #16

'desencriptamos la clave de acceso del archivo
DtOut = DecipherData(UsrPass, DtIn)        'NEW

GetUsrPassWord = Trim(DtOut)
Exit Function

er:
Close #16
GetUsrPassWord = "NoPassWord"
Resume Continuar
Exit Function

Continuar:
End Function

Sub CerrarArchivoInicio()

Dim StartIn, StartOut

StartIn = LoadResString(555)
StartOut = LoadResString(556)

'abrimos el archivo de configuracion
On Error GoTo StartError
Open App.path & AppConfigDir & AppInitFile For Output As #16
Write #16, StartIn, StartOut
Close #16
Exit Sub

StartError:
'Si ocurre un error creamos el reporte del mismo en un archivo
Close #16
ErrorMsg err.Number, LoadResString(333)
ErrorReporte LoadResString(333)
Resume Finalizar
End

Finalizar:
End Sub

Sub CrearArchivoInicio()

Dim StartIn, StartOut

StartIn = LoadResString(555)
StartOut = "-"

'abrimos el archivo de configuracion
On Error GoTo StartError
Open App.path & AppConfigDir & AppInitFile For Output As #15
Write #15, StartIn, StartOut
Close #15
Exit Sub

StartError:
Close #15
ErrorMsg err.Number, LoadResString(334)
ErrorReporte LoadResString(334)
Resume Finalizar
End

Finalizar:
End Sub

Sub ErrorReporte(Mensaje As String)

Dim StartData As String, EndData As String

StartData = "---- " & Date & " - " & time & " -------------------------------------------."
EndData = "----------------------------------------------------------------------."

On Error GoTo ReporteError
Open App.path & AppConfigDir & AppErrFile For Append As #18
Print #18, StartData
Print #18, Mensaje
Print #18, EndData
Print #18, ""
Close #18
Exit Sub

ReporteError:
Resume Next

End Sub

Sub OpenCUEFile(EstNumber As Long, WFileName As String)

Dim DD1 As String * 8
Dim Data1 As String
Dim DD2 As String * 8
Dim Data2 As String
Dim DD3 As String * 4
Dim Data3 As String
Dim DD4 As String * 4
Dim Data4 As String
Dim DD5 As String * 30
Dim Data5 As String
Dim DD6 As String * 30
Dim Data6 As String
Dim X As Integer
Dim NumReg As Long
Dim CaBeCera As String

If FileExist(WFileName) = False Then
    Exit Sub
End If

On Error GoTo FileOpenErr
Open WFileName For Random As #10 Len = Len(CUEData)

NumReg = 1

Select Case EstNumber
    Case 1  '*** ESTACION 1
        Get #10, NumReg, CUEData
        'extraemos y verificamos la cabecera
        CaBeCera = Trim(CUEData.CueHeader)
        If CaBeCera = CUEFHeader Then
            '*** extraemos y desencriptamos
            Data1 = DecipherData(CUEFilePass, CUEData.DisplayCUEStartTime)
            Data2 = DecipherData(CUEFilePass, CUEData.DisplayCUEEndTime)
            Data3 = DecipherData(CUEFilePass, CUEData.DisplayCUEStartMark)
            Data4 = DecipherData(CUEFilePass, CUEData.DisplayCUELengthMark)
            Data5 = DecipherData(CUEFilePass, CUEData.CUEStartByte)
            Data6 = DecipherData(CUEFilePass, CUEData.CUEEndByte)
            '*** actualizamos los datos
            Est01.Text1.text = Trim(Data1)
            Est01.Text2.text = Trim(Data2)
            Est01.E1Pos.SelStart = Trim(Data3)
            Est01.E1Pos.SelLength = Trim(Data4)
            Est01.LblStartCUE.Caption = Trim(Data5)
            Est01.LblEndCue.Caption = Trim(Data6)
            '*** actualizamos el EQ
            For X = 0 To 10
                Est01.fxsc(X).Value = CUEData.EQValue(X)
            Next X
            '*** cerramos
            Close #10
            Exit Sub
        Else
            'cabecera de archivo incorrecta
            GoSub HeaderErr
        End If
    Case 2  '*** ESTACION 2
        Get #10, NumReg, CUEData
        'extraemos y verificamos la cabecera
        CaBeCera = Trim(CUEData.CueHeader)
        If CaBeCera = CUEFHeader Then
            '*** extraemos y desencriptamos
            Data1 = DecipherData(CUEFilePass, CUEData.DisplayCUEStartTime)
            Data2 = DecipherData(CUEFilePass, CUEData.DisplayCUEEndTime)
            Data3 = DecipherData(CUEFilePass, CUEData.DisplayCUEStartMark)
            Data4 = DecipherData(CUEFilePass, CUEData.DisplayCUELengthMark)
            Data5 = DecipherData(CUEFilePass, CUEData.CUEStartByte)
            Data6 = DecipherData(CUEFilePass, CUEData.CUEEndByte)
            '*** actualizamos los datos
            Est02.Text1.text = Trim(Data1)
            Est02.Text2.text = Trim(Data2)
            Est02.E2Pos.SelStart = Trim(Data3)
            Est02.E2Pos.SelLength = Trim(Data4)
            Est02.LblStartCUE.Caption = Trim(Data5)
            Est02.LblEndCue.Caption = Trim(Data6)
            '*** actualizamos el EQ
            For X = 0 To 10
                Est02.fxsc(X).Value = CUEData.EQValue(X)
            Next X
            '*** cerramos
            Close #10
            Exit Sub
        Else
            'cabecera de archivo incorrecta
            GoSub HeaderErr
        End If
End Select
Exit Sub

FileOpenErr:
GoSub Restore
Close #10
Exit Sub

HeaderErr:
ErrorMsg err.Number, LoadResString(151)
ErrorReporte LoadResString(151)
GoSub Restore
Close #10
Exit Sub

Restore:
Select Case EstNumber
    Case 1
        Est01.Text1.text = "00:00:00"
        Est01.Text2.text = "00:00:00"
        Est01.E1Pos.SelStart = 0
        Est01.E1Pos.SelLength = 0
        Est01.LblStartCUE.Caption = "0"
        Est01.LblEndCue.Caption = "0"
        'restore the values for EQ
        For X = 0 To 10
            Est01.fxsc(X).Value = 10
        Next X
        Est01.fxsc(10).Value = 18
    Case 2
        Est02.Text1.text = "00:00:00"
        Est02.Text2.text = "00:00:00"
        Est02.E2Pos.SelStart = 0
        Est02.E2Pos.SelLength = 0
        Est02.LblStartCUE.Caption = "0"
        Est02.LblEndCue.Caption = "0"
        'restore the values for EQ
        For X = 0 To 10
            Est02.fxsc(X).Value = 10
        Next X
        Est02.fxsc(10).Value = 18
End Select
Return

End Sub

Sub SaveCUEFile(EstNumber As Long, CUEFileName As String)

Dim DD1 As String * 8
Dim Data1 As String
Dim DD2 As String * 8
Dim Data2 As String
Dim DD3 As String * 4
Dim Data3 As String
Dim DD4 As String * 4
Dim Data4 As String
Dim DD5 As String * 30
Dim Data5 As String
Dim DD6 As String * 30
Dim Data6 As String
Dim X As Integer

If CUEFileName = "" Or CUEFileName = " " Then
    Exit Sub
End If

'check the file for correct extension
If LCase(StripExtFromFile(CUEFileName)) = AppCUEFileExt Then
    CUEFileName = CUEFileName
Else
    CUEFileName = StripFileFromExt(CUEFileName) & AppCUEFileExt
End If

On Error GoTo FileSaveErr
Open CUEFileName For Random As #10 Len = Len(CUEData)

NumReg = 1

Select Case EstNumber
    Case 1  '*** ESTACION 1
        '*** actualizamos los datos
        DD1 = Trim(Est01.Text1.text)
        DD2 = Trim(Est01.Text2.text)
        DD3 = Trim(Est01.E1Pos.SelStart)
        DD4 = Trim(Est01.E1Pos.SelLength)
        DD5 = Trim(Est01.LblStartCUE.Caption)
        DD6 = Trim(Est01.LblEndCue.Caption)
        '*** encriptamos
        Data1 = CipherData(CUEFilePass, DD1)
        Data2 = CipherData(CUEFilePass, DD2)
        Data3 = CipherData(CUEFilePass, DD3)
        Data4 = CipherData(CUEFilePass, DD4)
        Data5 = CipherData(CUEFilePass, DD5)
        Data6 = CipherData(CUEFilePass, DD6)
        '*** guardamos los datos
        CUEData.id = NumReg
        CUEData.CueHeader = CUEFHeader  'cabecera del archivo sin encriptar
        CUEData.DisplayCUEStartTime = Data1
        CUEData.DisplayCUEEndTime = Data2
        CUEData.DisplayCUEStartMark = Data3
        CUEData.DisplayCUELengthMark = Data4
        CUEData.CUEStartByte = Data5
        CUEData.CUEEndByte = Data6
        'estos van sin encriptar
        For X = 0 To 10
            CUEData.EQValue(X) = Est01.fxsc(X).Value
        Next X
        Put #10, NumReg, CUEData
        '*** cerramos el archivo
        Close #10
        Exit Sub
    Case 2  'ESTACION 2
        '*** actualizamos los datos
        DD1 = Trim(Est02.Text1.text)
        DD2 = Trim(Est02.Text2.text)
        DD3 = Trim(Est02.E2Pos.SelStart)
        DD4 = Trim(Est02.E2Pos.SelLength)
        DD5 = Trim(Est02.LblStartCUE.Caption)
        DD6 = Trim(Est02.LblEndCue.Caption)
        '*** encriptamos
        Data1 = CipherData(CUEFilePass, DD1)
        Data2 = CipherData(CUEFilePass, DD2)
        Data3 = CipherData(CUEFilePass, DD3)
        Data4 = CipherData(CUEFilePass, DD4)
        Data5 = CipherData(CUEFilePass, DD5)
        Data6 = CipherData(CUEFilePass, DD6)
        '*** guardamos los datos ya encriptados
        CUEData.id = NumReg
        CUEData.CueHeader = CUEFHeader  'cabecera del archivo sin encriptar
        CUEData.DisplayCUEStartTime = Data1
        CUEData.DisplayCUEEndTime = Data2
        CUEData.DisplayCUEStartMark = Data3
        CUEData.DisplayCUELengthMark = Data4
        CUEData.CUEStartByte = Data5
        CUEData.CUEEndByte = Data6
        'estos van sin encriptar
        For X = 0 To 10
            CUEData.EQValue(X) = Est02.fxsc(X).Value
        Next X
        Put #10, NumReg, CUEData
        '*** cerramos el archivo
        Close #10
        Exit Sub
End Select
Exit Sub

FileSaveErr:
ErrorMsg err.Number, LoadResString(343)
ErrorReporte LoadResString(343)
Close #10

End Sub

Function SetCipherConfigData(DataIn As String) As String

Dim CipherIn As String, CipherOut As String

If DataIn = " " Or DataIn = "" Then
    SetCipherConfigData = ""
    Exit Function
Else
    CipherIn = DataIn
End If

'encriptamos
CipherOut = CipherData(CipherPass, CipherIn) 'NEW
'Encriptar CipherPass, CipherIn, CipherOut

SetCipherConfigData = CipherOut

End Function

Function ErrorMsg(ErNumber As Long, wMsg As String)

Select Case ErNumber
    Case 7, 31001
        MsgBox LoadResString(135), vbCritical
    Case 52
        MsgBox LoadResString(136), vbCritical
    Case 53, 76
        MsgBox LoadResString(137), vbCritical
    Case 55
        MsgBox LoadResString(138), vbCritical
    Case 57
        MsgBox LoadResString(139), vbCritical
    Case 61
        MsgBox LoadResString(140), vbCritical
    Case 68
        MsgBox LoadResString(141), vbCritical
    Case 70
        MsgBox LoadResString(142), vbCritical
    Case 71
        MsgBox LoadResString(143), vbCritical
    Case 321
        MsgBox LoadResString(144), vbCritical
    Case 322, 735
        MsgBox LoadResString(145), vbCritical
    Case 31036
        MsgBox LoadResString(146), vbCritical
    Case 31037
        MsgBox LoadResString(147), vbCritical
    Case Else
        MsgBox wMsg & LoadResString(148) & ErNumber, vbCritical
End Select

End Function

Private Function OpenEstDataFile(ByVal EstNumber As Long, ByVal FNombre As String) As String

Dim NewCont As Integer
Dim NumRg As Integer

If FileExist(FNombre) = False Then
    OpenEstDataFile = "NotOk"
    Exit Function
End If

On Error GoTo FileOpenErr
FNombre = Trim(FNombre)
Open FNombre For Random As #10 Len = Len(EstData)

Select Case EstNumber
    Case 1  'ESTACION 1
        For NewCont = 0 To 21
            NumRg = NewCont + 1
            Get #10, NumRg, EstData
            
            Est01.E11(NewCont).Caption = Trim(EstData.CCaption)
            Est01.E11(NewCont).ToolTipText = " Duración: " & Trim(EstData.FDuracion) & " "
            Est12Data.N1(NewCont).Caption = Trim(EstData.FName)
            Est12Data.c1(NewCont).Caption = Trim(EstData.CCaption)
            Est12Data.D1(NewCont).Caption = Trim(EstData.FDuracion)
            Est12Data.V1(NewCont).Caption = Trim(EstData.FType)

        Next NewCont
        Close #10
        OpenEstDataFile = "OK"
        Exit Function
        
    Case 2  'ESTACION 2
        For NewCont = 0 To 21
            NumRg = NewCont + 1
            Get #10, NumRg, EstData
            
            Est02.E21(NewCont).Caption = Trim(EstData.CCaption)
            Est02.E21(NewCont).ToolTipText = " Duración: " & Trim(EstData.FDuracion) & " "
            Est12Data.N2(NewCont).Caption = Trim(EstData.FName)
            Est12Data.c2(NewCont).Caption = Trim(EstData.CCaption)
            Est12Data.D2(NewCont).Caption = Trim(EstData.FDuracion)
            Est12Data.V2(NewCont).Caption = Trim(EstData.FType)
            
        Next NewCont
        Close #10
        OpenEstDataFile = "OK"
        Exit Function
End Select
Exit Function

FileOpenErr:
ErrorMsg err.Number, LoadResString(338)
ErrorReporte LoadResString(338)
OpenEstDataFile = "NotOk"
Close #10

End Function

Private Function SaveEstDataFile(ByVal EstNumber As Long, ByVal PageNumber As Long, ByVal FNombre As String) As String

Dim Contador As Integer
Dim NumReg As Long

If FNombre = "" Or FNombre = " " Then
    SaveEstDataFile = "NotOk"
    Exit Function
End If

On Error GoTo FileSaveErr
FNombre = FNombre & PageNumber
Open FNombre For Random As #10 Len = Len(EstData)

Select Case EstNumber
    Case 1  'ESTACION 1
        For Contador = 0 To 21
            EstData.id = Contador + 1
            EstData.Control = "E" & PageNumber
            EstData.CCaption = Est12Data.c1(Contador).Caption
            EstData.FName = Est12Data.N1(Contador).Caption
            EstData.FType = Est12Data.V1(Contador).Caption
            EstData.FDuracion = Est12Data.D1(Contador).Caption
            NumReg = Contador + 1
            Put #10, NumReg, EstData
        Next Contador
        Close #10
        SaveEstDataFile = "OK"
        Exit Function
    Case 2  'ESTACION 2
        For Contador = 0 To 21
            EstData.id = Contador + 1
            EstData.Control = "E" & PageNumber
            EstData.CCaption = Est12Data.c2(Contador).Caption
            EstData.FName = Est12Data.N2(Contador).Caption
            EstData.FType = Est12Data.V2(Contador).Caption
            EstData.FDuracion = Est12Data.D2(Contador).Caption
            NumReg = Contador + 1
            Put #10, NumReg, EstData
        Next Contador
        Close #10
        SaveEstDataFile = "OK"
        Exit Function
End Select
Exit Function

FileSaveErr:
ErrorMsg err.Number, LoadResString(344)
ErrorReporte LoadResString(344)
SaveEstDataFile = "NotOk"
Close #10

End Function

Function OpenEstFile(ByVal EstNum As Long, ByVal PageNum As Long, ByVal FName As String) As String

Dim NewName As String, Result As String
Dim A$, B$, C$, D$, e$, f$, G$, H$, i$

If FileExist(FName) = False Then
    OpenEstFile = "NotOk"
    Exit Function
End If

On Error GoTo OpenFileErr
Open FName For Input As #13
Input #13, A$, B$, C$, D$, e$, f$, G$, H$, i$
Close #13

    Select Case PageNum
        Case 1
            NewName = Trim(A$)
        Case 2
            NewName = Trim(B$)
        Case 3
            NewName = Trim(C$)
        Case 4
            NewName = Trim(D$)
        Case 5
            NewName = Trim(e$)
        Case 6
            NewName = Trim(f$)
        Case 7
            NewName = Trim(G$)
        Case 8
            NewName = Trim(H$)
        Case 9
            NewName = Trim(i$)
    End Select

Result = OpenEstDataFile(EstNum, NewName)

OpenEstFile = "OK"
Exit Function

OpenFileErr:
ErrorMsg err.Number, LoadResString(339)
ErrorReporte LoadResString(339)
OpenEstFile = "NotOk"
Close #13

End Function

Function SaveEstFile(ByVal EstNum As Long, ByVal PageNum As Long, ByVal FName As String) As String

Dim Result As String
Dim Contador As Integer
Dim Sumador As Integer

If FName = "" Or FName = " " Then
    SaveEstFile = "NotOk"
    Exit Function
End If

'primero guardamos los datos de las estaciones
Result = SaveEstDataFile(EstNum, PageNum, FName)
'... y verificamos que este ok
If Result = "NotOK" Then
    GoSub SaveFileErr
End If

'ahora guardamos el archivo principal de la estacion
On Error GoTo SaveFileErr
Open FName For Output As #11
    For Contador = 0 To 8
        Sumador = Contador + 1
        FName = Trim(FName)
        Write #11, FName & Sumador
    Next Contador
Close #11

SaveEstFile = "OK"
Exit Function

SaveFileErr:
ErrorMsg err.Number, LoadResString(345)
ErrorReporte LoadResString(345)
SaveEstFile = "NotOk"
Close #11

End Function
