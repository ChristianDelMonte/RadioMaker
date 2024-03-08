Attribute VB_Name = "RMVoiceData"

'////////////////////////////////////////////////////////
'*
'*  ////////// VOICE & FX module for Vb.6+ ///////////
'*  ** this module depends on 100% of "RMBass.bas" **
'*  ********* and is for Radiomaker 1.0 only *********
'*
'*     Copyright (c) 1987-2008 Only development Inc.
'*
'///////////////////////////////////////////////////////

Option Explicit
'archivos de audio MASCULINAS
Public Const FnameHoraESP = "H_Esp_Comp.Pak"         'hora                    ESPAÑOL
Public Const FnameHoraEN = "H_Eng_Comp.Pak"          'hora                    INGLES
Public Const FnameMinutoESP = "M_Esp_Comp.Pak"       'minutos                 ESPAÑOL
Public Const FnameMinutoEN = "M_Eng_Comp.Pak"        'minutos                 INGLES
Public Const FnameTempESP = "T_Esp_Comp.Pak"       'temperatura en grados   ESPAÑOL
Public Const FnameTempEN = "T_Eng_Comp.Pak"        'temperatura en grados   INGLES
Public Const FnameHumeESP = "Hu_Esp_Comp.Pak"        'humedad                 ESPAÑOL
Public Const FnameHumeEN = "Hu_Eng_Comp.Pak"         'humedad                 INGLES

'archivos de datos de audio (regiones a reproducir) MASCULINAS
Public Const HDataFileName_Esp = "H_Esp_Comp.Dat"      'hora                     ESPAÑOL
Public Const HDataFileName_Eng = "H_Eng_Comp.Dat"      'hora                     INGLES
Public Const MDataFileName_Esp = "M_Esp_Comp.Dat"      'minutos                  ESPAÑOL
Public Const MDataFileName_Eng = "M_Eng_Comp.Dat"      'minutos                  INGLES
Public Const TDataFileName_Esp = "T_Esp_Comp.Dat"    'temperatura en grados    ESPAÑOL
Public Const TDataFileName_Eng = "T_Eng_Comp.Dat"    'temperatura en grados    INGLES
Public Const HuDataFileName_Esp = "Hu_Esp_Comp.Dat"    'humedad                  ESPAÑOL
Public Const HuDataFileName_Eng = "Hu_Eng_Comp.Dat"    'humedad                  INGLES

Public Const ConfigFileName = "\RMVoice.cfg"     'archivo gral de configuracion
Public Const DatabaseFileName = "\World.dbs"     'base de datos de provincias u estados

Public Type CFG_Data
    Id As Integer
    Lng_Id As Integer               'lenguaje Id 1=español 2=ingles, 3=frances, 4=italiano, 5=portugues
    State_Id As String * 50         'provincia o estado id - ver general database
    Temp_Mode As Integer            'modo de temperatura 1=farenhit 2=centigrados
    Voice_Id As Integer             '1=masculino 2=femenino 3=personalizado
    Temp_RTime As Integer           'tiempo de refresco de datos climaticos en milisseconds
    HVoicePath As String * 255      'path de voz personalizada (hora)
    MVoicePath As String * 255      'path de voz personalizada (minutos)
    TmpVoicePath As String * 255    'path de voz personalizada (temperatura)
    HumVoicePath As String * 255    'path de voz personalizada (humedad)
End Type

Public Type GEN_State_Database
    Id As Integer
    State_Desc As String * 50   'provincia o estado, descripcion o nombre
    State_URL As String * 255   'url para la extraccion de datos climaticos correspondientes
End Type

Public ConfigData As CFG_Data
Public StateDatabase As GEN_State_Database

Public Pstate As Boolean 'estado de reproduccion CUIDADO!!!
Public FileState As Boolean
Public LastReg As Integer
Public Result As String

Function SearchURL_from_StateName(WStateName As String) As String

Dim i As Integer

On Error GoTo err
Open App.Path & DatabaseFileName For Random As #16 Len = Len(StateDatabase)

LastReg = GetStateLastReg
For i = 1 To LastReg
    Get #16, i, StateDatabase
    If Trim(StateDatabase.State_Desc) = Trim(WStateName) Then
        SearchURL_from_StateName = Trim(StateDatabase.State_URL)
        Close #16
        Exit Function
    Else
        If i >= LastReg Then
            SearchURL_from_StateName = "not found"
        End If
    End If
Next i
Close #16
Exit Function

'/// if there is an error ------------------------------------------
err:
DisplayMsg "Error en SearchURL_from_StateName > Module RMVoiceData", " Function_data: " & WStateName, err.Number, False
Close #16
SearchURL_from_StateName = "error"
End Function

Function ImportTextFile(WFileName As String) As Boolean

Dim DataId As String, Data As String

On Error GoTo err
Open WFileName For Input As #17
Do Until EOF(17)
Line Input #17, Data

    DataId = Left$(Data, 8)
    Data = Mid$(Data, 9, Len(Data))
    
    StateDatabase.State_URL = "http://www.weather.com/outlook/travel/businesstraveler/local/" & Trim(DataId) & "?from=search_city"
    StateDatabase.State_Desc = Trim(Data)
        
    FileState = SaveStateData(0, StateDatabase)
    
Loop
Close #17
ImportTextFile = True
Exit Function

err:
ImportTextFile = False
Close #17
DisplayMsg "Error en ImportTextFile > Module RMVoiceData - ", " Function_data: " & WFileName, err.Number, False
End Function

'cargamos los datos de la base de datos
Function GetStateData(WOptionalId As Integer) As GEN_State_Database

On Error GoTo err
Open App.Path & DatabaseFileName For Random As #18 Len = Len(StateDatabase)

If WOptionalId = 0 Or WOptionalId = -1 Then
    Close #18
    GetStateData.Id = -1
    Exit Function
Else
    LastReg = WOptionalId
End If

Get #18, LastReg, StateDatabase
GetStateData = StateDatabase
Close #18

Close #18
Exit Function

'/// if there is an error ------------------------------------------
err:
DisplayMsg "Error en GetStateData > Module RMVoiceData - ", " Function_data: " & WOptionalId, err.Number, False
Close #18
GetStateData.Id = -1
End Function

'guardamos los datos en la base de datos
Function SaveStateData(WOptionalId As Integer, Wdata As GEN_State_Database) As Boolean

'/// abrimos el archivo de productos
On Error GoTo err
Open App.Path & DatabaseFileName For Random As #15 Len = Len(StateDatabase)

If WOptionalId = 0 Or WOptionalId = -1 Then
    LastReg = GetStateLastReg
    LastReg = LastReg + 1
Else
    LastReg = WOptionalId
End If

'/// seteamos los datos del inventario a guardar
StateDatabase.Id = LastReg
StateDatabase.State_Desc = Trim(Wdata.State_Desc)
StateDatabase.State_URL = Trim(Wdata.State_URL)

'/// guardamos
Put #15, LastReg, StateDatabase
Close #15

SaveStateData = True
Exit Function

'/// if there is an error ------------------------------------------
err:
DisplayMsg "Error en SaveStateData > Module RMVoiceData - ", " Function_data: " & WOptionalId & " - " & Wdata.State_Desc & " - " & Wdata.State_URL, err.Number, False
Close #15
SaveStateData = False
End Function

'extraemos el ultimo registro de la base de datos
Function GetStateLastReg()

'/// abrimos el archivo
On Error GoTo err
Open App.Path & DatabaseFileName For Random As #24 Len = Len(StateDatabase)

'/// check for the last reg ID
LastReg = LOF(24) \ Len(StateDatabase)

GetStateLastReg = LastReg

'///end the function
Close #24
Exit Function

'/// if there is an error ------------------------------------------
err:
DisplayMsg "Error en GetStateLastReg > Module RMVoiceData - ", " Function_data: " & DatabaseFileName, err.Number, False
GetStateLastReg = -1
Close #24

End Function

'cargamos los datos de la configuracion
Function GetConfigData() As CFG_Data

On Error GoTo err
Open App.Path & ConfigFileName For Random As #12 Len = Len(ConfigData)

Get #12, 1, ConfigData
GetConfigData = ConfigData
Close #12

Close #12
Exit Function

'/// if there is an error ------------------------------------------
err:
DisplayMsg "Error en GetConfigData > Module RMVoiceData - ", " Function_data: " & ConfigFileName, err.Number, False
Close #12
GetConfigData.Id = -1
End Function

'guardamos los datos de la configuracion
Function SaveConfigData(Wdata As CFG_Data) As Boolean

'/// abrimos el archivo de productos
On Error GoTo err
Open App.Path & ConfigFileName For Random As #14 Len = Len(ConfigData)

'/// seteamos los datos del inventario a guardar
ConfigData.Id = 1
ConfigData.Lng_Id = Wdata.Lng_Id
ConfigData.State_Id = Trim(Wdata.State_Id)
ConfigData.Temp_Mode = Wdata.Temp_Mode
ConfigData.Voice_Id = Wdata.Voice_Id
ConfigData.Temp_RTime = Wdata.Temp_RTime
ConfigData.HVoicePath = Trim(Wdata.HVoicePath)
ConfigData.MVoicePath = Trim(Wdata.MVoicePath)
ConfigData.TmpVoicePath = Trim(Wdata.TmpVoicePath)
ConfigData.HumVoicePath = Trim(Wdata.HumVoicePath)

'/// guardamos
Put #14, 1, ConfigData
Close #14

SaveConfigData = True
Exit Function

'/// if there is an error ------------------------------------------
err:
DisplayMsg "Error en SaveConfigData > Module RMVoiceData - ", " Function_data: " & ConfigFileName, err.Number, False
Close #14
SaveConfigData = False
End Function

Function InitHora() As Long

ConfigData = GetConfigData

'procedemos a la carga del archivo de audio
If Stream01IsPlaying = BASSTRUE Then
    Result = Stream02Load(Trim(ConfigData.HVoicePath), "Stream")
    If Result = "999" Then Exit Function
    MainForm.LSTCommand.AddItem GetComLng_ByID(LNGDef, "150") 'reproducir hora
    MainForm.LblStrm.Caption = "2"
    Call SAYhora(2)
Else
    Result = Stream01Load(Trim(ConfigData.HVoicePath), "Stream")
    Debug.Print Trim(ConfigData.HVoicePath) & " Resultado: " & Result
    If Result = "999" Then Exit Function
    MainForm.LSTCommand.AddItem GetComLng_ByID(LNGDef, "150")
    MainForm.LblStrm.Caption = "1"
    Call SAYhora(1)
End If

End Function

Function InitMinutos() As Long

ConfigData = GetConfigData

If Stream01IsPlaying = BASSTRUE Then
    Result = Stream02Load(Trim(ConfigData.MVoicePath), "Stream")
    If Result = "999" Then Exit Function
    MainForm.LSTCommand.AddItem GetComLng_ByID(LNGDef, "151") 'reproducir minutos
    MainForm.LblStrm.Caption = "2"
    Call SAYminutos(2)
Else
    Result = Stream01Load(Trim(ConfigData.MVoicePath), "Stream")
    If Result = "999" Then Exit Function
    MainForm.LSTCommand.AddItem GetComLng_ByID(LNGDef, "151")
    MainForm.LblStrm.Caption = "1"
    Call SAYminutos(1)
End If

End Function

Function InitTemperatura() As Long

ConfigData = GetConfigData

'procedemos a la carga del archivo de audio de la temperatura
If Stream01IsPlaying = BASSTRUE Then
    Result = Stream02Load(Trim(ConfigData.TmpVoicePath), "Stream")
    If Result = "999" Then Exit Function
    MainForm.LSTCommand.AddItem GetComLng_ByID(LNGDef, "152") 'reproducir temperatura
    MainForm.LblStrm.Caption = "2"
    Call SAYtemperatura(2)
Else
    Result = Stream01Load(Trim(ConfigData.TmpVoicePath), "Stream")
    If Result = "999" Then Exit Function
    MainForm.LSTCommand.AddItem GetComLng_ByID(LNGDef, "152")
    MainForm.LblStrm.Caption = "1"
    Call SAYtemperatura(1)
End If

End Function

Function InitHumedad() As Long

ConfigData = GetConfigData

'procedemos a la carga del archivo de audio de la temperatura
If Stream01IsPlaying = BASSTRUE Then
    Result = Stream02Load(Trim(ConfigData.HumVoicePath), "Stream")
    If Result = "999" Then Exit Function
    MainForm.LSTCommand.AddItem GetComLng_ByID(LNGDef, "153") 'reproducir temperatura
    MainForm.LblStrm.Caption = "2"
    Call SAYhumedad(2)
Else
    Result = Stream01Load(Trim(ConfigData.HumVoicePath), "Stream")
    If Result = "999" Then Exit Function
    MainForm.LSTCommand.AddItem GetComLng_ByID(LNGDef, "153")
    MainForm.LblStrm.Caption = "1"
    Call SAYhumedad(1)
End If

End Function

Function SAYhora(StreamNum As Long) As Long

Dim Hora As String
Dim HR As Integer
Dim pos As Single, PosEnd As Long
Dim Npos As String, NPosEnd As String
Dim NewFileN As String

Hora = GetHH
HR = CInt(Hora)

If HR > 23 Then
    GoSub err
End If

ConfigData = GetConfigData
NewFileN = Trim(ConfigData.HVoicePath)
NewFileN = Left$(NewFileN, Len(NewFileN) - 4)
NewFileN = NewFileN & ".Dat"

On Error GoTo err
If StreamNum = 1 Then
    'play the file
    Stream01Play 0
    'Stream01SetVolumen 100
    Npos = GetRegion(NewFileN, CLng(HR), 1)
    If Npos = "999" Then GoSub err
    pos = CSng(Npos)
    pos = CSng(BASS_ChannelSeconds2Bytes(Strm1, CSng(Npos)))
    MainForm.LblStart.Caption = Str$(pos)
        NPosEnd = GetRegion(NewFileN, CLng(HR), 2)
        If NPosEnd = "999" Then GoSub err
    Stream01SetPosition pos, 2
    PosEnd = BASS_ChannelSeconds2Bytes(Strm1, CSng(NPosEnd))
    MainForm.LblEnd.Caption = Str$(PosEnd)
Else
    'play the file
    Stream02Play 0
        Npos = GetRegion(NewFileN, CLng(HR), 1)
        If Npos = "999" Then GoSub err
    pos = CSng(Npos)
    MainForm.LblStart.Caption = Str$(pos)
        NPosEnd = GetRegion(NewFileN, CLng(HR), 2)
        If NPosEnd = "999" Then GoSub err
    Stream02SetPosition pos, 2
    PosEnd = BASS_ChannelSeconds2Bytes(Strm2, CSng(NPosEnd))
    MainForm.LblEnd.Caption = Str$(PosEnd)
End If

MainForm.HTimer.Enabled = True
MainForm.HTimer.Interval = 1
Exit Function

err:
DisplayMsg "Error en SayHora > Module RMVoiceData - ", " Function_data: StrNum:" & StreamNum & " Filename:" & NewFileN & " Pos:" & Npos & " EndPos:" & NPosEnd, err.Number, False
End Function

Function SAYminutos(StreamNum As Long) As Long

Dim Minutos As String
Dim MN As Integer
Dim pos As Single, PosEnd As Long
Dim Npos As String, NPosEnd As String
Dim NewFileN As String

Minutos = GetMM
MN = CInt(Minutos)

If MN > 59 Then
    GoSub err
End If

ConfigData = GetConfigData
NewFileN = Trim(ConfigData.MVoicePath)
NewFileN = Left$(NewFileN, Len(NewFileN) - 4)
NewFileN = NewFileN & ".Dat"

On Error GoTo err
If StreamNum = 1 Then
    'play the file
    Stream01Play 0
        Npos = GetRegion(NewFileN, CLng(MN), 1)
    If Npos = "999" Then GoSub err
    pos = CSng(Npos)
    MainForm.LblStart.Caption = Str$(pos)
        NPosEnd = GetRegion(NewFileN, CLng(MN), 2)
    If NPosEnd = "999" Then GoSub err
    Stream01SetPosition pos, 1
    PosEnd = BASS_ChannelSeconds2Bytes(Strm1, CSng(NPosEnd))
    MainForm.LblEnd.Caption = Str$(PosEnd)
Else
    'play the file
    Stream02Play 0
        Npos = GetRegion(NewFileN, CLng(MN), 1)
    If Npos = "999" Then GoSub err
    pos = CSng(Npos)
    MainForm.LblStart.Caption = Str$(pos)
        NPosEnd = GetRegion(NewFileN, CLng(MN), 2)
    If NPosEnd = "999" Then GoSub err
    Stream02SetPosition pos, 1
    PosEnd = BASS_ChannelSeconds2Bytes(Strm2, CSng(NPosEnd))
    MainForm.LblEnd.Caption = Str$(PosEnd)
End If

MainForm.MTimer.Enabled = True
MainForm.MTimer.Interval = 1
Exit Function

err:
DisplayMsg "Error en SayMinutos > Module RMVoiceData - ", " Function_data: StrNum:" & StreamNum & " Filename:" & NewFileN & " Pos:" & Npos & " EndPos:" & NPosEnd, err.Number, False
End Function

Function SAYtemperatura(StreamNum As Long) As Long

Dim Temp As String
Dim TRp As Integer
Dim pos As Single, PosEnd As Long
Dim Npos As String, NPosEnd As String
Dim NewFileN As String

On Error GoTo err

Temp = GetTR
If Temp = "err" Then MsgBox "Error en temperatura!!": Exit Function
TRp = CInt(Temp)

If TRp > 49 Then
    GoSub err
End If

ConfigData = GetConfigData
NewFileN = Trim(ConfigData.TmpVoicePath)
NewFileN = Left$(NewFileN, Len(NewFileN) - 4)
NewFileN = NewFileN & ".Dat"

If StreamNum = 1 Then
    'play the file
    Stream01Play 0
        Npos = GetRegion(NewFileN, CLng(TRp), 1)
        If Npos = "999" Then GoSub err
    pos = CSng(Npos)
    MainForm.LblStart.Caption = Str$(pos)
        NPosEnd = GetRegion(NewFileN, CLng(TRp), 2)
        If NPosEnd = "999" Then GoSub err
    Stream01SetPosition pos, 1
    PosEnd = BASS_ChannelSeconds2Bytes(Strm1, CSng(NPosEnd))
    MainForm.LblEnd.Caption = Str$(PosEnd)
Else
    'play the file
    Stream02Play 0
        Npos = GetRegion(NewFileN, CLng(TRp), 1)
        If Npos = "999" Then GoSub err
    pos = CSng(Npos)
    MainForm.LblStart.Caption = Str$(pos)
        NPosEnd = GetRegion(NewFileN, CLng(TRp), 2)
        If NPosEnd = "999" Then GoSub err
    Stream02SetPosition pos, 1
    PosEnd = BASS_ChannelSeconds2Bytes(Strm2, CSng(NPosEnd))
    MainForm.LblEnd.Caption = Str$(PosEnd)
End If

MainForm.TTimer.Enabled = True
MainForm.TTimer.Interval = 1
Exit Function

err:
DisplayMsg "Error en SayTemperatura > Module RMVoiceData - ", " Function_data: StrNum:" & StreamNum & " Filename:" & NewFileN & " Pos:" & Npos & " EndPos:" & NPosEnd, err.Number, False
End Function

Function SAYhumedad(StreamNum As Long) As Long

Dim Hume As String
Dim Hup As Integer
Dim pos As Single, PosEnd As Long
Dim Npos As String, NPosEnd As String
Dim NewFileN As String

On Error GoTo err

Hume = GetHU
If Hume = "err" Then MsgBox "Error en humedad!!": Exit Function
Hup = CInt(Hume)

If Hup > 99 Then
    GoSub err
End If

ConfigData = GetConfigData
NewFileN = Trim(ConfigData.HumVoicePath)
NewFileN = Left$(NewFileN, Len(NewFileN) - 4)
NewFileN = NewFileN & ".Dat"

If StreamNum = 1 Then
    'play the file
    Stream01Play 0
        Npos = GetRegion(NewFileN, CLng(Hup), 1)
        If Npos = "999" Then GoSub err
    pos = CSng(Npos)
    MainForm.LblStart.Caption = Str$(pos)
        NPosEnd = GetRegion(NewFileN, CLng(Hup), 2)
        If NPosEnd = "999" Then GoSub err
    Stream01SetPosition pos, 1
    PosEnd = BASS_ChannelSeconds2Bytes(Strm1, CSng(NPosEnd))
    MainForm.LblEnd.Caption = Str$(PosEnd)
Else
    'play the file
    Stream02Play 0
        Npos = GetRegion(NewFileN, CLng(Hup), 1)
        If Npos = "999" Then GoSub err
    pos = CSng(Npos)
    MainForm.LblStart.Caption = Str$(pos)
        NPosEnd = GetRegion(NewFileN, CLng(Hup), 2)
        If NPosEnd = "999" Then GoSub err
    Stream02SetPosition pos, 1
    PosEnd = BASS_ChannelSeconds2Bytes(Strm2, CSng(NPosEnd))
    MainForm.LblEnd.Caption = Str$(PosEnd)
End If

MainForm.HuTimer.Enabled = True
MainForm.HuTimer.Interval = 1
Exit Function

err:
DisplayMsg "Error en SayHumedad > Module RMVoiceData - ", " Function_data: StrNum:" & StreamNum & " Filename:" & NewFileN & " Pos:" & Npos & " EndPos:" & NPosEnd, err.Number, False
End Function

Function GetTR() As String
'extraemos la temperatura
Dim lnga As Long
Dim pal As String
Dim TRa As String

pal = Trim(MainForm.Label1.Caption)
lnga = Len(pal)

Select Case lnga
    Case 2  '2°
        TRa = Left$(pal, 1)
        TRa = "0" & TRa
    
    Case 3  '12° or -2° or N/A
        TRa = Left$(pal, 1)
        Select Case TRa
            'Case "-"
            '    TRa = Mid$(pal, 2, 1)
            Case "N"
                TRa = "err"  'ERROR
            Case Else
                TRa = Left$(pal, 2)
        End Select
        
    Case 4  '-12° or 100°
        TRa = Left$(pal, 1)
        Select Case TRa
            'Case "-"
            '    TRa = Mid$(pal, 2, 2)
            Case Else
                TRa = Left$(pal, 3)
        End Select
        
    Case Else
        TRa = "err"  'ERROR
End Select

GetTR = TRa

End Function

Function GetHU() As String
'extraemos la humedad
Dim lng As Long
Dim pal, hu As String

pal = Trim(MainForm.Label2.Caption)
lng = Len(pal)

Select Case lng
    Case 2  '2%
        hu = Left$(pal, 1)
        hu = "0" & hu
    
    Case 3  '12% or N/A
        hu = Left$(pal, 1)
        Select Case hu
            Case "N"
                hu = "err"  'ERROR
            Case Else
                hu = Left$(pal, 2)
        End Select
        
    Case 4  '100%
        hu = Left$(pal, 1)
        Select Case hu
            Case Else
                hu = Left$(pal, 3)
        End Select
        
    Case Else
        hu = "err"  'ERROR
End Select

GetHU = hu

End Function

Function GetHH() As String
'extraemos la hora de la hora actual
Dim pal, hh As String

pal = time$     '04:29:49

hh = Left$(pal, 2)

GetHH = hh

End Function

Function GetMM() As String
'extraemos los minutos de la hora actual
Dim pal, mm As String

pal = time$     '04:29:49

mm = Mid$(pal, 4, 2)

GetMM = mm

End Function
