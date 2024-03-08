Attribute VB_Name = "NewRMBass"
'////////////////////////////////////////////////////////
'*
'*  ////// NEWRMBASS & FX module for RadioMaker //////
'*  ** this module depends on 100% of "modBass.bas" **
'*  ********* and is for Radiomaker 1+ only *********
'*
'*     Copyright (c) 1987-2024 Only development Inc.
'*     Christian A. Del Monte
'*     creadig@gmail.com / creadig@hotmail.com
'*
'///////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM SET POSITION FUNCTION - 03-03-2024
'Funcion para setear la posicion de reproduccion de un stream
'Wchan = canal a procesar
'WposOrWSeg =
'WType =
'//////////////////////////////////////////////////////////////////////////
Function GStreamSetPosition(ByVal WChan As Long, ByVal WPosOrWseg As Long, ByVal WType As Long) As Boolean

Private Rst As Long
Private RstS As Long

'wtype contants
'Const StrTime = 1
'Const StrByte = 2

'CHEQUEOS
Select Case WType
    Case StrTime
        If BASS_ChannelIsActive(WChan) = BASSTRUE Then
            RstS = BASS_ChannelSeconds2Bytes(WChan, WPosOrWseg)
            Rst = BASS_ChannelGetLength(WChan)
            If RstS > Rst Then  'compare is Ok
                DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
                GStreamSetPosition = BASSFALSE
            Else
                If BASS_ChannelSetPosition(WChan, RstS) = BASSFALSE Then
                    DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
                    GStreamSetPosition = BASSFALSE
                Else
                    GStreamSetPosition = BASSTRUE
                End If
            End If
        Else
            MsgBox "CANAL NO ACTIVO - CHEQUEAR: MOD NewRMBass - GStreamSetPosition"
            GStreamSetPosition = BASSFALSE
        End If
        
    Case StrByte
        If BASS_ChannelIsActive(WChan) = BASSTRUE Then
            Rst = BASS_ChannelGetLength(WChan)
            If WPosOrWseg > Rst Then  'compare is Ok
                DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
                GStreamSetPosition = BASSFALSE
            Else
                If BASS_ChannelSetPosition(WChan, WPosOrWseg) = BASSFALSE Then
                    DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
                    GStreamSetPosition = BASSFALSE
                Else
                    GStreamSetPosition = BASSTRUE
                End If
            End If
        Else
            GStreamSetPosition = BASSFALSE
        End If
End Select

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM LOAD FUNCTION - 03-03-2024
'Funcion para cargar un archivo stream
'WFileName =  path completo y nombre de archivo a cargar
'Wchan = canal anterior a liberar si es que hay cargado uno antes
'WChanType = si hay un anterior que tipo es = Music o Stream
'WEstNumber = numero de la estación que intenta cargar el archivo stream
'//////////////////////////////////////////////////////////////////////////
Function GStreamLoad(WFileName As String, ByVal WChan As Long, WChanType As String, WEstNumber As Long) As Boolean

'retorna BASSTRUE o BASSFALSE segun sea el caso por si hay algo mal

Dim StreamHandle1 As Long, StreamHandle2 As Long
Dim Mode1 As Long, Mode2 As Long, Mode3 As Long, Mode4 As Long, Mode5 As Long

'verificamos si hay un handle anterior y lo eliminamos
If WChanType = "Music" Then
    BASS_MusicFree WChan
Else
    GStreamClear WChan
End If

'*******************************************
'* see Bass Sample info flags for more info*
'*******************************************

'gets the config device data
ConfigData = OpenConfigFile

Select Case ConfigData.Aud_Type     '1=8bits    2=16bits
    Case 1
        Mode1 = BASS_SAMPLE_8BITS
    Case 2
        Mode1 = 0
    Case Else
        Mode1 = 0
End Select
Select Case ConfigData.Aud_Cual     '1=Mono     2=Stereo
    Case 1
        Mode2 = BASS_SAMPLE_MONO
    Case 2
        Mode2 = 0
    Case Else
        Mode2 = 0
End Select
Select Case ConfigData.Aud_Mode     '1=Normal   2=A3d   3=3d    4=Ogg
    Case 1
        Mode3 = 0
    Case 2
        Mode3 = 0
    Case 3
        Mode3 = BASS_SAMPLE_3D
    Case 4
        Mode3 = 0
    Case Else
        Mode3 = 0
End Select

Mode4 = BASS_MP3_SETPOS
Mode5 = BASS_SAMPLE_FX

'verificamos cual estación es la que quiere reproducir
Select Case WEstNumber
    Case 1 '**** Estación 1
        StreamHandle1 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, Mode1 Or Mode2 Or Mode3 Or Mode4 Or Mode5)
        If StreamHandle1 = 0 Then
            'DisplayMsg LoadResString(159)
            GStreamLoad = BASSFALSE
        Else
            Strm1 = StreamHandle1
            GStreamLoad = BASSTRUE
        End If
    Case 2 '**** Estación 2
        StreamHandle2 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, Mode1 Or Mode2 Or Mode3 Or Mode4 Or Mode5)
        If StreamHandle2 = 0 Then
            'DisplayMsg LoadResString(159)
            GStreamLoad = BASSFALSE
        Else
            Strm2 = StreamHandle2
            GStreamLoad = BASSTRUE
        End If
End Select

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM SET CHANNEL volume PAN FUNCTION - 03-03-2024
'Funcion para setear el paneo de un stream
'Wchan = canal a procesar = stream or music
'Wpan = valor del volumen a setear
'-100 = izq y +100 = der y 0=neutral o ambos canales
'//////////////////////////////////////////////////////////////////////////
Function GStreamSetPAN(ByVal WChan As Long, ByVal Wpan As Long) As Boolean

If BASS_ChannelIsActive(WChan) = BASSTRUE Then
    If Wpan < -100 Or Wpan > 100 Then
        DisplayMsg LoadResString(167)   'invalido
        GStreamSetPAN = BASSFALSE
    Else
        If BASS_ChannelSetAttributes(WChan, -1, -1, Wpan) = BASSTRUE Then
            GStreamSetPAN = BASSTRUE
        Else
            DisplayMsg LoadResString(168)   'no se puede
            GStreamSetPAN = BASSFALSE
        End If
    End If
Else
    GStreamSetPAN = BASSFALSE
End If

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM GET LENGHT FUNCTION - 03-03-2024
'Funcion para extraer la longitud de un stream
'Wchan = canal a extraer la longitud = stream or music
'wTypeDisplay = StrTime o StrByte
'//////////////////////////////////////////////////////////////////////////
Function GStreamGetLen(ByVal WChan As Long, ByVal WTypeDisplay As Long) As Long

'Const StrTime = 1
'Const StrByte = 2

Dim SByte As Long
Dim STime As Long

SByte = BASS_ChannelGetLength(WChan)    'get stream file lenght (Bytes)
STime = CLng(BASS_ChannelBytes2Seconds(WChan, SByte))

Select Case WTypeDisplay
    Case StrByte
        GStreamGetLen = SByte
        Exit Function
        
    Case StrTime
        GStreamGetLen = STime
        'BytesPS = Stream01GetBytesPS
        'Stream01GetLen = SLen / BytesPS
        Exit Function
End Select

GStreamGetLen = 0

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM Get Right volume level FUNCTION - 03-03-2024
'Funcion para extraer el volumen derecho de un stream
'Wchan = canal a examinar = stream or music
'//////////////////////////////////////////////////////////////////////////
Function GStreamGetRIGHTlevel(ByVal WChan As Long) As Long

Private Level As Long, RRRight As Long
Private B As Long

If BASS_ChannelIsActive(WChan) = BASSTRUE Then
    Level = BASS_ChannelGetLevel(WChan)  'stream file level meter
    B = 1
    If (B < 128) Then
        If HiWord(Level) >= B Then
            RRRight = HiWord(Level)
        Else
            RRRight = HiWord(Level)
            B = 2 * B - B / 2
        End If
    End If
    GStreamGetRIGHTlevel = RRRight
Else
    GStreamGetRIGHTlevel = 0
End If

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM Get Left volume level FUNCTION - 03-03-2024
'Funcion para extraer el volumen izquierdo de un stream
'Wchan = canal a examinar = stream or music
'//////////////////////////////////////////////////////////////////////////
Function GStreamGetLEFTlevel(ByVal WChan As Long) As Long

Private Level As Long, LLLeft As Long
Private A As Long

If BASS_ChannelIsActive(WChan) = BASSTRUE Then
    Level = BASS_ChannelGetLevel(WChan)  'stream file level meter
    A = 93
    If (A > 0) Then
        If LoWord(Level) >= A Then
            LLLeft = LoWord(Level)
        Else
            LLLeft = LoWord(Level)
            A = A * 2 / 3
        End If
    End If
    GStreamGetLEFTlevel = LLLeft
Else
    GStreamGetLEFTlevel = 0
End If

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM Get Bytes per Second FUNCTION - 03-03-2024
'Funcion para extraer los bytes por segundo de un stream
'Wchan = canal a examinar = stream or music
'//////////////////////////////////////////////////////////////////////////
Function GStreamGetBPS(ByVal WChan As Long) As Double

Private Flags As Long, bps As Long
Private Newf As BASS_CHANNELINFO

If BASS_ChannelGetAttributes(WChan, bps, 0, 0) = BASSTRUE Then
    Flags = BASS_ChannelGetInfo(WChan, Newf)
    If Not (Flags & BASS_SAMPLE_MONO) Then bps = bps * 2
    If Not (Flags & BASS_SAMPLE_8BITS) Then bps = bps * 2
    GStreamGetBPS = bps
Else
    GStreamGetBPS = 0
End If
 
End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM CLEAR FUNCTION - 03-03-2024
'Funcion para reiniciar y cerrar un stream
'Wchan = canal a reiniciar = stream or music
'//////////////////////////////////////////////////////////////////////////
Function GStreamClear(ByVal WChan As Long) As Boolean

'removes the last sync
'Result = StreamRmvSync(1)

If BASS_StreamFree(WChan) = BASSFALSE Then
    GStreamClear = BASSFALSE
Else
    GStreamClear = BASSTRUE
End If

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM GET POSITION FUNCTION - 03-03-2024
'Funcion para reinicio general de los streams
'Wchan = canal a extraer la posicion = stream or music
'wTypeDisplay = StrTime o StrByte
'//////////////////////////////////////////////////////////////////////////
Function GStreamGetPosition(ByVal WChan As Long, ByVal WTypeDisplay As Long) As Long

'Const StrTime = 1
'Const StrByte = 2

Dim PosByte As Long
Dim PosTime As Long

If BASS_ChannelIsActive(WChan) = BASSTRUE Then
    PosByte = BASS_ChannelGetPosition(WChan)  'get stream file position (Bytes)
    PosTime = CLng(BASS_ChannelBytes2Seconds(WChan, PosByte)) 'convert byte 2 sec.
    
    Select Case WTypeDisplay
        Case StrByte
            GStreamGetPosition = PosByte    'devolvemos la pos en bytes
            Exit Function
            
        Case StrTime
            GStreamGetPosition = PosTime    'devolvemos la pos en segundos
            'BytesPS = Stream01GetBytesPS
            'Stream01GetPosition = Position / BytesPS
            Exit Function
            
    End Select
Else
    GStreamGetPosition = 0
End If

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM RESTART FUNCTION - 03-03-2024
'Funcion para reinicio general de los streams
'Wchan = canal a reiniciar = stream or music
'//////////////////////////////////////////////////////////////////////////
Function GStreamRestart(ByVal WChan As Long) As Boolean

If BASS_ChannelSetPosition(WChan, 0) = BASSFALSE Then
    DisplayMsg LoadResString(160)   'no se puede
    GStreamRestart = BASSFALSE
    Exit Function
End If

GStreamRestart = BASSTRUE

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM VOLUMEN FUNCTION - 03-03-2024
'Funcion para el seteo del volumen general de los streams
'Wchan = canal a reproducir = stream or music
'Wvol = valor del volumen
'//////////////////////////////////////////////////////////////////////////
Function GStreamSetVolume(ByVal WChan As Long, ByVal WVol As Long) As Boolean

If BASS_ChannelIsActive(WChan) = BASSTRUE Then
    If WVol < 0 Or WVol > 100 Then
        DisplayMsg LoadResString(169)   'volumen invalido
        GStreamSetVolume = BASSFALSE
        Exit Function
    Else
        If BASS_ChannelSetAttributes(WChan, -1, WVol, -101) = BASSFALSE Then
            DisplayMsg LoadResString(170)   'no se puede
            GStreamSetVolume = BASSFALSE
            Exit Function
        End If
    End If
End If

GStreamSetVolume = BASSTRUE

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM PLAY FUNCTION - 03-03-2024
'Funcion para la reproducción general de los streams
'Wchan = canal a reproducir = stream or music
'Wflag = sample loop o another
'//////////////////////////////////////////////////////////////////////////
Function GStreamPlay(ByVal WChan As Long, ByVal WFlag As Long) As Boolean

Select Case WFlag
    Case BASS_SAMPLE_LOOP
         If BASS_ChannelPlay(WChan, BASSFALSE) = BASSFALSE Then
            'DisplayMsg LoadResString(183) & " " & LoadResString(172)
            GStreamPlay = BASSFALSE
            Exit Function
        End If
    Case Else
        If BASS_ChannelPlay(WChan, BASSFALSE) = BASSFALSE Then
            'DisplayMsg LoadResString(159)
            GStreamPlay = BASSFALSE
            Exit Function
        End If
End Select

GStreamPlay = BASSTRUE

End Function

'//////////////////////////////////////////////////////////////////////////
'GENERAL STREAM STOP FUNCTION - 03-03-2024
'Funcion para detener la reproducción general de los streams
'Wchan = canal a detener = stream or music
'//////////////////////////////////////////////////////////////////////////
Function GStreamStop(ByVal WChan As Long) As Boolean

'Stop the stream
If BASS_ChannelStop(WChan) = BASSFALSE Then
    'DisplayMsg LoadResString(182)
    GStreamStop = BASSFALSE
    Exit Function
End If

GStreamStop = BASSTRUE

End Function
