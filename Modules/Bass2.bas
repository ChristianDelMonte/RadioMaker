Attribute VB_Name = "RMBass"
'////////////////////////////////////////////////////////
'*
'*  //////// MULTIMEDIA & FX module for Vb.6+ ////////
'*  ** this module depends on 100% of "modBass.bas" **
'*  ********* and is for Radiomaker 1.0 only *********
'*
'*     Copyright (c) 1987-2002 Only development Inc.
'*     Christian A. Del Monte
'///////////////////////////////////////////////////////

Option Explicit

Public Const StrTime = 1    '(stream) result in time
Public Const StrByte = 2    '(stream) result in byte
Public Const MscRowCol = 1  '(music) result in row/col
Public Const MscByte = 2    '(music) result in byte

'/// Stream / Music file handle dims
'/// Estacion  01
Public Strm1 As Long
Public Msc1 As Long

'/// Estacion 02
Public Strm2 As Long
Public Msc2 As Long

'/// PH handle
Public StrmPH As Long
Public StrmLen As Long
Public Muslen As Long

'/// Stream / Music Sync Handles
Public StrmYNC1 As Long    'tanda 01
Public StrmYNC2 As Long    'tanda 02
Public PHYNC As Long       'ph handle

'/// MOD Visualization Types
Public Enum defScopeMode
    ScopeSideBySide
    ScopeDouble
    ScopeLeftOnly
    ScopeRightOnly
End Enum

'/////////////////////////////////////
' This is the number of bands
' = 128 / sSize
'------------------------------------
' For 64 band, use 2
' For 32 band, use 4
' For 21 band, use 6
' For 16 band, use 8
' For 12 band, use 10
'/////////////////////////////////////
'Public Const sSize = 6   ' Boosts the volume on the fft for better visuals
'/////////////////////////////////////

'/// Visualization Gain
Public Const VisGain = 5   '5

'/// fx parameters (number of FX to use)
Dim fx1(11) As Long        ' 9 EQ band + reverb (+ other for est01 (not implemented))
Dim fx2(11) As Long        ' 9 EQ band + reverb (+ other for est02 (not implemented))

Public Sub SendMiniFFT(ByVal NumStream As Long, TypeStream As String, sSize)
    
'this is the FFT data to send to RMPlayer.dll control
'to visualize the sound

Dim D() As Single
Dim TopLevel As Long
Static Peak(256) As Long
Dim A As Long
Dim sCount As Long
Dim f As String
Dim Z As Long
Dim sLeft As Integer
Dim RetVal As Long

ReDim D(512) As Single  '1024
    
If NumStream = 1 Then
    Select Case TypeStream
        Case "Stream"
            RetVal = BASS_ChannelGetData(Strm1, D(0), BASS_DATA_FFT1024)
        Case "Music"
            RetVal = BASS_ChannelGetData(Msc1, D(0), BASS_DATA_FFT1024)
        Case Else
            'xxx nothing
    End Select
Else
    If NumStream = 2 Then
        Select Case TypeStream
            Case "Stream"
                RetVal = BASS_ChannelGetData(Strm2, D(0), BASS_DATA_FFT1024)
            Case "Music"
                RetVal = BASS_ChannelGetData(Msc2, D(0), BASS_DATA_FFT1024)
            Case Else
                'xxx nothing
        End Select
    Else
        Exit Sub
    End If
End If

'sending the fft data to rmm-mini player control
TopMenu.RMPlugIn.DrawMiniFFT D, sSize

End Sub
Public Sub SendMiniScope(ByVal NumStream As Long, TypeStream As String)

'this is the SCOPE data to send to RMPlayer.dll control
'to visualize the sound

Dim LLft As Long
Dim RRgt As Long
    
If NumStream = 1 Then
    Select Case TypeStream
        Case "Stream"
            LLft = Stream01GetLEFTLevel
            RRgt = Stream01GetRIGHTLevel
        Case "Music"
            LLft = Music01GetLEFTLevel
            RRgt = Music01GetRIGHTLevel
        Case Else
            'xxx nothing
    End Select
Else
    If NumStream = 2 Then
        Select Case TypeStream
            Case "Stream"
                LLft = Stream02GetLEFTLevel
                RRgt = Stream02GetRIGHTLevel
            Case "Music"
                LLft = Music02GetLEFTLevel
                RRgt = Music02GetRIGHTLevel
            Case Else
                'xxx nothing
        End Select
    Else
        Exit Sub
    End If
End If

'sending the fft data to rmm-mini player control
TopMenu.RMPlugIn.SetBuffLevel LLft, RRgt

End Sub

Public Sub CheckForTimers(ByVal WState As Long)

'wastate:   1=activar   0=desactivar

'gets the config device data
ConfigData = OpenConfigFile

If Stream01IsPlaying = True Or Music01IsPlaying = True Then
    Select Case Est12Control.Origen1.Caption
        Case "E1"
            Select Case WState
                Case 0
                    TopMenu.ProcTimer.Interval = 0
                    TopMenu.ProcTimer.Enabled = False
                    Est01.TmrScopeLite.Interval = 0
                    Est01.TmrScopeLite.Enabled = False
                    
                Case 1
                    TopMenu.ProcTimer.Enabled = True
                    TopMenu.ProcTimer.Interval = 1
                    If ConfigData.Aud_Show_FTT = 1 Or ConfigData.Aud_Show_SCOPE = 1 Then
                        'activate the level meter
                        Est01.TmrScopeLite.Enabled = True
                        Est01.TmrScopeLite.Interval = 25
                    End If

            End Select
        Case "T1"
            Select Case WState
                Case 0
                    TopMenu.ProcTimer.Interval = 0
                    TopMenu.ProcTimer.Enabled = False
                
                Case 1
                    TopMenu.ProcTimer.Enabled = True
                    TopMenu.ProcTimer.Interval = 1
            
            End Select
    End Select
End If

If Stream02IsPlaying = True Or Music02IsPlaying = True Then
    Select Case Est12Control.Origen2.Caption
        Case "E1"
            Select Case WState
                Case 0
                    TopMenu.ProcTimer.Interval = 0
                    TopMenu.ProcTimer.Enabled = False
                    Est02.TmrScopeLite2.Interval = 0
                    Est02.TmrScopeLite2.Enabled = False
                    
                Case 1
                    TopMenu.ProcTimer.Enabled = True
                    TopMenu.ProcTimer.Interval = 1
                    If ConfigData.Aud_Show_FTT = 1 Or ConfigData.Aud_Show_SCOPE = 1 Then
                        'activate the level meter
                        Est02.TmrScopeLite2.Enabled = True
                        Est02.TmrScopeLite2.Interval = 25
                    End If
                
            End Select
        Case "T1"
            Select Case WState
                Case 0
                    TopMenu.ProcTimer.Interval = 0
                    TopMenu.ProcTimer.Enabled = False
                    
                Case 1
                    TopMenu.ProcTimer.Enabled = True
                    TopMenu.ProcTimer.Interval = 1
                
            End Select
    End Select
End If

End Sub

Function FileLenGetBytesPS() As Double

'Esta funcion es para uso interno de la libreria.
Dim Flags As Long, bps As Long

 On Error GoTo None
 Call BASS_ChannelGetAttributes(StrmLen, bps, 0, 0)
 
 'flags = BASS_ChannelGetFlags(StrmLen)
  Flags = BASS_ChannelGetLength(StrmLen)
  
 If Not (Flags & BASS_SAMPLE_MONO) Then bps = bps * 2
 If Not (Flags & BASS_SAMPLE_8BITS) Then bps = bps * 2
 
 FileLenGetBytesPS = bps

Exit Function
 
None:
 FileLenGetBytesPS = 0
End Function

Sub AutoUpVol()

Dim vel As Integer

'seteamos la velocidad del timer de acuerdo
'a lo seleccionado de entre 0% a 100% por el
'usuario
Select Case FrmTime.SldVel.Value
    Case Is = 100
        vel = 10
    Case Is > 90, Is < 100, Is = 90
        vel = 9
    Case Is > 80, Is < 90, Is = 80
        vel = 8
    Case Is > 70, Is < 80, Is = 70
        vel = 7
    Case Is > 60, Is < 70, Is = 60
        vel = 6
    Case Is > 50, Is < 60, Is = 50
        vel = 5
    Case Is > 40, Is < 50, Is = 40
        vel = 4
    Case Is > 30, Is < 40, Is = 30
        vel = 3
    Case Is > 20, Is < 30, Is = 20
        vel = 2
    Case Is > 10, Is < 20, Is = 10
        vel = 1
    Case Is < 10
        vel = 1
End Select

'lets check the devices
If Stream01IsPlaying = True Or Music01IsPlaying = True Then
    Select Case Est12Control.Origen1.Caption
        Case "E1"
            'subir volumen e1
            FrmTime.E1Vin.Enabled = True
            FrmTime.E1Vin.Interval = vel
        Case "T1"
            'subir volumen t1 1
            FrmTime.T1VIn.Enabled = True
            FrmTime.T1VIn.Interval = vel
    End Select
End If
If Stream02IsPlaying = True Or Music02IsPlaying = True Then
    Select Case Est12Control.Origen2.Caption
        Case "E2"
            'subir volumen e2
            FrmTime.E2VIn.Enabled = True
            FrmTime.E2VIn.Interval = vel
        Case "T2"
            'subir volumen t1 2
            FrmTime.T2VIn.Enabled = True
            FrmTime.T2VIn.Interval = vel
    End Select
End If

End Sub
Sub AutoDwVol()

Dim vel As Integer

'seteamos la velocidad del timer de acuerdo
'a lo seleccionado de entre 0% a 100% por el
'usuario
Select Case FrmTime.SldVel.Value
    Case Is = 100
        vel = 10
    Case Is > 90, Is < 100, Is = 90
        vel = 9
    Case Is > 80, Is < 90, Is = 80
        vel = 8
    Case Is > 70, Is < 80, Is = 70
        vel = 7
    Case Is > 60, Is < 70, Is = 60
        vel = 6
    Case Is > 50, Is < 60, Is = 50
        vel = 5
    Case Is > 40, Is < 50, Is = 40
        vel = 4
    Case Is > 30, Is < 40, Is = 30
        vel = 3
    Case Is > 20, Is < 30, Is = 20
        vel = 2
    Case Is > 10, Is < 20, Is = 10
        vel = 1
    Case Is < 10
        vel = 1
End Select

'lets check the devices
If Stream01IsPlaying = True Or Music01IsPlaying = True Then
    Select Case Est12Control.Origen1.Caption
        Case "E1"
            'bajar volumen e1
            FrmTime.E1VOut.Enabled = True
            FrmTime.E1VOut.Interval = vel
        Case "T1"
            'bajar volumen t1 1
            FrmTime.T1VOut.Enabled = True
            FrmTime.T1VOut.Interval = vel
    End Select
End If
If Stream02IsPlaying = True Or Music02IsPlaying = True Then
    Select Case Est12Control.Origen2.Caption
        Case "E2"
            'bajar volumen e2
            FrmTime.E2Vout.Enabled = True
            FrmTime.E2Vout.Interval = vel
        Case "T2"
            'bajar volumen t1 2
            FrmTime.T2VOut.Enabled = True
            FrmTime.T2VOut.Interval = vel
    End Select
End If

End Sub
Function StreamPHRmvSync() As String

Dim Rslt As Integer

Rslt = BASS_ChannelRemoveSync(StrmPH, PHYNC)
If Rslt = BASSFALSE Then
    StreamPHRmvSync = "NotOk"
Else
    StreamPHRmvSync = "Ok"
End If

End Function

Function StreamPHSetSync() As String

'PHYNC=PH handle sync
PHYNC = BASS_ChannelSetSync(StrmPH, BASS_SYNC_END, 0, AddressOf SYNCPROC_PH, 0)

StreamPHSetSync = "Ok"

End Function

Sub StreamPHPlay()

'Play stream
'If BASS_StreamPlay(StrmPH, BASSFALSE, 0) = BASSFALSE Then
'    DisplayMsg LoadResString(158)   'no se puede reproducir
'    Exit Sub
'End If

End Sub

Function StreamPHLoad(WFileName As String) As String

'retorna NotOk si hay algo mal
'retorna Stream (new handle) si fue satisfactorio

Dim PHHandle1 As Long

Call PHRmv
BASS_StreamFree StrmPH   'stream

PHHandle1 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, 0)
If PHHandle1 = 0 Then
    DisplayMsg LoadResString(159)   'no se puede cargar
    StreamPHLoad = "NotOk"
Else
    StrmPH = PHHandle1
    StreamPHLoad = "Stream"
End If

End Function

Function FileLenGetLen(WFileType As String) As Long

'Funcion para utilizar unicamente con RadioMaker XPlorer

Dim SLen As Long
Dim BytesPS As Double

On Error GoTo None
Select Case WFileType
    Case "Stream"
        SLen = BASS_ChannelGetLength(StrmLen)
        'SLen = BASS_StreamGetLength(StrmLen)  'stream file lenght (Bytes)
        BytesPS = FileLenGetBytesPS
        FileLenGetLen = SLen / BytesPS
        
    Case "Music"
        FileLenGetLen = 0
        
End Select
Exit Function

None:
FileLenGetLen = 0
End Function

Function FileLoadLen(WFileName As String, WFileType As String) As Long

'Funcion para utilizar unicamente con RadioMaker XPlorer
'retorna NotOk si hay algo mal
'retorna filelen si fue satisfactorio
Dim ResultFile As Boolean
Dim StreamHandle1 As Long, MusicHandle1 As Long
Dim LenFile As Long

'chequeamos si el archivo existe
ResultFile = FileExist(WFileName)
If ResultFile = False Then
    FileLoadLen = 999
    Exit Function
End If

'removemos los handles anteriores
BASS_MusicFree StrmLen     'music
BASS_StreamFree Muslen   'stream

'load the file
Select Case WFileType
    Case "Stream"
        StreamHandle1 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, 0)
        If StreamHandle1 = 0 Then
            DisplayMsg LoadResString(159)   'no se puede cargar
            FileLoadLen = "NotOk"
            Exit Function
        Else
            StrmLen = StreamHandle1
            'load file len info
            LenFile = FileLenGetLen("Stream")
            'return the result
            FileLoadLen = LenFile
        End If
    Case "Music"
        MusicHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, 0, 0)
        If MusicHandle1 = 0 Then
            DisplayMsg LoadResString(159)   'no se puede cargar
            FileLoadLen = "NotOk"
            Exit Function
        Else
            Muslen = MusicHandle1
            DisplayMsg LoadResString(1000)  'no implementado todavia...
            FileLoadLen = 0
        End If
End Select

'removemos los handles anteriores
BASS_MusicFree StrmLen     'music
BASS_StreamFree Muslen   'stream

End Function

Function StreamRmvSync(ByVal WStream As Long) As String

Dim Rslt As Integer

On Error GoTo None
Select Case WStream
    Case 1
        Rslt = BASS_ChannelRemoveSync(Strm1, StrmYNC1)
    Case 2
        Rslt = BASS_ChannelRemoveSync(Strm2, StrmYNC2)
End Select

If Rslt = BASSFALSE Then
    StreamRmvSync = "NotOk"
Else
    StreamRmvSync = "Ok"
End If
Exit Function

None:
StreamRmvSync = "NotOk"

End Function

Function StreamSetSyncPos(ByVal WStream As Long, ByVal WSegs As Long) As String

Dim CurrPos As Long
Dim ln As Long

Select Case WStream
    Case 1  '/////////////////////////////////////////////// stream01
        ln = Stream01GetLen(1)  'get stream1 lenght (seconds)
        CurrPos = ln - WSegs
        Tanda01.SyncLabel.Caption = CurrPos
        Tanda01.SyncStream.Caption = "Stream01"
        'enable the timer sync
        Tanda01.SyncTimer.Enabled = True
        Tanda01.SyncTimer.Interval = 1
        
    Case 2  '/////////////////////////////////////////////// stream02
        ln = Stream02GetLen(1)  'get stream2 lenght (seconds)
        CurrPos = ln - WSegs
        Tanda01.SyncLabel.Caption = CurrPos
        Tanda01.SyncStream.Caption = "Stream02"
        'enable the timer sync
        Tanda01.SyncTimer.Enabled = True
        Tanda01.SyncTimer.Interval = 1
        
End Select

StreamSetSyncPos = "Ok"

End Function

Public Sub UpdateFX02(ByVal B As Integer)

Dim v As Integer
  
v = Est02.fxsc(B).Value
  
Select Case B
    Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
        Dim P As BASS_FXPARAMEQ
        Call BASS_FXGetParameters(fx2(B), P)
        P.fGain = 10 - v
        Call BASS_FXSetParameters(fx2(B), P)
    Case 10
        Dim p1 As BASS_FXREVERB
        Call BASS_FXGetParameters(fx2(B), p1)
        p1.fReverbMix = -0.012 * v * v * v
        Call BASS_FXSetParameters(fx2(B), p1)
    Case 11
        'xxx add other effect
        'not implemented yet...
End Select
  
End Sub

Public Sub UpdateFX01(ByVal B As Integer)

Dim v As Integer

v = Est01.fxsc(B).Value
  
Select Case B
    Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
        Dim P As BASS_FXPARAMEQ
        Call BASS_FXGetParameters(fx1(B), P)
        P.fGain = 10 - v
        Call BASS_FXSetParameters(fx1(B), P)
    Case 10
        Dim p1 As BASS_FXREVERB
        Call BASS_FXGetParameters(fx1(B), p1)
        p1.fReverbMix = -0.012 * v * v * v
        Call BASS_FXSetParameters(fx1(B), p1)
    Case 11
        'xxx add other effect
        'not implemented yet...
End Select
  
End Sub

Sub InitEffect(ByVal channel As Long, ChanType As String)

Dim P As BASS_FXPARAMEQ
Dim count As Integer

On Error Resume Next
Select Case channel
    Case 1
        Select Case ChanType
            Case "Stream"   'stream channel
                fx1(0) = BASS_ChannelSetFX(Strm1, BASS_FX_PARAMEQ, 0) 'bass
                fx1(1) = BASS_ChannelSetFX(Strm1, BASS_FX_PARAMEQ, 0) 'mid
                fx1(2) = BASS_ChannelSetFX(Strm1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(3) = BASS_ChannelSetFX(Strm1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(4) = BASS_ChannelSetFX(Strm1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(5) = BASS_ChannelSetFX(Strm1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(6) = BASS_ChannelSetFX(Strm1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(7) = BASS_ChannelSetFX(Strm1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(8) = BASS_ChannelSetFX(Strm1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(9) = BASS_ChannelSetFX(Strm1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(10) = BASS_ChannelSetFX(Strm1, BASS_FX_REVERB, 0) 'reverb
                'fx1(11) = not implemented... yet...
            Case "Music"    'music channel
                fx1(0) = BASS_ChannelSetFX(Msc1, BASS_FX_PARAMEQ, 0) 'bass
                fx1(1) = BASS_ChannelSetFX(Msc1, BASS_FX_PARAMEQ, 0) 'mid
                fx1(2) = BASS_ChannelSetFX(Msc1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(3) = BASS_ChannelSetFX(Msc1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(4) = BASS_ChannelSetFX(Msc1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(5) = BASS_ChannelSetFX(Msc1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(6) = BASS_ChannelSetFX(Msc1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(7) = BASS_ChannelSetFX(Msc1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(8) = BASS_ChannelSetFX(Msc1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(9) = BASS_ChannelSetFX(Msc1, BASS_FX_PARAMEQ, 0) 'treble
                fx1(10) = BASS_ChannelSetFX(Msc1, BASS_FX_REVERB, 0) 'reverb
                'fx1(11) = not implemented... yet...
        End Select
        P.fGain = 0
        P.fBandwidth = 18
        P.fCenter = 60
        Call BASS_FXSetParameters(fx1(0), P)
        P.fCenter = 170                    ' bass (125hz)
        Call BASS_FXSetParameters(fx1(1), P)
        P.fCenter = 310
        Call BASS_FXSetParameters(fx1(2), P)
        P.fCenter = 600
        Call BASS_FXSetParameters(fx1(3), P)
        P.fCenter = 1000                   ' mid (1khz)
        Call BASS_FXSetParameters(fx1(4), P)
        P.fCenter = 3000
        Call BASS_FXSetParameters(fx1(5), P)
        P.fCenter = 6000                   ' treble (8khz)
        Call BASS_FXSetParameters(fx1(6), P)
        P.fCenter = 12000
        Call BASS_FXSetParameters(fx1(7), P)
        P.fCenter = 14000
        Call BASS_FXSetParameters(fx1(8), P)
        P.fCenter = 16000
        Call BASS_FXSetParameters(fx1(9), P)
        ' you can add more EQ bands with changing:
        ' p.fCenter = N [hz] N>=80 and N<=16000
        For count = 0 To 10
            Call UpdateFX01(count)  'update the fx
        Next count
    Case 2
        Select Case ChanType
            Case "Stream"
                fx2(0) = BASS_ChannelSetFX(Strm2, BASS_FX_PARAMEQ, 0) 'bass
                fx2(1) = BASS_ChannelSetFX(Strm2, BASS_FX_PARAMEQ, 0) 'mid
                fx2(2) = BASS_ChannelSetFX(Strm2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(3) = BASS_ChannelSetFX(Strm2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(4) = BASS_ChannelSetFX(Strm2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(5) = BASS_ChannelSetFX(Strm2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(6) = BASS_ChannelSetFX(Strm2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(7) = BASS_ChannelSetFX(Strm2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(8) = BASS_ChannelSetFX(Strm2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(9) = BASS_ChannelSetFX(Strm2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(10) = BASS_ChannelSetFX(Strm2, BASS_FX_REVERB, 0) 'reverb
            Case "Music"
                fx2(0) = BASS_ChannelSetFX(Msc2, BASS_FX_PARAMEQ, 0) 'bass
                fx2(1) = BASS_ChannelSetFX(Msc2, BASS_FX_PARAMEQ, 0) 'mid
                fx2(2) = BASS_ChannelSetFX(Msc2, BASS_FX_PARAMEQ, 0)  'treble
                fx2(3) = BASS_ChannelSetFX(Msc2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(4) = BASS_ChannelSetFX(Msc2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(5) = BASS_ChannelSetFX(Msc2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(6) = BASS_ChannelSetFX(Msc2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(7) = BASS_ChannelSetFX(Msc2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(8) = BASS_ChannelSetFX(Msc2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(9) = BASS_ChannelSetFX(Msc2, BASS_FX_PARAMEQ, 0) 'treble
                fx2(10) = BASS_ChannelSetFX(Msc2, BASS_FX_REVERB, 0) 'reverb
        End Select
        P.fGain = 0
        P.fBandwidth = 18
        P.fCenter = 60
        Call BASS_FXSetParameters(fx2(0), P)
        P.fCenter = 170                    ' bass (125hz)
        Call BASS_FXSetParameters(fx2(1), P)
        P.fCenter = 310
        Call BASS_FXSetParameters(fx2(2), P)
        P.fCenter = 600
        Call BASS_FXSetParameters(fx2(3), P)
        P.fCenter = 1000                   ' mid (1khz)
        Call BASS_FXSetParameters(fx2(4), P)
        P.fCenter = 3000
        Call BASS_FXSetParameters(fx2(5), P)
        P.fCenter = 6000                   ' treble (8khz)
        Call BASS_FXSetParameters(fx2(6), P)
        P.fCenter = 12000
        Call BASS_FXSetParameters(fx2(7), P)
        P.fCenter = 14000
        Call BASS_FXSetParameters(fx2(8), P)
        P.fCenter = 16000
        Call BASS_FXSetParameters(fx2(9), P)
        ' you can add more EQ bands with changing:
        ' p.fCenter = N [hz] N>=80 and N<=16000
        For count = 0 To 10
            Call UpdateFX02(count)
        Next count
    Case Else
        'xxx nothing
End Select

End Sub
Public Function RoundDown(IntDone As Long, IntMax As Long, MaxAmount As Long) As Long
    
    Dim D As Long
    
    On Error Resume Next
    
    D = Int(MaxAmount * IntDone / IntMax)
    
    RoundDown = CInt(D)
    
End Function

Public Function Percentage(IntDone As Long, IntMax As Long) As Long
    
    Dim D As Long
    
    On Error Resume Next
    
    D = Int(100 * IntDone / IntMax)
    
    Percentage = CInt(D)
    
End Function

Public Sub DrawScope(ByVal Color1 As Long, ByVal Color2 As Long, ByVal aLeft As Long, ByVal aTop As Long, ByVal MaxX As Long, ByVal MaxY As Long, ByVal Stream As Long, WTypeStream As String, ScopeMode As defScopeMode)
    
    Static PeakVal As Integer
    Dim RetVal As Long
    Dim S() As Integer
    Dim A As Long
    Dim T As Long
    Dim sBool As Boolean
    Dim sVal As Long
    Dim Fudge As Integer
    Dim Thickness As Integer
    
    Fudge = CInt(MaxX / 2)
    Thickness = 1
    
    If ScopeMode = ScopeDouble Then
        ReDim S(MaxX * 2) As Integer
        
        If Stream = 1 Then
            Select Case WTypeStream
                Case "Stream"
                    RetVal = BASS_ChannelGetData(Strm1, S(0), MaxX * 2)
                Case "Music"
                    RetVal = BASS_ChannelGetData(Msc1, S(0), MaxX * 2)
            End Select
            Est01.Picfft1.Cls
            ' ScopeDouble
            T = 0
            For A = 0 To MaxX * 2
                sBool = Not sBool
                sVal = S(A) Xor 0
                sVal = RoundDown(sVal + 16384, 32768, MaxY)
                If sBool = False Then
                    Est01.Picfft1.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color1
                Else
                    Est01.Picfft1.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color2
                End If
                T = T + 1
                If T >= MaxX Then Exit For
            Next
        Else
            If Stream = 2 Then
                Select Case WTypeStream
                    Case "Stream"
                        RetVal = BASS_ChannelGetData(Strm2, S(0), MaxX * 2)
                    Case "Music"
                        RetVal = BASS_ChannelGetData(Msc2, S(0), MaxX * 2)
                End Select
                Est02.Picfft2.Cls
                ' ScopeDouble
                T = 0
                For A = 0 To MaxX * 2
                    sBool = Not sBool
                    sVal = S(A) Xor 0
                    sVal = RoundDown(sVal + 16384, 32768, MaxY)
                    If sBool = False Then
                        Est02.Picfft2.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color1
                    Else
                        Est02.Picfft2.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color2
                    End If
                    T = T + 1
                    If T >= MaxX Then Exit For
                Next
            Else
                Exit Sub
            End If
        End If
    End If
    
    If ScopeMode = ScopeLeftOnly Then
        ReDim S(MaxX * 2) As Integer
        
        If Stream = 1 Then
            Select Case WTypeStream
                Case "Stream"
                    RetVal = BASS_ChannelGetData(Strm1, S(0), MaxX * 2)
                Case "Music"
                    RetVal = BASS_ChannelGetData(Msc1, S(0), MaxX * 2)
            End Select
            Est01.Picfft1.Cls
            ' ScopeDouble
            T = 0
            For A = 0 To MaxX * 2
                sBool = Not sBool
                sVal = S(A) Xor 0
                sVal = RoundDown(sVal + 16384, 32768, MaxY)
                If sBool = False Then
                    Est01.Picfft1.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color1
                Else
'                    est01.picfft1.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color2
                End If
                T = T + 1
                If T >= MaxX Then Exit For
            Next
        Else
            If Stream = 2 Then
                Select Case WTypeStream
                    Case "Stream"
                        RetVal = BASS_ChannelGetData(Strm2, S(0), MaxX * 2)
                    Case "Music"
                        RetVal = BASS_ChannelGetData(Msc2, S(0), MaxX * 2)
                End Select
                Est02.Picfft2.Cls
                ' ScopeDouble
                T = 0
                For A = 0 To MaxX * 2
                    sBool = Not sBool
                    sVal = S(A) Xor 0
                    sVal = RoundDown(sVal + 16384, 32768, MaxY)
                    If sBool = False Then
                        Est02.Picfft2.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color1
                    Else
'                       est02.picfft2.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color2
                    End If
                    T = T + 1
                    If T >= MaxX Then Exit For
                Next
            Else
                Exit Sub
            End If
        End If
    End If
    
    If ScopeMode = ScopeRightOnly Then
        ReDim S(MaxX * 2) As Integer
        
        If Stream = 1 Then
            Select Case WTypeStream
                Case "Stream"
                    RetVal = BASS_ChannelGetData(Strm1, S(0), MaxX * 2)
                Case "Music"
                    RetVal = BASS_ChannelGetData(Msc1, S(0), MaxX * 2)
            End Select
            Est01.Picfft1.Cls
            ' ScopeDouble
            T = 0
            For A = 0 To MaxX * 2
                sBool = Not sBool
                sVal = S(A) Xor 0
                sVal = RoundDown(sVal + 16384, 32768, MaxY)
                If sBool = False Then
'                    est01.picfft1.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color1
                Else
                    Est01.Picfft1.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color2
                End If
                T = T + 1
                If T >= MaxX Then Exit For
            Next
        Else
            If Stream = 2 Then
                Select Case WTypeStream
                    Case "Stream"
                        RetVal = BASS_ChannelGetData(Strm2, S(0), MaxX * 2)
                    Case "Music"
                        RetVal = BASS_ChannelGetData(Msc2, S(0), MaxX * 2)
                End Select
                Est02.Picfft2.Cls
                ' ScopeDouble
                T = 0
                For A = 0 To MaxX * 2
                    sBool = Not sBool
                    sVal = S(A) Xor 0
                    sVal = RoundDown(sVal + 16384, 32768, MaxY)
                    If sBool = False Then
'                       est02.picfft2.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color1
                    Else
                        Est02.Picfft2.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color2
                    End If
                    T = T + 1
                    If T >= MaxX Then Exit For
                Next
            Else
                Exit Sub
            End If
        End If
    End If
    
    If ScopeMode = ScopeSideBySide Then
        ReDim S(MaxX * 2) As Integer
        
        If Stream = 1 Then
            Select Case WTypeStream
                Case "Stream"
                    RetVal = BASS_ChannelGetData(Strm1, S(0), MaxX)
                Case "Music"
                    RetVal = BASS_ChannelGetData(Msc1, S(0), MaxX)
            End Select
            Est01.Picfft1.Cls
            ' ScopeDouble
            T = 0
            For A = 0 To MaxX / 2
                sBool = Not sBool
                sVal = S(A) Xor 0
                sVal = RoundDown(sVal + 16384, 32768, MaxY)
                If sBool = False Then
                    Est01.Picfft1.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color1
                Else
                    Est01.Picfft1.Line (aLeft + Fudge + T, aTop + sVal)-(aLeft + Fudge + T + Thickness, aTop + sVal), Color2
                End If
                T = T + 1
                If T >= MaxX Then Exit For
            Next
        Else
            If Stream = 2 Then
                Select Case WTypeStream
                    Case "Stream"
                        RetVal = BASS_ChannelGetData(Strm2, S(0), MaxX)
                    Case "Music"
                        RetVal = BASS_ChannelGetData(Msc2, S(0), MaxX)
                End Select
                Est02.Picfft2.Cls
                ' ScopeDouble
                T = 0
                For A = 0 To MaxX / 2
                    sBool = Not sBool
                    sVal = S(A) Xor 0
                    sVal = RoundDown(sVal + 16384, 32768, MaxY)
                    If sBool = False Then
                        Est02.Picfft2.Line (aLeft + T, aTop + sVal)-(aLeft + T + Thickness, aTop + sVal), Color1
                    Else
                        Est02.Picfft2.Line (aLeft + Fudge + T, aTop + sVal)-(aLeft + Fudge + T + Thickness, aTop + sVal), Color2
                    End If
                    T = T + 1
                    If T >= MaxX Then Exit For
                Next
            Else
                Exit Sub
            End If
        End If
    End If
End Sub

Public Sub DrawFFT(ByVal NumStream As Long, TypeStream As String, sSize)
    
Dim D() As Single
Dim TopLevel As Long
Static Peak(256) As Long
Dim A As Long
Dim sCount As Long
Dim f As String
Dim Z As Long
Dim sLeft As Integer
Dim RetVal As Long

ReDim D(512) As Single  '1024
    
If NumStream = 1 Then
    Select Case TypeStream
        Case "Stream"
            RetVal = BASS_ChannelGetData(Strm1, D(0), BASS_DATA_FFT1024)
        Case "Music"
            RetVal = BASS_ChannelGetData(Msc1, D(0), BASS_DATA_FFT1024)
        Case Else
            'xxx nothing
    End Select
    If RetVal = 0 Then Exit Sub
    sLeft = Est01.Picfft1.ScaleWidth - 256
    sLeft = sLeft / 2
    ' This should be even number between 2 and 10
    Est01.Picfft1.Cls
    For A = 0 To 256 Step sSize
        Z = (D(A) * 1000)
        If Z > 10 Then Z = Z + VisGain
        If Z > 40 Then Z = 40
        If Z > Peak(A) Then
            Peak(A) = Z
        Else
            If Peak(A) > 45 Then
                Peak(A) = Peak(A) - 2
            Else
                Peak(A) = Peak(A) - 1
            End If
        End If
        ' Draw blue
        Est01.Picfft1.Line (sLeft + A, 45)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF
        ' Draw blue/red
        TopLevel = 10
        If Z > TopLevel Then
            Est01.Picfft1.Line (sLeft + A, 45 - TopLevel)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF
        End If
        ' Draw red
        TopLevel = 20
        If Z > TopLevel Then
            Est01.Picfft1.Line (sLeft + A, 45 - TopLevel)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF              '&HFFFF&, BF
        End If
        ' Draw red/orange
        TopLevel = 30
        If Z > TopLevel Then
            Est01.Picfft1.Line (sLeft + A, 45 - TopLevel)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF            '&H69CAFE, BF
        End If
        ' Draw Orange
        TopLevel = 35
        If Z > TopLevel Then
            Est01.Picfft1.Line (sLeft + A, 45 - TopLevel)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF            '&H80FF&, BF   'top
        End If
        ' Draw orange/yellow
        TopLevel = 40
        If Z > TopLevel Then
            Est01.Picfft1.Line (sLeft + A, 45 - TopLevel)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF
        End If
        Est01.Picfft1.Line (sLeft + A, 45 - Peak(A))-(sLeft + A + (sSize / 2), 45 - Peak(A)), &H808000, BF
        sCount = sCount + 1
    Next
Else
    If NumStream = 2 Then
        Select Case TypeStream
            Case "Stream"
                RetVal = BASS_ChannelGetData(Strm2, D(0), BASS_DATA_FFT1024)
            Case "Music"
                RetVal = BASS_ChannelGetData(Msc2, D(0), BASS_DATA_FFT1024)
            Case Else
                'xxx nothing
        End Select
        If RetVal = 0 Then Exit Sub
        sLeft = Est02.Picfft2.ScaleWidth - 256
        sLeft = sLeft / 2
        ' This should be even number between 2 and 10
        Est02.Picfft2.Cls
        For A = 0 To 256 Step sSize
            Z = (D(A) * 1000)
            If Z > 10 Then Z = Z + VisGain
            If Z > 40 Then Z = 40
            If Z > Peak(A) Then
                Peak(A) = Z
            Else
                If Peak(A) > 45 Then
                    Peak(A) = Peak(A) - 2
                Else
                    Peak(A) = Peak(A) - 1
                End If
            End If
            ' Draw blue
            Est02.Picfft2.Line (sLeft + A, 45)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF
            ' Draw blue/red
            TopLevel = 10
            If Z > TopLevel Then
                Est02.Picfft2.Line (sLeft + A, 45 - TopLevel)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF
            End If
            ' Draw red
            TopLevel = 20
            If Z > TopLevel Then
                Est02.Picfft2.Line (sLeft + A, 45 - TopLevel)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF              '&HFFFF&, BF
            End If
            ' Draw red/orange
            TopLevel = 30
            If Z > TopLevel Then
                Est02.Picfft2.Line (sLeft + A, 45 - TopLevel)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF            '&H69CAFE, BF
            End If
            ' Draw Orange
            TopLevel = 35
            If Z > TopLevel Then
                Est02.Picfft2.Line (sLeft + A, 45 - TopLevel)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF            '&H80FF&, BF   'top
            End If
            ' Draw orange/yellow
            TopLevel = 40
            If Z > TopLevel Then
                Est02.Picfft2.Line (sLeft + A, 45 - TopLevel)-(sLeft + A + (sSize / 2), 45 - Z), &HC0C000, BF
            End If
            Est02.Picfft2.Line (sLeft + A, 45 - Peak(A))-(sLeft + A + (sSize / 2), 45 - Peak(A)), &H808000, BF
            sCount = sCount + 1
        Next
    Else
        Exit Sub
    End If
End If
    
    
End Sub


Public Function Stream02GetBytesPS() As Double

'Funcion solo para uso interno de la libreria
 Dim Flags As Long, bps As Long
 Dim Newf As BASS_CHANNELINFO
 
 On Error GoTo None

 Call BASS_ChannelGetAttributes(Strm2, bps, 0, 0)
 
 'flags = BASS_ChannelGetFlags(Strm2)
 Flags = BASS_ChannelGetInfo(Strm2, Newf)
 
 If Not (Flags & BASS_SAMPLE_MONO) Then bps = bps * 2
 If Not (Flags & BASS_SAMPLE_8BITS) Then bps = bps * 2
 
 Stream02GetBytesPS = bps
Exit Function
 
None:
 Stream02GetBytesPS = 0
End Function

Function Stream01GetBytesPS() As Double

'Esta funcion es para uso interno de la libreria.
Dim Flags As Long, bps As Long
 Dim Newf As BASS_CHANNELINFO

 On Error GoTo None
 Call BASS_ChannelGetAttributes(Strm1, bps, 0, 0)
 
 'flags = BASS_ChannelGetFlags(Strm1)
 Flags = BASS_ChannelGetInfo(Strm1, Newf)
 
 If Not (Flags & BASS_SAMPLE_MONO) Then bps = bps * 2
 If Not (Flags & BASS_SAMPLE_8BITS) Then bps = bps * 2
 
 Stream01GetBytesPS = bps

Exit Function
 
None:
 Stream01GetBytesPS = 0
End Function

Function Music02IsPlaying() As Boolean

If BASS_ChannelIsActive(Msc2) = BASSTRUE Then
    Music02IsPlaying = True
Else
    Music02IsPlaying = False
End If

End Function

Function Music01IsPlaying() As Boolean

If BASS_ChannelIsActive(Msc1) = BASSTRUE Then
    Music01IsPlaying = True
Else
    Music01IsPlaying = False
End If

End Function

Function Stream02IsPlaying() As Boolean

If BASS_ChannelIsActive(Strm2) = BASSTRUE Then
    Stream02IsPlaying = True
Else
    Stream02IsPlaying = False
End If

End Function

Function Stream01IsPlaying() As Boolean

If BASS_ChannelIsActive(Strm1) = BASSTRUE Then
    Stream01IsPlaying = True
Else
    Stream01IsPlaying = False
End If

End Function

Sub Music01SetPosition(ByVal WOrder As Long, ByVal WRow As Long)

'Dim Rst As Long, LOrder As Long, LRow As Long

'If Music01IsPlaying = True Then
'    Rst = BASS_MusicGetLength(Msc1, BASSFALSE)
'    LOrder = GetLoWord(Rst)
'    LRow = GetHiWord(Rst)
'    If WOrder > LOrder Then
'        DisplayMsg LoadResString(160)   'no se puede...
'    Else
'        If BASS_ChannelSetPosition(Msc1, MakeLong(WOrder, WRow)) = BASSFALSE Then
'            DisplayMsg LoadResString(160)   '...posicion incorrecta
'        End If
'    End If
'End If

End Sub

Sub Music02SetPosition(ByVal WOrder As Long, ByVal WRow As Long)

'Dim Rst As Long, LOrder As Long, LRow As Long

'CHEQUEOS
'If Music02IsPlaying = True Then
'    Rst = BASS_MusicGetLength(Msc2, BASSFALSE)
'    LOrder = GetLoWord(Rst)
'    LRow = GetHiWord(Rst)
'    If WOrder > LOrder Then
'        DisplayMsg LoadResString(160)   'no se puede...
'    Else
'        If BASS_ChannelSetPosition(Msc2, MakeLong(WOrder, WRow)) = BASSFALSE Then
'            DisplayMsg LoadResString(160)   '...posicion incorrecta
'        End If
'    End If
'End If

End Sub

Sub Stream02SetPosition(ByVal WPosOrWseg As Long, ByVal WType As Long)

Dim Rst As Long
Dim RstS As Long

'wtype contants
'Const StrTime = 1
'Const StrByte = 2

'CHEQUEOS
Select Case WType
    Case StrTime
        If Stream02IsPlaying = True Then
            RstS = BASS_ChannelSeconds2Bytes(Strm2, WPosOrWseg)
            'RstS = (WPosOrWseg * 60000) * 3   'convert the seg into byte
            'Rst = BASS_StreamGetLength(Strm2)   'get the real lenght (byte)
            Rst = BASS_ChannelGetLength(Strm2)
            If RstS > Rst Then  'compare is Ok
                DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
            Else
                If BASS_ChannelSetPosition(Strm2, RstS) = BASSFALSE Then
                    DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
                End If
            End If
        End If

    Case StrByte
        If Stream02IsPlaying = True Then
            'Rst = BASS_StreamGetLength(Strm2)   'get the real lenght (byte)
            Rst = BASS_ChannelGetLength(Strm2)
            If WPosOrWseg > Rst Then  'compare is Ok
                'DisplayMsg "The play position can´t exceed the file position."
                DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
            Else
                If BASS_ChannelSetPosition(Strm2, WPosOrWseg) = BASSFALSE Then
                    DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
                End If
            End If
        End If

End Select

End Sub

Sub Stream01SetPosition(ByVal WPosOrWseg As Long, ByVal WType As Long)

Dim Rst As Long
Dim RstS As Long

'wtype contants
'Const StrTime = 1
'Const StrByte = 2

'CHEQUEOS
Select Case WType
    Case StrTime
        If Stream01IsPlaying = True Then
            RstS = BASS_ChannelSeconds2Bytes(Strm1, WPosOrWseg)
            'RstS = (WPosOrWseg * 60000) * 3 'convert the seg into byte
            'Rst = BASS_StreamGetLength(Strm1)   'get the real lenght (byte)
            Rst = BASS_ChannelGetLength(Strm1)
            If RstS > Rst Then  'compare is Ok
                DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
            Else
                If BASS_ChannelSetPosition(Strm1, RstS) = BASSFALSE Then
                    DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
                End If
            End If
        End If
        
    Case StrByte
        If Stream01IsPlaying = True Then
            'Rst = BASS_StreamGetLength(Strm1)   'get the real lenght (byte)
            Rst = BASS_ChannelGetLength(Strm1)
            If WPosOrWseg > Rst Then  'compare is Ok
                DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
            Else
                If BASS_ChannelSetPosition(Strm1, WPosOrWseg) = BASSFALSE Then
                    DisplayMsg LoadResString(160)   '...posicion espec. incorrecta
                End If
            End If
        End If
        
End Select

End Sub

Function Stream02GetLen(ByVal WTypeDisplay As Long) As Long

'Const StrTime = 1
'Const StrByte = 2

Dim SByte As Long
Dim STime As Long

'SByte = BASS_StreamGetLength(Strm2)  'get stream file lenght (Bytes)
SByte = BASS_ChannelGetLength(Strm2)
STime = CLng(BASS_ChannelBytes2Seconds(Strm2, SByte))

Select Case WTypeDisplay
    Case StrByte
        Stream02GetLen = SByte
    
    Case StrTime
        Stream02GetLen = STime
        'BytesPS = Stream02GetBytesPS
        'Stream02GetLen = SLen / BytesPS

End Select

End Function

Function Stream01GetLen(ByVal WTypeDisplay As Long) As Long

'Const StrTime = 1
'Const StrByte = 2

Dim SByte As Long
Dim STime As Long

'SByte = BASS_StreamGetLength(Strm1)  'get stream file lenght (Bytes)
SByte = BASS_ChannelGetLength(Strm1)
STime = CLng(BASS_ChannelBytes2Seconds(Strm1, SByte))

Select Case WTypeDisplay
    Case StrByte
        Stream01GetLen = SByte
    
    Case StrTime
        Stream01GetLen = STime
        'BytesPS = Stream01GetBytesPS
        'Stream01GetLen = SLen / BytesPS

End Select

End Function

Function Music02GetLen(ByVal WTypeDisplay As Long) As String

'Const MscRowCol = 1
'Const MscByte = 2

'Dim MLen As Long, LOrder As Long, LRow As Long

'Select Case WTypeDisplay
'    Case MscByte
'        MLen = BASS_MusicGetLength(Msc2, BASSTRUE) 'music length (in bytes)
'        Music02GetLen = Trim(Str$(MLen))
'    Case MscRowCol
'        MLen = BASS_MusicGetLength(Msc2, BASSFALSE) 'music length (row/col)
'        LOrder = GetLoWord(MLen)
'        LRow = GetHiWord(MLen)
'        Music02GetLen = Trim(Str$(LOrder)) & "," & Trim(Str$(LRow))
'    Case Else
'        Music02GetLen = "0"
'End Select

End Function

Function Music01GetLen(ByVal WTypeDisplay As Long) As String

'Const MscRowCol = 1
'Const MscByte = 2

'Dim MLen As Long, LOrder As Long, LRow As Long

'Select Case WTypeDisplay
'    Case MscByte
'        MLen = BASS_MusicGetLength(Msc1, BASSTRUE) 'music length (in bytes)
'        Music01GetLen = Trim(Str$(MLen))
    
'    Case MscRowCol
'        MLen = BASS_MusicGetLength(Msc1, BASSFALSE) 'music length (Order/Row)
'        LOrder = GetLoWord(MLen)
'        LRow = GetHiWord(MLen)
'        Music01GetLen = Trim(Str$(LOrder)) & "," & Trim(Str$(LRow))

'    Case Else
'        Music01GetLen = "0"
'End Select

Exit Function

None:
Music01GetLen = "0"
End Function

Sub CloseDevice(WLHandle1 As String, WLHandle2 As String)

' Stop digital output
BASS_Stop

' Free the first handle
Select Case WLHandle1
    Case "Stream"
        BASS_StreamFree Strm1   'stream
    Case "Music"
        BASS_MusicFree Msc1     'music
    Case Else
        'xxx    NOTHING
End Select

' Free the second handle
Select Case WLHandle2
    Case "Stream"
        BASS_StreamFree Strm2   'stream
    Case "Music"
        BASS_MusicFree Msc2     'music
    Case Else
        'xxx    NOTHING
End Select

' Close digital sound system
BASS_Free

End Sub

Function InitDevice(ByVal Whwnd As Long) As String

Dim ParmResult As String
Dim bi As BASS_INFO
Dim Mode1 As Long, Mode2 As Long, Mode3 As Long
Dim Msg As String, Msg0 As String, Msg3 As String, Msg4 As String
Dim Style, Title, Response

' Check that BASS 1.4 was loaded

' check the correct BASS was loaded
If (HiWord(BASS_GetVersion) <> BASSVERSION) Then
    DisplayMsg LoadResString(161)   'no se puede iniciar bass.dll
    'Call MsgBox("An incorrect version of BASS.DLL was loaded", vbCritical)
    InitDevice = "NotOk"
    Exit Function
End If

'If BASS_GetStringVersion <> "1.8" Then
'    DisplayMsg LoadResString(161)   'no se puede iniciar bass.dll
'    InitDevice = "NotOk"
'    Exit Function
'End If

'**********************
'* Device setup flags *
'**********************
'Global Const BASS_DEVICE_8BITS = 1     'use 8 bit resolution, else 16 bit
'Global Const BASS_DEVICE_MONO = 2      'use mono, else stereo
'Global Const BASS_DEVICE_3D = 4        'enable 3D functionality
' If the BASS_DEVICE_3D flag is not specified when initilizing BASS,
' then the 3D flags (BASS_SAMPLE_3D and BASS_MUSIC_3D) are ignored when
' loading/creating a sample/stream/music.
'Global Const BASS_DEVICE_A3D = 8       'enable A3D functionality
'Global Const BASS_DEVICE_NOSYNC = 16   'disable synchronizers
'Global Const BASS_DEVICE_LEAVEVOL = 32 'leave volume as it is
'Global Const BASS_DEVICE_OGG = 64      'enable OGG support (requires OGG.DLL & VORBIS.DLL)
'Global Const BASS_DEVICE_NOTHREAD = 128 'update buffers manually (using BASS_Update)

InitComp:
'get the config device data
ConfigData = OpenConfigFile

Select Case ConfigData.Aud_Type     '1=8bits    2=16bits
    Case 1
        Mode1 = BASS_DEVICE_8BITS
    Case 2
        Mode1 = 0
    Case Else
        Mode1 = 0
End Select
Select Case ConfigData.Aud_Cual     '1=Mono     2=Stereo
    Case 1
        Mode2 = BASS_DEVICE_MONO
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
        Mode3 = BASS_DEVICE_3D
    Case 4
        Mode3 = 0
    Case Else
        Mode3 = 0
End Select

' Initialize digital sound - default device, 44100hz, stereo, 16 bits
If BASS_Init(-1, 44100, Mode1 Or Mode2 Or Mode3, Whwnd, 0) = BASSFALSE Then
    DisplayMsg LoadResString(134)
    InitDevice = "NotOk"
    Exit Function
End If

' Start digital output
If BASS_Start = BASSFALSE Then
    DisplayMsg LoadResString(134)
    InitDevice = "NotOk"
    Exit Function
End If

' check for DX8 drivers.
'bi.Size = LenB(bi)      'LenB(..) returns a byte data
Call BASS_GetInfo(bi)
If (bi.dsver < 8) Then
    Msg0 = LoadResString(166)
    Msg3 = " "
    Msg4 = LoadResString(132)
    Msg = Msg0 & " " & Msg3 & " " & Msg4
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "Rm100 - DirectX 8 no instalado!!!!"
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then
        Est12Control.LblFX.Caption = "NoFX"
    Else
        BASS_Free
        End
    End If
End If

InitDevice = "Ok"
End Function

Sub Stream02SetPan(ByVal Wpan As Long)

If Stream02IsPlaying = True Then
    If Wpan < -100 Or Wpan > 100 Then
        DisplayMsg LoadResString(167)   'invalido
    Else
        If BASS_ChannelSetAttributes(Strm2, -1, -1, Wpan) = BASSFALSE Then
            DisplayMsg LoadResString(168)   'no se puede
        End If
    End If
End If

End Sub

Sub Stream01SetPan(ByVal Wpan As Long)

If Stream01IsPlaying = True Then
    If Wpan < -100 Or Wpan > 100 Then
        DisplayMsg LoadResString(167)   'invalido
    Else
        If BASS_ChannelSetAttributes(Strm1, -1, -1, Wpan) = BASSFALSE Then
            DisplayMsg LoadResString(168)   'no se puede
        End If
    End If
End If

End Sub

Sub Music02SetPan(ByVal Wpan As Long)

If Music02IsPlaying = True Then
    If Wpan < -100 Or Wpan > 100 Then
        DisplayMsg LoadResString(167)   'invalido
    Else
        If BASS_ChannelSetAttributes(Msc2, -1, -1, Wpan) = BASSFALSE Then
            DisplayMsg LoadResString(168)   'no se puede
        End If
    End If
End If

End Sub

Sub Music01SetPan(ByVal Wpan As Long)

If Music01IsPlaying = True Then
    If Wpan < -100 Or Wpan > 100 Then
        DisplayMsg LoadResString(167)   'invalido
    Else
        If BASS_ChannelSetAttributes(Msc1, -1, -1, Wpan) = BASSFALSE Then
            DisplayMsg LoadResString(168)   'no se puede
        End If
    End If
End If

End Sub

Sub Stream02SetVolume(ByVal WVol As Long)

If Stream02IsPlaying = True Then
    If WVol < 0 Or WVol > 100 Then
        DisplayMsg LoadResString(169)   'invalido
    Else
        If BASS_ChannelSetAttributes(Strm2, -1, WVol, -101) = BASSFALSE Then
            DisplayMsg LoadResString(170)   'no se puede
        End If
    End If
End If

End Sub

Sub Stream01SetVolume(ByVal WVol As Long)

If Stream01IsPlaying = True Then
    If WVol < 0 Or WVol > 100 Then
        DisplayMsg LoadResString(169)   'invalido
    Else
        If BASS_ChannelSetAttributes(Strm1, -1, WVol, -101) = BASSFALSE Then
            DisplayMsg LoadResString(170)   'no se puede
        End If
    End If
End If

End Sub

Sub Music02SetVolume(ByVal WVol As Long)

If Music02IsPlaying = True Then
    If WVol < 0 Or WVol > 100 Then
        DisplayMsg LoadResString(169)   'invalido
    Else
        If BASS_ChannelSetAttributes(Msc2, -1, WVol, -101) = BASSFALSE Then
            DisplayMsg LoadResString(170)   'no se puede
        End If
    End If
End If

End Sub

Sub Music01SetVolume(ByVal WVol As Long)

If Music02IsPlaying = True Then
    If WVol < 0 Or WVol > 100 Then
        DisplayMsg LoadResString(169)   'invalido
    Else
        If BASS_ChannelSetAttributes(Msc1, -1, WVol, -101) = BASSFALSE Then
            DisplayMsg LoadResString(170)   'no se puede
        End If
    End If
End If

End Sub

Sub DisplayMsg(Message As String)

'Display error dialogues
Dim ErrorNum As Long

ErrorNum = BASS_ErrorGetCode
MsgBox Message & vbCrLf & vbCrLf & " Error Codigo: " & ErrorNum & vbCrLf & BASS_ErrorGetCode, vbCritical, "RMBass Error"

End Sub

Function Music02GetRIGHTLevel() As Long

'Dim Level As Long, RRRight As Long
'Dim b As Long

'If Music02IsPlaying = True Then
'    Level = BASS_ChannelGetLevel(Msc2)  'music file level meter
'    b = 1
'    If (b < 128) Then
'        If GetHiWord(Level) >= b Then
'            RRRight = GetHiWord(Level)
'        Else
'            RRRight = GetHiWord(Level)
'            b = 2 * b - b / 2
'        End If
'    End If
'    Music02GetRIGHTLevel = RRRight
'Else
'    Music02GetRIGHTLevel = 0
'End If

End Function

Function Music01GetRIGHTLevel() As Long

'Dim Level As Long, RRRight As Long
'Dim b As Long

'If Music01IsPlaying = True Then
'    Level = BASS_ChannelGetLevel(Msc1)  'music file level meter
'    b = 1
'    If (b < 128) Then
'        If GetHiWord(Level) >= b Then
'            RRRight = GetHiWord(Level)
'        Else
'            RRRight = GetHiWord(Level)
'            b = 2 * b - b / 2
'        End If
'    End If
'    Music01GetRIGHTLevel = RRRight
'Else
'    Music01GetRIGHTLevel = 0
'End If

End Function

Function Stream02GetRIGHTLevel() As Long

Dim Level As Long, RRRight As Long
Dim B As Long

If Stream02IsPlaying = True Then
    Level = BASS_ChannelGetLevel(Strm2)  'stream file level meter
    B = 1
    If (B < 128) Then
        If HiWord(Level) >= B Then
            RRRight = HiWord(Level)
        Else
            RRRight = HiWord(Level)
            B = 2 * B - B / 2
        End If
    End If
    Stream02GetRIGHTLevel = RRRight
Else
    Stream02GetRIGHTLevel = 0
End If

End Function

Function Stream01GetRIGHTLevel() As Long

Dim Level As Long, RRRight As Long
Dim B As Long

If Stream01IsPlaying = True Then
    Level = BASS_ChannelGetLevel(Strm1)  'stream file level meter
    B = 1
    If (B < 128) Then
        If HiWord(Level) >= B Then
            RRRight = HiWord(Level)
        Else
            RRRight = HiWord(Level)
            B = 2 * B - B / 2
        End If
    End If
    Stream01GetRIGHTLevel = RRRight
Else
    Stream01GetRIGHTLevel = 0
End If

End Function

Function Music02GetLEFTLevel() As Long

'Dim Level As Long, LLLeft As Long
'Dim a As Long

'If Music02IsPlaying = True Then
'    Level = BASS_ChannelGetLevel(Msc2)  'Music file level meter
'    a = 93
'    If (a > 0) Then
'        If GetLoWord(Level) >= a Then
'            LLLeft = GetLoWord(Level)
'        Else
'            LLLeft = GetLoWord(Level)
'            a = a * 2 / 3
'        End If
'    End If
'    Music02GetLEFTLevel = LLLeft
'Else
'    Music02GetLEFTLevel = 0
'End If

End Function

Function Music01GetLEFTLevel() As Long

'Dim Level As Long, LLLeft As Long
'Dim a As Long

'If Music01IsPlaying = True Then
'    Level = BASS_ChannelGetLevel(Msc1)  'music file level meter
'    a = 93
'    If (a > 0) Then
'        If GetLoWord(Level) >= a Then
'            LLLeft = GetLoWord(Level)
'        Else
'            LLLeft = GetLoWord(Level)
'            a = a * 2 / 3
'        End If
'    End If
'    Music01GetLEFTLevel = LLLeft
'Else
'    Music01GetLEFTLevel = 0
'End If

End Function

Function Stream02GetLEFTLevel() As Long

Dim Level As Long, LLLeft As Long
Dim A As Long

If Stream02IsPlaying = True Then
    Level = BASS_ChannelGetLevel(Strm2)  'stream file level meter
        A = 93
    If (A > 0) Then
        If LoWord(Level) >= A Then
            LLLeft = LoWord(Level)
        Else
            LLLeft = LoWord(Level)
            A = A * 2 / 3
        End If
    End If
    Stream02GetLEFTLevel = LLLeft
Else
    Stream02GetLEFTLevel = 0
End If

End Function

Function Music02GetPosition(ByVal WTypeDisplay As Long) As String

'Const MscRowCol = 1
'Const MscByte = 2

'Dim Position As Long, LOrder As Long, LRow As Long

'If Music02IsPlaying = True Then
'    Position = BASS_ChannelGetPosition(Msc2)  'music file position (Order/Row)
'    Select Case WTypeDisplay
'        Case MscByte
'            Music02GetPosition = Trim(Str$(Position))
'        Case MscRowCol
'            LOrder = GetLoWord(Position)
'            LRow = GetHiWord(Position)
'            Music02GetPosition = Trim(Str$(LOrder)) & "," & Trim(Str$(LRow))
'        Case Else
'            Music02GetPosition = "0"
'    End Select
'Else
'    Music02GetPosition = "0"
'End If

End Function

Function Music01GetPosition(ByVal WTypeDisplay As Long) As String

'Const MscRowCol = 1
'Const MscByte = 2

'Dim Position As Long, LOrder As Long, LRow As Long

'If Music01IsPlaying = True Then
'    Position = BASS_ChannelGetPosition(Msc1)  'music file position (Order/Row)
'    Select Case WTypeDisplay
'        Case MscByte
'            Music01GetPosition = Trim(Str$(Position))
'        Case MscRowCol
'            LOrder = GetLoWord(Position)
'            LRow = GetHiWord(Position)
'            Music01GetPosition = Trim(Str$(LOrder)) & "," & Trim(Str$(LRow))
'        Case Else
'            Music01GetPosition = "0"
'    End Select
'Else
'    Music01GetPosition = "0"
'End If

End Function

Function Stream02GetPosition(ByVal WTypeDisplay As Long) As Long

'Const StrTime = 1
'Const StrByte = 2

Dim PosByte As Long
Dim PosTime As Long

If Stream02IsPlaying = True Then
    PosByte = BASS_ChannelGetPosition(Strm2)  'stream file position (Bytes)
    PosTime = CLng(BASS_ChannelBytes2Seconds(Strm2, PosByte))
    
    Select Case WTypeDisplay
        Case StrByte
            Stream02GetPosition = PosByte
            
        Case StrTime
            Stream02GetPosition = PosTime
            'BytesPS = Stream02GetBytesPS
            'Stream02GetPosition = Position / BytesPS

    End Select
Else
    Stream02GetPosition = 0
End If

End Function

Sub Stream02Restart()

If BASS_ChannelSetPosition(Strm2, 0) = BASSFALSE Then
    DisplayMsg LoadResString(160)   'no se puede
    Exit Sub
End If

End Sub

Sub Stream01Restart()

If BASS_ChannelSetPosition(Strm1, 0) = BASSFALSE Then
    DisplayMsg LoadResString(160)   'no se puede
    Exit Sub
End If

End Sub

Sub Music02Restart()

' Play the music from the start
'If BASS_MusicPlayEx(Msc2, 0, -1, BASSTRUE) = BASSFALSE Then
'    DisplayMsg LoadResString(160)   'no se puede
'    Exit Sub
'End If

End Sub

Sub Music01Restart()

' Play the music from the start
'If BASS_MusicPlayEx(Msc1, 0, -1, BASSTRUE) = BASSFALSE Then
'    DisplayMsg LoadResString(160)   'no se puede
'    Exit Sub
'End If

End Sub

Function Music02Load(WFileName As String, ByVal WFlagMusic As Long, LastHandle As String) As String

'retorna NotOk si hay algo mal
'retorna Music (new handle) si fue satisfactorio

Dim ModHandle2 As Long
Dim Mode1 As Long, Mode2 As Long, Mode3 As Long, Mode4 As Long, Mode5 As Long
Dim Mode6 As Long, Mode7 As Long

'verificamos si hay un handle anterior y lo eliminamos
If LastHandle = "Music" Then
    BASS_MusicFree Msc2     'music
Else
    If LastHandle = "Stream" Then
        BASS_StreamFree Strm2   'stream
    Else
        BASS_StreamFree Strm2   'stream
    End If
End If

'***************
'* Music flags *
'***************
'Global Const BASS_MUSIC_RAMP = 1       ' normal ramping
'Global Const BASS_MUSIC_RAMPS = 2      ' sensitive ramping
' Ramping doesn't take a lot of extra processing and improves
' the sound quality by removing "clicks". Sensitive ramping will
' leave sharp attacked samples, unlike normal ramping.
'Global Const BASS_MUSIC_LOOP = 4       ' loop music
'Global Const BASS_MUSIC_FT2MOD = 16    ' play .MOD as FastTracker 2 does
'Global Const BASS_MUSIC_PT1MOD = 32    ' play .MOD as ProTracker 1 does
'Global Const BASS_MUSIC_MONO = 64      ' force mono mixing (less CPU usage)
'Global Const BASS_MUSIC_3D = 128       ' enable 3D functionality
'Global Const BASS_MUSIC_SURROUND = 512 'surround sound
'Global Const BASS_MUSIC_SURROUND2 = 1024 'surround sound (mode 2)
'Global Const BASS_MUSIC_FX = 4096      'enable DX8 effects
'Global Const BASS_MUSIC_CALCLEN = 8192 'calculate playback length

'gets the config device data
ConfigData = OpenConfigFile

Select Case ConfigData.Aud_Cual     '1=Mono     2=Stereo
    Case 1
        Mode1 = BASS_MUSIC_MONO
    Case Else
        Mode1 = 0
End Select
Select Case ConfigData.Aud_Mode     '1=Normal   2=A3d   3=3d    4=Ogg
    Case 3
        Mode2 = BASS_MUSIC_3D
    Case Else
        Mode2 = 0
End Select
Select Case ConfigData.Aud_Mod_Type     '1=Ramp Normal      2=Ramp Sensitive
    Case 2
        Mode3 = BASS_MUSIC_RAMPS
    Case Else
        Mode3 = BASS_MUSIC_RAMP
End Select
Select Case ConfigData.Aud_Mod_Cual     '0=None     1=Surround
    Case 1
        Mode4 = BASS_MUSIC_SURROUND
    Case Else
        Mode4 = 0
End Select
Select Case ConfigData.Aud_Mod_Mode     '1=as FT2   2=as PT2
    Case 2
        Mode5 = BASS_MUSIC_PT1MOD
    Case Else
        Mode5 = BASS_MUSIC_FT2MOD
End Select

Mode6 = BASS_MUSIC_FX
Mode7 = BASS_MUSIC_CALCLEN

Select Case WFlagMusic
    Case BASS_MUSIC_LOOP      ' loop music
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, Mode1 Or Mode2 Or Mode3 Or Mode4 Or Mode5 Or Mode6 Or Mode7 Or BASS_MUSIC_LOOP, 0)
        If ModHandle2 = 0 Then
            DisplayMsg LoadResString(159)   '& " " & LoadResString(172)
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If
        
    Case Else
        ModHandle2 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, Mode1 Or Mode2 Or Mode3 Or Mode4 Or Mode5 Or Mode6 Or Mode7, 0)
        If ModHandle2 = 0 Then
            DisplayMsg LoadResString(159)
            Music02Load = "NotOk"
            Exit Function
        Else
            Msc2 = ModHandle2
            Music02Load = "Music"
        End If
End Select

End Function

Function Music01Load(WFileName As String, ByVal WFlagMusic As Long, LastHandle As String) As String

'retorna NotOk si hay algo mal
'retorna Music (new handle) si fue satisfactorio

Dim ModHandle1 As Long
Dim Mode1 As Long, Mode2 As Long, Mode3 As Long, Mode4 As Long, Mode5 As Long
Dim Mode6 As Long, Mode7 As Long

'verificamos si hay un handle anterior y lo eliminamos
If LastHandle = "Music" Then
    BASS_MusicFree Msc1     'music
Else
    If LastHandle = "Stream" Then
        BASS_StreamFree Strm1   'stream
    Else
        BASS_StreamFree Strm1   'stream
    End If
End If

'***************
'* Music flags *
'***************
'Global Const BASS_MUSIC_RAMP = 1       ' normal ramping
'Global Const BASS_MUSIC_RAMPS = 2      ' sensitive ramping
' Ramping doesn't take a lot of extra processing and improves
' the sound quality by removing "clicks". Sensitive ramping will
' leave sharp attacked samples, unlike normal ramping.
'Global Const BASS_MUSIC_LOOP = 4       ' loop music
'Global Const BASS_MUSIC_FT2MOD = 16    ' play .MOD as FastTracker 2 does
'Global Const BASS_MUSIC_PT1MOD = 32    ' play .MOD as ProTracker 1 does
'Global Const BASS_MUSIC_MONO = 64      ' force mono mixing (less CPU usage)
'Global Const BASS_MUSIC_3D = 128       ' enable 3D functionality
'Global Const BASS_MUSIC_SURROUND = 512 'surround sound
'Global Const BASS_MUSIC_SURROUND2 = 1024 'surround sound (mode 2)
'Global Const BASS_MUSIC_FX = 4096      'enable DX8 effects
'Global Const BASS_MUSIC_CALCLEN = 8192 'calculate playback length

'gets the config device data
ConfigData = OpenConfigFile

Select Case ConfigData.Aud_Cual     '1=Mono     2=Stereo
    Case 1
        Mode1 = BASS_MUSIC_MONO
    Case Else
        Mode1 = 0
End Select
Select Case ConfigData.Aud_Mode     '1=Normal   2=A3d   3=3d    4=Ogg
    Case 3
        Mode2 = BASS_MUSIC_3D
    Case Else
        Mode2 = 0
End Select
Select Case ConfigData.Aud_Mod_Type     '1=Ramp Normal      2=Ramp Sensitive
    Case 2
        Mode3 = BASS_MUSIC_RAMPS
    Case Else
        Mode3 = BASS_MUSIC_RAMP
End Select
Select Case ConfigData.Aud_Mod_Cual     '0=None     1=Surround
    Case 1
        Mode4 = BASS_MUSIC_SURROUND
    Case Else
        Mode4 = 0
End Select
Select Case ConfigData.Aud_Mod_Mode     '1=as FT2   2=as PT2
    Case 2
        Mode5 = BASS_MUSIC_PT1MOD
    Case Else
        Mode5 = BASS_MUSIC_FT2MOD
End Select

Mode6 = BASS_MUSIC_FX
Mode7 = BASS_MUSIC_CALCLEN

Select Case WFlagMusic
    Case BASS_MUSIC_LOOP      ' loop music
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, Mode1 Or Mode2 Or Mode3 Or Mode4 Or Mode5 Or Mode6 Or Mode7 Or BASS_MUSIC_LOOP, 0)
        If ModHandle1 = 0 Then
            DisplayMsg LoadResString(159)   '& " " & LoadResString(172)
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If
        
    Case Else
        ModHandle1 = BASS_MusicLoad(BASSFALSE, WFileName, 0, 0, Mode1 Or Mode2 Or Mode3 Or Mode4 Or Mode5 Or Mode6 Or Mode7, 0)
        If ModHandle1 = 0 Then
            DisplayMsg LoadResString(159)
            Music01Load = "NotOk"
            Exit Function
        Else
            Msc1 = ModHandle1
            Music01Load = "Music"
        End If
End Select

End Function

Function Stream02Load(WFileName As String, LastHandle As String) As String

'retorna NotOk si hay algo mal
'retorna Stream (new handle) si fue satisfactorio

Dim StreamHandle2 As Long
Dim Mode1 As Long, Mode2 As Long, Mode3 As Long, Mode4 As Long, Mode5 As Long

'verificamos si hay un handle anterior y lo eliminamos
If LastHandle = "Music" Then
    BASS_MusicFree Msc2     'music
Else
    If LastHandle = "Stream" Then
        Stream02Clear
        'BASS_StreamFree Strm2   'stream
    Else
        Stream02Clear
        'BASS_StreamFree Strm2   'stream
    End If
End If

'*********************
'* Sample info flags *
'*********************
'Global Const BASS_SAMPLE_8BITS = 1             ' 8 bit, else 16 bit
'Global Const BASS_SAMPLE_MONO = 2              ' mono, else stereo
'Global Const BASS_SAMPLE_3D = 8                ' 3D functionality enabled
'Global Const BASS_SAMPLE_FX = 128              ' the DX8 effects are enabled
'Global Const BASS_MP3_HALFRATE = 65536         ' reduced quality MP3/MP2/MP1 (half sample rate)
'Global Const BASS_MP3_SETPOS = 131072          ' enable pin-point seeking on the MP3/MP2/MP1/OGG

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
        Mode3 = 0
    Case 4
        Mode3 = 0
    Case Else
        Mode3 = 0
End Select

Mode4 = BASS_MP3_SETPOS
Mode5 = BASS_SAMPLE_FX

StreamHandle2 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, Mode1 Or Mode2 Or Mode3 Or Mode4 Or Mode5)
If StreamHandle2 = 0 Then
    DisplayMsg LoadResString(159)
    Stream02Load = "NotOk"
Else
    Strm2 = StreamHandle2
    Stream02Load = "Stream"
End If

End Function

Function Stream01Load(WFileName As String, LastHandle As String) As String

'retorna NotOk si hay algo mal
'retorna Stream (new handle) si fue satisfactorio

Dim StreamHandle1 As Long
Dim Mode1 As Long, Mode2 As Long, Mode3 As Long, Mode4 As Long, Mode5 As Long

'verificamos si hay un handle anterior y lo eliminamos
If LastHandle = "Music" Then
    BASS_MusicFree Msc1     'music
Else
    If LastHandle = "Stream" Then
        Stream01Clear
    Else
        Stream01Clear
    End If
End If

'*********************
'* Sample info flags *
'*********************
'Global Const BASS_SAMPLE_8BITS = 1             ' 8 bit, else 16 bit
'Global Const BASS_SAMPLE_MONO = 2              ' mono, else stereo
'Global Const BASS_SAMPLE_3D = 8                ' 3D functionality enabled
'Global Const BASS_SAMPLE_FX = 128              ' the DX8 effects are enabled
'Global Const BASS_MP3_HALFRATE = 65536         ' reduced quality MP3/MP2/MP1 (half sample rate)
'Global Const BASS_MP3_SETPOS = 131072          ' enable pin-point seeking on the MP3/MP2/MP1/OGG

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

StreamHandle1 = BASS_StreamCreateFile(BASSFALSE, WFileName, 0, 0, Mode1 Or Mode2 Or Mode3 Or Mode4 Or Mode5)
If StreamHandle1 = 0 Then
    DisplayMsg LoadResString(159)
    Stream01Load = "NotOk"
Else
    Strm1 = StreamHandle1
    Stream01Load = "Stream"
End If

End Function

Sub Music02Clear()

BASS_MusicFree Msc2

End Sub

Sub Music01Clear()

BASS_MusicFree Msc1

End Sub

Sub Music02Stop()

If BASS_ChannelStop(Msc2) = BASSFALSE Then
    DisplayMsg LoadResString(181)
    Exit Sub
End If

End Sub

Sub Music01Stop()

If BASS_ChannelStop(Msc1) = BASSFALSE Then
    DisplayMsg LoadResString(181)
    Exit Sub
End If

End Sub

Sub Music02Play()

'If BASS_MusicPlay(Msc2) = BASSFALSE Then
'    DisplayMsg LoadResString(159)
'    Exit Sub
'End If

End Sub

Sub Music01Play()

'If BASS_MusicPlay(Msc1) = BASSFALSE Then
'    DisplayMsg LoadResString(159)
'    Exit Sub
'End If

End Sub

Sub Stream02Clear()

'remove the last sync
'Result = StreamRmvSync(2)

BASS_StreamFree Strm2

End Sub

Sub Stream01Clear()

'removes the last sync
'Result = StreamRmvSync(1)

BASS_StreamFree Strm1

End Sub

Sub Stream02Stop()

' Stop the stream
If BASS_ChannelStop(Strm2) = BASSFALSE Then
    'DisplayMsg LoadResString(182)
    Exit Sub
End If

End Sub

Sub Stream01Stop()

'Stop the stream
If BASS_ChannelStop(Strm1) = BASSFALSE Then
    'DisplayMsg LoadResString(182)
    Exit Sub
End If

End Sub

Sub Stream02Play(ByVal WFlagStrmSample As Long)

'Play stream, not flushed
Select Case WFlagStrmSample
    Case BASS_SAMPLE_LOOP
        'If BASS_StreamPlay(Strm2, BASSFALSE, BASS_SAMPLE_LOOP) = BASSFALSE Then
        If BASS_ChannelPlay(Strm2, BASSFALSE) = BASSFALSE Then
            DisplayMsg LoadResString(183) & " " & LoadResString(172)
        End If
    Case Else
        'If BASS_StreamPlay(Strm2, BASSFALSE, 0) = BASSFALSE Then
        If BASS_ChannelPlay(Strm2, BASSFALSE) = BASSFALSE Then
            DisplayMsg LoadResString(159)
        End If
End Select

End Sub

Sub Stream01Play(ByVal WFlagStrmSample As Long)

'Play stream
Select Case WFlagStrmSample
    Case BASS_SAMPLE_LOOP
        'If BASS_StreamPlay(Strm1, BASSFALSE, BASS_SAMPLE_LOOP) = BASSFALSE Then
         If BASS_ChannelPlay(Strm1, BASSFALSE) = BASSFALSE Then
            DisplayMsg LoadResString(183) & " " & LoadResString(172)
        End If
    Case Else
        'If BASS_StreamPlay(Strm1, BASSFALSE, 0) = BASSFALSE Then
        If BASS_ChannelPlay(Strm1, BASSFALSE) = BASSFALSE Then
            DisplayMsg LoadResString(159)
        End If
End Select

End Sub

Function Stream01GetPosition(ByVal WTypeDisplay As Long) As Long

'Const StrTime = 1
'Const StrByte = 2

Dim PosByte As Long
Dim PosTime As Long

If Stream01IsPlaying = True Then
    PosByte = BASS_ChannelGetPosition(Strm1)  'get stream file position (Bytes)
    PosTime = CLng(BASS_ChannelBytes2Seconds(Strm1, PosByte)) 'convert byte 2 sec.
    
    Select Case WTypeDisplay
        Case StrByte
            Stream01GetPosition = PosByte
            
        Case StrTime
            Stream01GetPosition = PosTime
            'BytesPS = Stream01GetBytesPS
            'Stream01GetPosition = Position / BytesPS
            
    End Select
Else
    Stream01GetPosition = 0
End If

End Function

Function Stream01GetLEFTLevel() As Long

Dim Level As Long, LLLeft As Long
Dim A As Long

If Stream01IsPlaying = True Then
    Level = BASS_ChannelGetLevel(Strm1)  'stream file level meter
    A = 93
    If (A > 0) Then
        If LoWord(Level) >= A Then
            LLLeft = LoWord(Level)
        Else
            LLLeft = LoWord(Level)
            A = A * 2 / 3
        End If
    End If
    Stream01GetLEFTLevel = LLLeft
Else
    Stream01GetLEFTLevel = 0
End If

End Function


Sub PHPlay()

'//////////////// MAL FUNCIONAMIENTO /////////////////////
'/////////////// modificar y eliminar ///////////////////

Dim Result As String
Dim Oldi As Integer
Dim Newi As Integer

'cargar el tema horario
Result = StreamPHLoad(TopMenu.PHName.Caption)
If Result = "NotOk" Then Exit Sub
'Setear el syncronismo para volver el volumen a normal
StreamPHSetSync
'reproducir el tema horario
StreamPHPlay
'bajar el volumen del tema que se esta reproduciendo
Call AutoDwVol
'prepararse para continuar con el siguiente
Oldi = CInt(TopMenu.NumberIdx.Caption)
Newi = Oldi + 1
TopMenu.NumberIdx.Caption = Newi

End Sub


Sub PHRmv()

Dim RR As String

RR = StreamPHRmvSync

End Sub

Public Sub Tanda02Play(WFName As String, WFTitle As String, WFType As String, WSync As String)

Dim FileN As String, FileTP As String, SSTitle As String
Dim IntrSeg As Long
Dim Result As String, ResultDisplay As String

'extraccion de datos necesarios para la reproduccion
FileN = WFName       'nombre y path del archivo
FileTP = WFType      'tipo de archivo Music or Stream?
SSTitle = WFTitle     'Titulo del archivo

Select Case WSync
    Case "Yes" '-------------------------------------------------------------------
        If FileTP = "Stream" Then
            If Est12Control.StopLabel2.Caption = "Stream" Then
                Result = Stream02Load(FileN, "Stream")
            Else
                If Est12Control.StopLabel2.Caption = "Music" Then
                    Result = Stream02Load(FileN, "Music")
                Else
                    Result = Stream02Load(FileN, "Stream")
                End If
            End If
            'lets sets the sync
            IntrSeg = CLng(Tanda01.Intr.text)
            ResultDisplay = StreamSetSyncPos(2, IntrSeg)
            'continue
            Tanda01.T2Name.Caption = SSTitle
            Tanda01.T2Name.ForeColor = &HFFFF00     'celeste claro(activado)
            Est12Control.StopLabel2.Caption = Result
            Est12Control.Origen2.Caption = "T2"
            'activamos el fx
            If Est12Control.LblFX.Caption = "NoFX" Then
                'xxxxxx
            Else
                Call InitEffect(2, "Stream")
            End If
            'lets play the file
            Stream02Play (0)
            Tanda01.Caption = LoadResString(1007)   'reproduciendo
        Else
            If FileTP = "Music" Then
                If Est12Control.StopLabel2.Caption = "Music" Then
                    '0 = default music load
                    Result = Music02Load(FileN, 0, "Music")
                Else
                    If Est12Control.StopLabel2.Caption = "Stream" Then
                        Result = Music02Load(FileN, 0, "Stream")
                    Else
                        Result = Music02Load(FileN, 0, "Music")
                    End If
                End If
                'get config settings, module get name?
                ConfigData = OpenConfigFile
                If ConfigData.Gen_AutoName = 1 Then
                    'If BASS_MusicGetNameString(Msc2) = "" Then
                    '    Tanda01.T2Name.Caption = SSTitle
                    'Else
                        'Tanda01.T2Name.Caption = BASS_MusicGetNameString(Msc2)
                    'End If
                Else
                    Tanda01.T2Name.Caption = SSTitle
                End If
                Tanda01.T2Name.ForeColor = &HFFFF00     'celeste claro(activado)
                Est12Control.StopLabel2.Caption = Result
                Est12Control.Origen2.Caption = "T2"
                'activamos el fx
                If Est12Control.LblFX.Caption = "NoFX" Then
                    'xxxxxx
                Else
                    Call InitEffect(2, "Music")
                End If
                'lets play the music
                Music02Play
                Tanda01.Caption = LoadResString(1007)   'reproduciendo
            Else
                Tanda01.T2Name.ForeColor = &H808000     'celeste oscuro(desactivado)
                Tanda01.T2Name.Caption = "---"
            End If
        End If
    Case "No" '-------------------------------------------------------------------
        If FileTP = "Stream" Then
            If Est12Control.StopLabel2.Caption = "Stream" Then
                Result = Stream02Load(FileN, "Stream")
            Else
                If Est12Control.StopLabel2.Caption = "Music" Then
                    Result = Stream02Load(FileN, "Music")
                Else
                    Result = Stream02Load(FileN, "Stream")
                End If
            End If
            Tanda01.T2Name.Caption = SSTitle
            Tanda01.T2Name.ForeColor = &HFFFF00     'celeste claro(activado)
            Est12Control.StopLabel2.Caption = Result
            Est12Control.Origen2.Caption = "T2"
            'activamos el fx
            If Est12Control.LblFX.Caption = "NoFX" Then
                'xxxxxx
            Else
                Call InitEffect(2, "Stream")
            End If
            'lets play the file
            Stream02Play (0)
            Tanda01.Caption = LoadResString(1007)   'reproduciendo
        Else
            If FileTP = "Music" Then
                If Est12Control.StopLabel2.Caption = "Music" Then
                    '0 = default music load
                    Result = Music02Load(FileN, 0, "Music")
                Else
                    If Est12Control.StopLabel2.Caption = "Stream" Then
                        Result = Music02Load(FileN, 0, "Stream")
                    Else
                        Result = Music02Load(FileN, 0, "Music")
                    End If
                End If
                'get config settings, module get name?
                ConfigData = OpenConfigFile
                If ConfigData.Gen_AutoName = 1 Then
                    'If BASS_MusicGetNameString(Msc2) = "" Then
                    '    Tanda01.T2Name.Caption = SSTitle
                    'Else
                    '    Tanda01.T2Name.Caption = BASS_MusicGetNameString(Msc2)
                    'End If
                Else
                    Tanda01.T2Name.Caption = SSTitle
                End If
                Tanda01.T2Name.ForeColor = &HFFFF00     'celeste claro(activado)
                Est12Control.StopLabel2.Caption = Result
                Est12Control.Origen2.Caption = "T2"
                'activamos el fx
                If Est12Control.LblFX.Caption = "NoFX" Then
                    'xxxxxx
                Else
                    Call InitEffect(2, "Music")
                End If
                'lets play the music
                Music02Play
                Tanda01.Caption = LoadResString(1007)   'reproduciendo
            Else
                Tanda01.T2Name.ForeColor = &H808000     'celeste oscuro(desactivado)
                Tanda01.T2Name.Caption = "---"
            End If
        End If
End Select

End Sub

Public Sub Tanda01Play(WFName As String, WFTitle As String, WFType As String, WSync As String)

Dim FileN As String, FileTP As String, SSTitle As String
Dim IntrSeg As Long
Dim Result As String, ResultDisplay As String

'extraccion de datos necesarios para la reproduccion
FileN = WFName       'nombre y path del archivo
FileTP = WFType      'tipo de archivo Music or Stream?
SSTitle = WFTitle     'Titulo del archivo

Select Case WSync
    Case "Yes" '-----------------------------------------------------------------
        If FileTP = "Stream" Then
            If Est12Control.StopLabel1.Caption = "Stream" Then
                Result = Stream01Load(FileN, "Stream")
            Else
                If Est12Control.StopLabel1.Caption = "Music" Then
                    Result = Stream01Load(FileN, "Music")
                Else
                    Result = Stream01Load(FileN, "Stream")
                End If
            End If
            'sets the file sync
            IntrSeg = CLng(Tanda01.Intr.text)
            ResultDisplay = StreamSetSyncPos(1, IntrSeg)
            'continue
            Tanda01.T1Name.Caption = SSTitle
            Tanda01.T1Name.ForeColor = &HFFFF00     'celeste claro(activado)
            Est12Control.StopLabel1.Caption = Result
            Est12Control.Origen1.Caption = "T1"
            'activamos el fx
            If Est12Control.LblFX.Caption = "NoFX" Then
                'xxxxxx
            Else
                Call InitEffect(1, "Stream")
            End If
            'lets play the file
            Stream01Play (0)
            Tanda01.Caption = LoadResString(1007)   'reproduciendo
        Else
            If FileTP = "Music" Then
                If Est12Control.StopLabel1.Caption = "Music" Then
                    '0 = default music load
                    Result = Music01Load(FileN, 0, "Music")
                Else
                    If Est12Control.StopLabel1.Caption = "Stream" Then
                        Result = Music01Load(FileN, 0, "Stream")
                    Else
                        Result = Music01Load(FileN, 0, "Music")
                    End If
                End If
                'get config settings, module get name?
                ConfigData = OpenConfigFile
                If ConfigData.Gen_AutoName = 1 Then
                    'If BASS_MusicGetNameString(Msc1) = "" Then
                    '    Tanda01.T1Name.Caption = SSTitle
                    'Else
                    '    Tanda01.T1Name.Caption = BASS_MusicGetNameString(Msc1)
                    'End If
                Else
                    Tanda01.T1Name.Caption = SSTitle
                End If
                Tanda01.T1Name.ForeColor = &HFFFF00     'celeste claro(activado)
                Est12Control.StopLabel1.Caption = Result
                Est12Control.Origen1.Caption = "T1"
                'activamos el fx
                If Est12Control.LblFX.Caption = "NoFX" Then
                    'xxxxxx
                Else
                    Call InitEffect(1, "Music")
                End If
                'lets play the music
                Music01Play
                Tanda01.Caption = LoadResString(1007)   'reproduciendo
            Else
                Tanda01.T1Name.ForeColor = &H808000     'celeste oscuro(desactivado)
                Tanda01.T1Name.Caption = "---"
            End If
        End If
    Case "No" '-----------------------------------------------------------------
        If FileTP = "Stream" Then
            If Est12Control.StopLabel1.Caption = "Stream" Then
                Result = Stream01Load(FileN, "Stream")
            Else
                If Est12Control.StopLabel1.Caption = "Music" Then
                    Result = Stream01Load(FileN, "Music")
                Else
                    Result = Stream01Load(FileN, "Stream")
                End If
            End If
            Tanda01.T1Name.Caption = SSTitle
            Tanda01.T1Name.ForeColor = &HFFFF00     'celeste claro(activado)
            Est12Control.StopLabel1.Caption = Result
            Est12Control.Origen1.Caption = "T1"
            'activamos el fx
            If Est12Control.LblFX.Caption = "NoFX" Then
                'xxxxxx
            Else
                Call InitEffect(1, "Stream")
            End If
            'lets play the file nosync
            Stream01Play (0)
            Tanda01.Caption = LoadResString(1007)   'reproduciendo
        Else
            If FileTP = "Music" Then
                If Est12Control.StopLabel1.Caption = "Music" Then
                    '0 = default music load
                    Result = Music01Load(FileN, 0, "Music")
                Else
                    If Est12Control.StopLabel1.Caption = "Stream" Then
                        Result = Music01Load(FileN, 0, "Stream")
                    Else
                        Result = Music01Load(FileN, 0, "Music")
                    End If
                End If
                'get config settings, module get name?
                ConfigData = OpenConfigFile
                If ConfigData.Gen_AutoName = 1 Then
                    'If BASS_MusicGetNameString(Msc1) = "" Then
                    '    Tanda01.T1Name.Caption = SSTitle
                    'Else
                    '    Tanda01.T1Name.Caption = BASS_MusicGetNameString(Msc1)
                    'End If
                Else
                    Tanda01.T1Name.Caption = SSTitle
                End If
                Tanda01.T1Name.ForeColor = &HFFFF00     'celeste claro(activado)
                Est12Control.StopLabel1.Caption = Result
                Est12Control.Origen1.Caption = "T1"
                'activamos el fx
                If Est12Control.LblFX.Caption = "NoFX" Then
                    'xxxxxx
                Else
                    Call InitEffect(1, "Music")
                End If
                'lets play the music
                Music01Play
                Tanda01.Caption = LoadResString(1007)   'reproduciendo
            Else
                Tanda01.T1Name.ForeColor = &H808000     'celeste oscuro(desactivado)
                Tanda01.T1Name.Caption = "---"
            End If
        End If
End Select

End Sub

Function Estacion02Play(ByVal WConNum As Integer) As String

Dim FileN As String, FileTP As String, SSTitle As String
Dim ResultFile As Boolean, Result As String

'extraccion de datos necesarios para la reproduccion
FileN = Est12Data.N2(WConNum).Caption       'nombre y path del archivo
FileTP = Est12Data.V2(WConNum).Caption      'tipo de archivo Music or Stream?
SSTitle = Est12Data.c2(WConNum).Caption     'Titulo del archivo

'chequeos necesarios
If FileN = "" Or FileN = " " Then
    Est02.Label1.ForeColor = &H808000     'celeste oscuro(desactivado)
    Est02.Label1.Caption = "---"
    Estacion02Play = "NotOk"
    Exit Function
End If

'chequeamos si el archivo existe
ResultFile = FileExist(FileN)
If ResultFile = False Then
    Estacion02Play = "NotOk"
    Exit Function
End If

If Est12Control.StopLabel2.Caption = "Stream" Then
    Stream02Stop       'stream stop
Else
    If Est12Control.StopLabel2.Caption = "Music" Then
        Music02Stop         'music stop
    End If
End If

If FileTP = "Stream" Then
    If Est12Control.StopLabel2.Caption = "Stream" Then
        Result = Stream02Load(FileN, "Stream")
    Else
        If Est12Control.StopLabel2.Caption = "Music" Then
            Result = Stream02Load(FileN, "Music")
        Else
            Result = Stream02Load(FileN, "Stream")
        End If
    End If
    Est02.Label1.Caption = SSTitle
    Est02.Label1.ForeColor = &HFFFF00     'celeste claro(activado)
    Est12Control.StopLabel2.Caption = Result
    Est12Control.Origen2.Caption = "E2"
    'activamos el fx
    If Est12Control.LblFX.Caption = "NoFX" Then
        'xxxxxx
    Else
        Call InitEffect(2, "Stream")
    End If
    'lets play the file
    If Est02.LAplay.ForeColor = &HFFFF00 Then   'autoplay claro
        If Est02.pcontup.Visible = True Then    'play continuous activado?
            Stream02Play (BASS_SAMPLE_LOOP)
            Est02.Caption = LoadResString(1004)     'reproduciendo
        Else
            Stream02Play (0)
            Est02.Caption = LoadResString(1004) 'reproduciendo
        End If
    End If
Else
    If FileTP = "Music" Then
        If Est12Control.StopLabel2.Caption = "Music" Then
            If Est02.pcontup.Visible = True Then
                Result = Music02Load(FileN, BASS_MUSIC_LOOP, "Music")
            Else
                Result = Music02Load(FileN, 0, "Music")
            End If
        Else
            If Est12Control.StopLabel2.Caption = "Stream" Then
                If Est02.pcontup.Visible = True Then
                    Result = Music02Load(FileN, BASS_MUSIC_LOOP, "Stream")
                Else
                    Result = Music02Load(FileN, 0, "Stream")
                End If
            Else
                If Est02.pcontup.Visible = True Then
                    Result = Music02Load(FileN, BASS_MUSIC_LOOP, "Music")
                Else
                    Result = Music02Load(FileN, 0, "Music")
                End If
            End If
        End If
        'get config settings, module get name?
        ConfigData = OpenConfigFile
        If ConfigData.Gen_AutoName = 1 Then
            'If BASS_MusicGetNameString(Msc2) = "" Then
                Est02.Label1.Caption = SSTitle
            'Else
            '    Est02.Label1.Caption = BASS_MusicGetNameString(Msc2)
            'End If
        Else
            Est02.Label1.Caption = SSTitle
        End If
        Est02.Label1.ForeColor = &HFFFF00     'celeste claro(activado)
        Est12Control.StopLabel2.Caption = Result
        Est12Control.Origen2.Caption = "E2"
        'activamos el fx
        If Est12Control.LblFX.Caption = "NoFX" Then
            'xxxxxx
        Else
            Call InitEffect(2, "Music")
        End If
        'lets play the music
        If Est02.LAplay.ForeColor = &HFFFF00 Then 'claro
            Music02Play
            Est02.Caption = LoadResString(1004) 'reproduciendo
        End If
    Else
        Est02.Label1.ForeColor = &H808000     'celeste oscuro(desactivado)
        Est02.Label1.Caption = "---"
    End If
End If

Estacion02Play = "Ok"

End Function
Function Estacion01Play(ByVal WConNum As Integer) As String

Dim FileN As String, FileTP As String, SSTitle As String
Dim ResultFile As Boolean, Result As String

'extraccion de datos necesarios para la reproduccion
FileN = Est12Data.N1(WConNum).Caption       'nombre y path del archivo
FileTP = Est12Data.V1(WConNum).Caption      'tipo de archivo Music or Stream?
SSTitle = Est12Data.c1(WConNum).Caption     'Titulo del archivo

'chequeos necesarios
If FileN = "" Or FileN = " " Then
    Est01.Label1.ForeColor = &H808000     'celeste oscuro(desactivado)
    Est01.Label1.Caption = "---"
    Estacion01Play = "NotOk"
    Exit Function
End If

'chequeamos si el archivo existe
ResultFile = FileExist(FileN)
If ResultFile = False Then
    Estacion01Play = "NotOk"
    Exit Function
End If

If Est12Control.StopLabel1.Caption = "Stream" Then
    Stream01Stop       'stream stop
Else
    If Est12Control.StopLabel1.Caption = "Music" Then
        Music01Stop         'music stop
    End If
End If

If FileTP = "Stream" Then
    If Est12Control.StopLabel1.Caption = "Stream" Then
        Result = Stream01Load(FileN, "Stream")
    Else
        If Est12Control.StopLabel1.Caption = "Music" Then
            Result = Stream01Load(FileN, "Music")
        Else
            Result = Stream01Load(FileN, "Stream")
        End If
    End If
    Est01.Label1.Caption = SSTitle
    Est01.Label1.ForeColor = &HFFFF00     'celeste claro(activado)
    Est12Control.StopLabel1.Caption = Result
    Est12Control.Origen1.Caption = "E1"
    'activamos el fx
    If Est12Control.LblFX.Caption = "NoFX" Then
        'xxxxxx
    Else
        Call InitEffect(1, "Stream")
    End If
    'lets play the file
    If Est01.LAplay.ForeColor = &HFFFF00 Then   'autoplay claro
        If Est01.pcontup.Visible = True Then    'play continuo activado?
            Stream01Play (BASS_SAMPLE_LOOP)
            Est01.TitelBar1.Caption = LoadResString(1001) 'reproduciendo
        Else
            Stream01Play (0)
            Est01.TitelBar1.Caption = LoadResString(1001) 'reproduciendo
        End If
    End If
Else
    If FileTP = "Music" Then
        If Est12Control.StopLabel1.Caption = "Music" Then
            If Est01.pcontup.Visible = True Then
                Result = Music01Load(FileN, BASS_MUSIC_LOOP, "Music")
            Else
                Result = Music01Load(FileN, 0, "Music")
            End If
        Else
            If Est12Control.StopLabel1.Caption = "Stream" Then
                If Est01.pcontup.Visible = True Then
                    Result = Music01Load(FileN, BASS_MUSIC_LOOP, "Stream")
                Else
                    Result = Music01Load(FileN, 0, "Stream")
                End If
            Else
                If Est01.pcontup.Visible = True Then
                    Result = Music01Load(FileN, BASS_MUSIC_LOOP, "Music")
                Else
                    Result = Music01Load(FileN, 0, "Music")
                End If
            End If
        End If
        'get config settings, module get name?
        ConfigData = OpenConfigFile
        If ConfigData.Gen_AutoName = 1 Then
            'If BASS_MusicGetNameString(Msc1) = "" Then
                Est01.Label1.Caption = SSTitle
            'Else
            '    Est01.Label1.Caption = BASS_MusicGetNameString(Msc1)
            'End If
        Else
            Est01.Label1.Caption = SSTitle
        End If
        Est01.Label1.ForeColor = &HFFFF00     'celeste claro(activado)
        Est12Control.StopLabel1.Caption = Result
        Est12Control.Origen1.Caption = "E1"
        'activamos el fx
        If Est12Control.LblFX.Caption = "NoFX" Then
            'xxxxxx
        Else
            Call InitEffect(1, "Music")
        End If
        'lets play the music
        If Est01.LAplay.ForeColor = &HFFFF00 Then 'claro
            Music01Play
            Est01.TitelBar1.Caption = LoadResString(1001) 'reproduciendo
        End If
    Else
        Est01.Label1.ForeColor = &H808000     'celeste oscuro(desactivado)
        Est01.Label1.Caption = "---"
    End If
End If

Estacion01Play = "Ok"
End Function


Sub SYNCPROC_PH(ByVal handle As Long, ByVal channel As Long, ByVal Data As Long, ByVal user As Long)
    
    'CALLBACK FUNCTION !!!
    
    'Similarly in here, write what to do when sync function
    'is called, i.e screen flash etc.
    
    ' NOTE: a sync callback function should be very
    ' quick (eg. just posting a message) as other syncs cannot be processed
    ' until it has finished.
    ' handle : The sync that has occured
    ' channel: Channel that the sync occured in
    ' data   : Additional data associated with the sync's occurance
    ' user   : The 'user' parameter given when calling BASS_ChannelSetSync */
    
    'Sync proc in PH only
    'set the volumen to normal
    Call AutoUpVol
    'call the next file and next sync
    Call FrmTime.PHActive_Click
    
End Sub

