Attribute VB_Name = "NetShowProccess"
'********************* RM100 *********************
'    RADIO MAKER IMPORT/EXPORT FILE MODULE
'COPYRIGHT (C) 1987-2002 ONLY development inc.
'*************************************************
'Autor: Christian A. Del Monte
'---------------------------------------------------------------
'Modulo para procesar archivos Region/PlayList de NetShow Player
'y convertirlos en información CUE para su posterior proceso en
'Radio Maker.
'---------------------------------------------------------------

Option Explicit

Public Function GetNetShowAudioRegion(WFileName As String, ByVal WEstNum As Long) As String

Dim Wname As String
Dim L1, L2, L3, L4
Dim TimeIn As String
Dim TimeOut As String
Dim Result As String
Dim DataIn As String
Dim DataOut As String

If FileExist(WFileName) = False Then GoSub NoGetAudio

Wname = StripFileFromExt(WFileName)

On Error GoTo NoGetAudio
Open Wname & ".txt" For Input As #30
Line Input #30, L1
Line Input #30, L2
Line Input #30, L3
Line Input #30, L4
Close #30

If L1 = "start_marker_table" Then
    DataIn = Left$(L2, 10)   '=00:00:00.0  h/m/s/ss
    DataOut = Left$(L3, 10)  '=00:00:00.0  h/m/s/ss
    TimeIn = ConvMinToSec(Left$(DataIn, 8))
    TimeOut = ConvMinToSec(Left$(DataOut, 8))
    If TimeIn = "0" Or 0 Then
        MsgBox LoadResString(155), vbCritical
        GoSub NoGetAudio
    End If
    If TimeOut = "0" Or 0 Then
        MsgBox LoadResString(156), vbCritical
        GoSub NoGetAudio
    End If
    Select Case WEstNum
        Case 1
            Result = SetNetShowAudioRegion(TimeIn, TimeOut, 1)
        Case 2
            Result = SetNetShowAudioRegion(TimeIn, TimeOut, 2)
    End Select
Else
    MsgBox LoadResString(157), vbCritical
    GoSub NoGetAudio
End If
GetNetShowAudioRegion = "Ok"
Exit Function

NoGetAudio:
GetNetShowAudioRegion = "NotOk"
End Function

Public Function SetNetShowAudioRegion(WRStart As String, WREnd As String, ByVal WEstNum As Long) As String

Dim BytesPS As Double
Dim nStart As Long, NEnd As Long
Dim FStart As Long, FEnd As Long

'chequeos y asignaciones
If WRStart = "" Or WRStart = " " Then GoSub nop
If WREnd = "" Or WREnd = " " Then GoSub nop
nStart = CLng(WRStart)
NEnd = CLng(WREnd)

Select Case WEstNum
    Case 1
        'covertimos el tiempo (segundos) en bytes
        BytesPS = Stream01GetBytesPS
        If BytesPS = 0 Then
            MsgBox LoadResString(154), vbCritical
            GoSub nop
        End If
        FStart = nStart * BytesPS
        FEnd = NEnd * BytesPS
        'Seteamos las posiciones de comienzo y final de reproduccion
        Est01.E1Pos.SelStart = nStart           'sets the start point in seconds (value)
        Est01.E1Pos.SelLength = NEnd - nStart   'sets the end point in seconds  (value)
        Est01.LblStartCUE.Caption = FStart  'sets the start point in bytes
        Est01.LblEndCue.Caption = FEnd      'sets the end point in bytes
        SetNetShowAudioRegion = "Ok"
        Exit Function
    Case 2
        'covertimos el tiempo (segundos) en bytes
        BytesPS = Stream02GetBytesPS
        If BytesPS = 0 Then
            MsgBox LoadResString(154), vbCritical
            GoSub nop
        End If
        FStart = nStart * BytesPS
        FEnd = NEnd * BytesPS
        'Seteamos las posiciones de comienzo y final de reproduccion
        Est02.E2Pos.SelStart = nStart           'sets the start point in seconds (value)
        Est02.E2Pos.SelLength = NEnd - nStart   'sets the end point in seconds  (value)
        Est02.LblStartCUE.Caption = FStart  'sets the start point in bytes
        Est02.LblEndCue.Caption = FEnd      'sets the end point in bytes
        SetNetShowAudioRegion = "Ok"
        Exit Function
End Select
Exit Function

nop:
SetNetShowAudioRegion = "NotOk"
End Function
