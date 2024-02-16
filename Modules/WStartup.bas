Attribute VB_Name = "Startup"
'//////////////////////////////////////////////////////////
'*
'*  /// Only Radiomaker 1.0 modulo principal for Vb.6+ ///
'*  *********** and is for Radiomaker 1.0 only ***********
'*
'*     Copyright (c) 2002-2022 Only development Inc.
'*     Christian A. Del Monte
'/////////////////////////////////////////////////////////

Option Explicit

Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

'/////////////////////////////////////////
' Get the screen resolution info.
' ** Must be 1024 x 768 (minimum) **
'/////////////////////////////////////////

Private Function GetScreenInfo() As String

'If TopMenu.SysInfo1.WorkAreaWidth < "15360" And TopMenu.SysInfo1.WorkAreaHeight < "11520" Then
'    GetScreenInfo = "NotOk"
'    Exit Function
'Else
'    If TopMenu.SysInfo1.WorkAreaWidth >= "15360" And TopMenu.SysInfo1.WorkAreaHeight >= "11520" Then
'        GetScreenInfo = "Ok"
'        Exit Function
'    End If
'End If

End Function

Private Sub SystemCheck()

Dim CallResult As String
Dim Msg As String, Msg0 As String, Msg3 As String, Msg4 As String
Dim Style, Title, Response

'Extraemos la informacion del disco duro
CallResult = GetHdInfo("c:\")
Select Case CallResult
    Case "Ok"
        'xxx
    Case "NotOk"
        'xxx
End Select

'Extraemos la informacion de la Fecha y Hora del sistema
CallResult = GetTimeDateInfo
Select Case CallResult
    Case "Ok"
        'xxx
    Case "NotOk"
        'xxx
End Select

'Extraemos la informacion de resolucion de pantalla
CallResult = GetScreenInfo
Select Case CallResult
    Case "Ok"
        'xxxxxxxxxx
        'nothing to do... continue the load
    Case "NotOk"
        Msg0 = LoadResString(131)   'resolucion no apropiada
        Msg3 = " "
        Msg4 = LoadResString(132)   'desea continuar?
        Msg = Msg0 & " " & Msg3 & " " & Msg4
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Rm100 - Resolución no soportada."
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            End
        Else
            End
        End If
End Select

'Extraemos la informacion de los dispositivos de audio
CallResult = GetAudioInfo
Select Case CallResult
    Case "None"     'no hay dispositivos de audio
        Msg0 = LoadResString(133)
        Msg3 = " "
        Msg4 = LoadResString(132)
        Msg = Msg0 & " " & Msg3 & " " & Msg4
        Style = vbYesNo + vbCritical + vbDefaultButton2
        Title = "Rm100 - Detección de audio"
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            'xxxxx nothing to do... continue the load
        Else
            End
        End If
    Case "One"      'se encontro un solo dispositivo
        'xxx nothing to do... continue the load
    Case "Two"      'se encontraron dos o mas dispositivos
        'xxx nothing to do... continue the load
End Select

End Sub

Private Function GetHdInfo(Whd As String) As String

'Extraccion de la Informacion del Disco duro
GetHdInfo = "OK"
Exit Function

Error:
GetHdInfo = "NotOk"

End Function

Private Function GetTimeDateInfo() As String

'Extraccion de la informacion de la Fecha y la Hora del sistema
GetTimeDateInfo = "OK"
Exit Function

Error:
GetTimeDateInfo = "NotOK"

End Function

Private Function GetAudioInfo() As String

Dim rtn As Integer

rtn = waveOutGetNumDevs()

'Extraccion de la informacion de los dispositivos de audio
If rtn = 0 Then
   GetAudioInfo = "None"   'NO HAY PLACAS DE AUDIO
Else
    If rtn = 1 Then
        GetAudioInfo = "One"   'HAY UNA SOLA PLACA DE AUDIO
    Else
        If rtn >= 2 Then
            GetAudioInfo = "Two"      'HAY DOS PLACAS DE AUDIO
        Else
            GetAudioInfo = "Two"
        End If
    End If
End If

End Function

'///////////////////////
' Principal RUN module
'///////////////////////

Sub Main()

Dim CallResult As String
Dim Msg As String, Msg0 As String, Msg2 As String
Dim Style, Title

'/// Desabilitamos el timer del RMSplash (Splash Screen)
RMSplash.Timer1.Interval = 0
RMSplash.Timer1.Enabled = False

'/// chequeamos los requisitos del sistema
Call SystemCheck

'/// Habilitamos las ventanas principales de datos
Est12Data.Show
Est12Data.Visible = False
Est12Control.Show
Est12Control.Visible = False

'/// inicializamos el sistema de audio
CallResult = InitDevice(0)
If CallResult = "NotOk" Then        'error al inicializar
    Msg0 = LoadResString(134)       'error display
    Msg2 = " "
    Msg = Msg0 & " " & Msg2
    Style = vbCritical
    Title = "Rm100 - Error."
    MsgBox Msg, Style, Title
    CloseDevice "Stream", "Stream"  'cerramos el dispositivo de audio
    End                             'finalizamos la ejecución del programa
End If

'/// extraemos la informacion del config para los displays
ConfigData = OpenConfigFile

Select Case ConfigData.Aud_Disp_Time    '1=normal   2=rest
    Case 1
        TopMenu.LType.Caption = "Normal"
    Case Else
        TopMenu.LType.Caption = "Restante"
End Select
Select Case ConfigData.Aud_Disp_Wave    '1=normal   2=rest
    Case 1
        TopMenu.OType.Caption = "Normal"
    Case Else
        TopMenu.OType.Caption = "Restante"
End Select
Select Case ConfigData.Aud_Disp_Samp    '1=normal   2=rest
    Case 1
        TopMenu.SType.Caption = "Normal"
    Case Else
        TopMenu.SType.Caption = "Restante"
End Select
    
'/// Actualizamos los displays (por defecto)
'Call RestoreDisplay(1)   'EST01
'Call RestoreDisplay(2)   'EST02
'Call RestoreDisplay(3)   'TANDA 01
'Call RestoreDisplay(4)   'TANDA 02
'Call RestoreDisplay(5)   'TOTAL TIME TANDA 01 Y 02
'Call RestoreDisplay(6)   'LUNCH TIME in PHTIMER
'Call RestoreDisplay(7)   'PROG TANDAS TIMER module
'Call RestoreDisplay(10)  'PROG TANDAS TIME DISPLAY

'/// extraemos el ultimo estado del programa
CallResult = GetState
If CallResult = "NotOk" Then
    'display all default windows
    'ShowWindow "TopMenu"
    'ShowWindow "DwMenu"
    'ShowWindow "Est01"
    'ShowWindow "Est02"
    'ShowWindow "Prg01"
    'ShowWindow "Tnd01"
End If

'/// habilitamos el reloj del topmenu
TopMenu.ClockTimer.Enabled = True
TopMenu.ClockTimer.Interval = 1000

'/// Deshabilitamos el splash screen
Unload RMSplash

'/// chequeamos por la existencia de Plug-Ins
'TopMenu.GetPlugInList "RMPlayer.dll"
'TopMenu.GetPlugInList "RMRipper.dll"
'TopMenu.GetPlugInList "RMVoice.dll"
'TopMenu.GetPlugInList "RMController.dll"
'TopMenu.GetPlugInList "RMXModule.dll"
'TopMenu.GetPlugInList "RMFilter.dll"
'TopMenu.GetPlugInList "RMEditec.dll"
'TopMenu.GetPlugInList "RMDatabase.dll"

'/// Mostramos las sugerencias del dia y Retornamos el control al usuario
'frmTip.Show: frmTip.cmdOK.SetFocus

End Sub
