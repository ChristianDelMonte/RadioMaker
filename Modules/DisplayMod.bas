Attribute VB_Name = "DisplayModule"
'********************* RM100 *********************
'         RADIO MAKER DISPLAY MODULE
'COPYRIGHT (C) 1987-2008 ONLY development inc.
'  Christian A. Del Monte
'*************************************************

Option Explicit

'viariabled de manejo de tiempo en relojes TopMenu
Dim LHora As String, LMinutos As String, LSegundos As String
Dim NHora(1 To 2) As String, NMinutos(1 To 2) As String, NSegundos(1 To 2) As String

'variabled de manejo de fecha en relojes TopMenu
Dim LMes As String, LDia As String, LAno As String
Dim NMes(1 To 2) As String, NDia(1 To 2) As String, NAno(1 To 4) As String

'variables para el manejo climatico en relojes topmenu
Dim TempLen As Long, HumeLen As Long
Dim DTemp(1 To 5) As String, DHume(1 To 3) As String

Dim Contador As Integer

Function ClimaDisplay(WMode As Long)

'esta funcion muestra o desabilita el display del clima en el top menu

Select Case WMode
    Case 1  'show
        TopMenu.c1.Visible = True
        TopMenu.c2.Visible = True
        TopMenu.c3.Visible = True
        TopMenu.c4.Visible = True
        TopMenu.c5.Visible = True
        TopMenu.c6.Visible = True
        TopMenu.c7.Visible = True
        TopMenu.c8.Visible = True
        TopMenu.c9.Visible = True
        TopMenu.c0.Visible = True
        
    Case -1 'hide
        TopMenu.c1.Visible = False
        TopMenu.c2.Visible = False
        TopMenu.c3.Visible = False
        TopMenu.c4.Visible = False
        TopMenu.c5.Visible = False
        TopMenu.c6.Visible = False
        TopMenu.c7.Visible = False
        TopMenu.c8.Visible = False
        TopMenu.c9.Visible = False
        TopMenu.c0.Visible = False
End Select
End Function

Function DateDisplay(WMode As Long)

'esta funcion muestra o desabilita el display de la fecha en el top menu

Select Case WMode
    Case 1  'show
        TopMenu.f1.Visible = True
        TopMenu.f2.Visible = True
        TopMenu.f3.Visible = True
        TopMenu.f4.Visible = True
        TopMenu.f5.Visible = True
        TopMenu.f6.Visible = True
        TopMenu.f7.Visible = True
        TopMenu.f8.Visible = True
        TopMenu.f9.Visible = True
        TopMenu.f10.Visible = True
        
    Case -1 'hide
        TopMenu.f1.Visible = False
        TopMenu.f2.Visible = False
        TopMenu.f3.Visible = False
        TopMenu.f4.Visible = False
        TopMenu.f5.Visible = False
        TopMenu.f6.Visible = False
        TopMenu.f7.Visible = False
        TopMenu.f8.Visible = False
        TopMenu.f9.Visible = False
        TopMenu.f10.Visible = False
End Select

End Function

Sub RestoreAllActiveColor(ByVal EstNum As Long)

Dim NumCon As Integer

Select Case EstNum
    Case 1
        For NumCon = 0 To 21
            Est01.E11(NumCon).BackColor = &H404040        'GRIS (cara del boton)
        Next NumCon
    Case 2
        For NumCon = 0 To 21
            Est02.E21(NumCon).BackColor = &H404040        'GRIS (cara del boton)
        Next NumCon

End Select

End Sub

Sub RestoreActiveColor(ByVal index As Integer, ByVal EstNum As Long)

Select Case EstNum
    Case 1
        Est01.E11(index).BackColor = &H404040         'GRIS (cara del boton)
    Case 2
        Est02.E21(index).BackColor = &H404040        'GRIS (cara del boton)
End Select

End Sub

Sub ChangeActiveColor(ByVal index As Integer, ByVal EstNum As Long)

Select Case EstNum
    Case 1
        Est01.E11(index).BackColor = &HC0C000            'celeste
    Case 2
        Est02.E21(index).BackColor = &HC0C000            'celeste
End Select

End Sub

Sub SetDefControl(ByVal ContNum As Long)

Select Case ContNum
    Case 1
        For Contador = 0 To 21
            Est01.E11(Contador).Caption = ""
            Est01.E11(Contador).ToolTipText = "Duración:"
            Est12Data.N1(Contador).Caption = ""
            Est12Data.c1(Contador).Caption = ""
            Est12Data.D1(Contador).Caption = ""
            Est12Data.V1(Contador).Caption = ""
        Next Contador
    Case 2
        For Contador = 0 To 21
            Est02.E21(Contador).Caption = ""
            Est02.E21(Contador).ToolTipText = "Duración:"
            Est12Data.N2(Contador).Caption = ""
            Est12Data.c2(Contador).Caption = ""
            Est12Data.D2(Contador).Caption = ""
            Est12Data.V2(Contador).Caption = ""
        Next Contador
End Select

End Sub

Sub RestoreDisplay(ByVal DispNum As Long)

'setea los displays predeterminados a 0
'por defecto 00:00:00.

Select Case DispNum
    Case 1      'ESTACION 01
        'wave only
        Est01.E1p0.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1p1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1p2.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1p3.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1p4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1p5.Picture = TopMenu.SmallClip.GraphicCell(10)
        'time only
        Est01.E1p6.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1p7.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1p8.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1p9.Picture = TopMenu.SmallClip.GraphicCell(12) ':
        Est01.E1p10.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1p11.Picture = TopMenu.SmallClip.GraphicCell(10)
        'otro
        Est01.E1t0.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1t1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1t2.Picture = TopMenu.SmallClip.GraphicCell(12) ':
        Est01.E1t3.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1t4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1t5.Picture = TopMenu.SmallClip.GraphicCell(12) ':
        Est01.E1t6.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1t7.Picture = TopMenu.SmallClip.GraphicCell(10)
        'otro
        Est01.E1f0.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1f1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1f2.Picture = TopMenu.SmallClip.GraphicCell(12) ':
        Est01.E1f3.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1f4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1f5.Picture = TopMenu.SmallClip.GraphicCell(12) ':
        Est01.E1f6.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est01.E1f7.Picture = TopMenu.SmallClip.GraphicCell(10)
        
    Case 2      'ESTACION 02
        'wave only
        Est02.E2p0.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2p1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2p2.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2p3.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2p4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2p5.Picture = TopMenu.SmallClip.GraphicCell(10)
        'time only
        Est02.E2p6.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2p7.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2p8.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2p9.Picture = TopMenu.SmallClip.GraphicCell(12)
        Est02.E2p10.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2p11.Picture = TopMenu.SmallClip.GraphicCell(10)
        'otro
        Est02.E2t0.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2t1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2t2.Picture = TopMenu.SmallClip.GraphicCell(12) ':
        Est02.E2t3.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2t4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2t5.Picture = TopMenu.SmallClip.GraphicCell(12) ':
        Est02.E2t6.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2t7.Picture = TopMenu.SmallClip.GraphicCell(10)
        'otro
        Est02.E2f0.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2f1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2f2.Picture = TopMenu.SmallClip.GraphicCell(12) ':
        Est02.E2f3.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2f4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2f5.Picture = TopMenu.SmallClip.GraphicCell(12) ':
        Est02.E2f6.Picture = TopMenu.SmallClip.GraphicCell(10)
        Est02.E2f7.Picture = TopMenu.SmallClip.GraphicCell(10)

    Case 3      'TANDA 01
        'time only
        Tanda01.T1p0.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1p1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1p2.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1p3.Picture = TopMenu.SmallClip.GraphicCell(12)
        Tanda01.T1p4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1p5.Picture = TopMenu.SmallClip.GraphicCell(10)
    Case 4      'TANDA 02
        'Time only
        Tanda01.T1p6.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1p7.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1p8.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1p9.Picture = TopMenu.SmallClip.GraphicCell(12)
        Tanda01.T1p10.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1p11.Picture = TopMenu.SmallClip.GraphicCell(10)
        
    Case 5      'TOTAL time in TANDA 01 Y 02
        'Time only
        Tanda01.T1t1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1t2.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1t3.Picture = TopMenu.SmallClip.GraphicCell(12)
        Tanda01.T1t4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1t5.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1t6.Picture = TopMenu.SmallClip.GraphicCell(12)
        Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(10)
        
        Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1I3.Picture = TopMenu.SmallClip.GraphicCell(12)
        Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1I6.Picture = TopMenu.SmallClip.GraphicCell(12)
        Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(10)
        
        Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1F3.Picture = TopMenu.SmallClip.GraphicCell(12)
        Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1F6.Picture = TopMenu.SmallClip.GraphicCell(12)
        Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(10)
        Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(10)
        
    Case 6      'TIME DISPLAY FOR PH TIMING
        'time only
        TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(10)
        TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(10)
        TopMenu.Pht3.Picture = TopMenu.SmallClip.GraphicCell(12)
        TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(10)
        TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(10)
        TopMenu.Pht6.Picture = TopMenu.SmallClip.GraphicCell(12)
        TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(10)
        TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(10)
        
    Case 7      'TIME DISPLAY FOR PROGTANDAS ACTIVATE MODULE
        TopMenu.PrgT1.Picture = TopMenu.SmallClip.GraphicCell(10)
        TopMenu.PrgT2.Picture = TopMenu.SmallClip.GraphicCell(10)
        TopMenu.PrgT3.Picture = TopMenu.SmallClip.GraphicCell(12)
        TopMenu.PrgT4.Picture = TopMenu.SmallClip.GraphicCell(10)
        TopMenu.PrgT5.Picture = TopMenu.SmallClip.GraphicCell(10)
        TopMenu.PrgT6.Picture = TopMenu.SmallClip.GraphicCell(12)
        TopMenu.PrgT7.Picture = TopMenu.SmallClip.GraphicCell(10)
        TopMenu.PrgT8.Picture = TopMenu.SmallClip.GraphicCell(10)
        
    Case 8      'AUDIO PROP DISPLAY
        'file display
        AudioProp.T1t1.Picture = TopMenu.SmallClip.GraphicCell(10)
        AudioProp.T1t2.Picture = TopMenu.SmallClip.GraphicCell(10)
        AudioProp.T1t3.Picture = TopMenu.SmallClip.GraphicCell(12)
        AudioProp.T1t4.Picture = TopMenu.SmallClip.GraphicCell(10)
        AudioProp.T1t5.Picture = TopMenu.SmallClip.GraphicCell(10)
    Case 9      'AUDIO PROP DISPLAY
        'mix display
        AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(10)
        AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(10)
        AudioProp.T1M3.Picture = TopMenu.SmallClip.GraphicCell(12)
        AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(10)
        AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(10)
        
    Case 10     'TOTAL TIME IN PROGRAMACION DE TANDAS
        'Time only
        Prg01.p1t1.Picture = TopMenu.SmallClip.GraphicCell(10)
        Prg01.p1t2.Picture = TopMenu.SmallClip.GraphicCell(10)
        Prg01.p1t3.Picture = TopMenu.SmallClip.GraphicCell(12)
        Prg01.p1t4.Picture = TopMenu.SmallClip.GraphicCell(10)
        Prg01.p1t5.Picture = TopMenu.SmallClip.GraphicCell(10)
        Prg01.p1t6.Picture = TopMenu.SmallClip.GraphicCell(12)
        Prg01.p1t7.Picture = TopMenu.SmallClip.GraphicCell(10)
        Prg01.p1t8.Picture = TopMenu.SmallClip.GraphicCell(10)

    Case 11     'DISPLAY DEL CLIMA EN EL TOPMENU
        TopMenu.c1.Picture = TopMenu.TempClip.GraphicCell(10)
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(10)
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(10)
        TopMenu.c4.Picture = TopMenu.TempClip.GraphicCell(10)
        TopMenu.c5.Picture = TopMenu.TempClip.GraphicCell(10)
        TopMenu.c6.Picture = TopMenu.TempClip.GraphicCell(10)
        TopMenu.c7.Picture = TopMenu.TempClip.GraphicCell(10)
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(10)
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(10)
        TopMenu.c0.Picture = TopMenu.TempClip.GraphicCell(12)
End Select

End Sub

'---------------------------------------------------------------------
'Funcion para establecer la temperatura y humedad del dia en el reloj
'principal del topmenu.
'EN REVISION 11-03-24
'---------------------------------------------------------------------'
Sub TopClima(Wtemp As String, Whume As String)

'wdata debe estar formateado de la siguiente manera
'034°C-030% o 044°F-030%

'If Trim(Wtemp) Or Trim(Whume) = "--" Or Trim(Wtemp) Or Trim(Whume) = "er" Then GoTo NoClima

TempLen = Len(Trim(Wtemp))

Select Case TempLen
    Case 5  '034°C o 100°C o -34°C      -------------------------------------------------------------------
        If Left$(Wtemp, 1) = "0" Then
            DTemp(1) = "x"                  '0
            DTemp(2) = Mid$(Wtemp, 2, 1)    '3
            DTemp(3) = Mid$(Wtemp, 3, 1)    '4
            DTemp(4) = "°"
            DTemp(5) = Right$(Wtemp, 1)     'C o F
        Else
            DTemp(1) = Left$(Wtemp, 1)      '-
            DTemp(2) = Mid$(Wtemp, 2, 1)    '3
            DTemp(3) = Mid$(Wtemp, 3, 1)    '4
            DTemp(4) = "°"
            DTemp(5) = Right$(Wtemp, 1)     'C o F
        End If
    
    Case 4  '34°C o -4°C                -------------------------------------------------------------------
        If Left$(Wtemp, 1) = "-" Then
            DTemp(1) = "x"
            DTemp(2) = Left$(Wtemp, 1)      '-
            DTemp(3) = Mid$(Wtemp, 2, 1)    '4
            DTemp(4) = "°"
            DTemp(5) = Right$(Wtemp, 1)     'C o F
        Else
            DTemp(1) = "x"
            DTemp(2) = Left$(Wtemp, 1)      '3
            DTemp(3) = Mid$(Wtemp, 2, 1)    '4
            DTemp(4) = "°"
            DTemp(5) = Right$(Wtemp, 1)     'C o F
        End If
    
    Case 3  '4°C                        -------------------------------------------------------------------
        DTemp(1) = "x"
        DTemp(2) = "x"
        DTemp(3) = Left$(Wtemp, 1)      '4
        DTemp(4) = "°"
        DTemp(5) = Right$(Wtemp, 1)     'C o F

End Select

HumeLen = Len(Trim(Whume))

Select Case HumeLen
    Case 4  '100% o 085%                -------------------------------------------------------------------
        If Left$(Whume, 1) = "0" Then
            DHume(1) = "x"
            DHume(2) = Mid$(Whume, 2, 1)
            DHume(3) = Mid$(Whume, 3, 1)
        Else
            DHume(1) = Left$(Whume, 1)
            DHume(2) = Mid$(Whume, 2, 1)
            DHume(3) = Mid$(Whume, 3, 1)
        End If

    Case 3  '85% o 08%                 -------------------------------------------------------------------
        If Left$(Whume, 1) = "0" Then
            DHume(1) = "x"
            DHume(2) = "x"
            DHume(3) = Mid$(Whume, 2, 1)
        Else
            DHume(1) = "x"
            DHume(2) = Left$(Whume, 1)
            DHume(3) = Mid$(Whume, 2, 1)
        End If
    
    Case 2  '3%                       -------------------------------------------------------------------
        DHume(1) = "x"
        DHume(2) = "x"
        DHume(3) = Left$(Whume, 1)
        
End Select

'--------------------------------------------------------------------------------------------------------
'tempertura, display
If DTemp(1) = "-" Then TopMenu.c1.Picture = TopMenu.TempClip.GraphicCell(13)
If DTemp(1) = "x" Then TopMenu.c1.Picture = TopMenu.TempClip.GraphicCell(10)
If DTemp(1) = "0" Then TopMenu.c1.Picture = TopMenu.TempClip.GraphicCell(0)
If DTemp(1) = "1" Then TopMenu.c1.Picture = TopMenu.TempClip.GraphicCell(1)

Select Case DTemp(2)
    Case "x"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(10)
    Case "-"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(13)
    Case "0"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(0)
    Case "1"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(1)
    Case "2"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(2)
    Case "3"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(3)
    Case "4"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(4)
    Case "5"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(5)
    Case "6"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(6)
    Case "7"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(7)
    Case "8"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(8)
    Case "9"
        TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(9)
End Select
Select Case DTemp(3)
    Case "x"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(10)
    Case "-"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(13)
    Case "0"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(0)
    Case "1"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(1)
    Case "2"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(2)
    Case "3"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(3)
    Case "4"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(4)
    Case "5"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(5)
    Case "6"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(6)
    Case "7"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(7)
    Case "8"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(8)
    Case "9"
        TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(9)
End Select

TopMenu.c4.Picture = TopMenu.TempClip.GraphicCell(16) '°
If DTemp(5) = "C" Then
    TopMenu.c5.Picture = TopMenu.TempClip.GraphicCell(14)   'C
Else
    TopMenu.c5.Picture = TopMenu.TempClip.GraphicCell(15)   'F
End If

'----------------------------------------------------------------------------------
TopMenu.c6.Picture = TopMenu.TempClip.GraphicCell(10)   'separador
'----------------------------------------------------------------------------------

'humedad, display
If DHume(1) = "x" Then TopMenu.c7.Picture = TopMenu.TempClip.GraphicCell(10)
If DHume(1) = "0" Then TopMenu.c7.Picture = TopMenu.TempClip.GraphicCell(0)
If DHume(1) = "1" Then TopMenu.c7.Picture = TopMenu.TempClip.GraphicCell(1)
If DHume(1) = "2" Then TopMenu.c7.Picture = TopMenu.TempClip.GraphicCell(2)

Select Case DHume(2)
    Case "x"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(10)
    Case "-"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(13)
    Case "0"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(0)
    Case "1"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(1)
    Case "2"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(2)
    Case "3"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(3)
    Case "4"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(4)
    Case "5"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(5)
    Case "6"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(6)
    Case "7"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(7)
    Case "8"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(8)
    Case "9"
        TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(9)
End Select
Select Case DHume(3)
    Case "x"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(10)
    Case "-"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(13)
    Case "0"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(0)
    Case "1"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(1)
    Case "2"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(2)
    Case "3"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(3)
    Case "4"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(4)
    Case "5"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(5)
    Case "6"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(6)
    Case "7"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(7)
    Case "8"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(8)
    Case "9"
        TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(9)
End Select

TopMenu.c0.Picture = TopMenu.TempClip.GraphicCell(11)

Exit Sub

NoClima:
TopMenu.c1.Picture = TopMenu.TempClip.GraphicCell(10)
TopMenu.c2.Picture = TopMenu.TempClip.GraphicCell(10)
TopMenu.c3.Picture = TopMenu.TempClip.GraphicCell(10)
TopMenu.c4.Picture = TopMenu.TempClip.GraphicCell(10)
TopMenu.c5.Picture = TopMenu.TempClip.GraphicCell(10)
TopMenu.c6.Picture = TopMenu.TempClip.GraphicCell(10)
TopMenu.c7.Picture = TopMenu.TempClip.GraphicCell(10)
TopMenu.c8.Picture = TopMenu.TempClip.GraphicCell(10)
TopMenu.c9.Picture = TopMenu.TempClip.GraphicCell(10)
TopMenu.c0.Picture = TopMenu.TempClip.GraphicCell(12)

End Sub

'---------------------------------------------------------------------
'Funcion para establecer la fecha del dia en el reloj principal
'del topmenu.
'---------------------------------------------------------------------'
Sub TopDate(WNumber As String)

On Error GoTo NoClock
LMes = Left$(WNumber, 2)
LDia = Mid$(WNumber, 4, 2)
LAno = Right$(WNumber, 4)

'mes
NMes(1) = Left$(LMes, 1)
NMes(2) = Right$(LMes, 1)
'dia
NDia(1) = Left$(LDia, 1)
NDia(2) = Right$(LDia, 1)
'año
NAno(1) = Left$(LAno, 1)
NAno(2) = Mid$(LAno, 2, 1)
NAno(3) = Mid$(LAno, 3, 1)
NAno(4) = Right$(LAno, 1)

'display the day
If NDia(1) = "0" Then TopMenu.f1.Picture = TopMenu.BigClip.GraphicCell(0)
If NDia(1) = "1" Then TopMenu.f1.Picture = TopMenu.BigClip.GraphicCell(1)
If NDia(1) = "2" Then TopMenu.f1.Picture = TopMenu.BigClip.GraphicCell(2)
If NDia(1) = "3" Then TopMenu.f1.Picture = TopMenu.BigClip.GraphicCell(3)

If NDia(2) = "0" Then TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(0)
If NDia(2) = "1" Then TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(1)
If NDia(2) = "2" Then TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(2)
If NDia(2) = "3" Then TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(3)
If NDia(2) = "4" Then TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(4)
If NDia(2) = "5" Then TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(5)
If NDia(2) = "6" Then TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(6)
If NDia(2) = "7" Then TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(7)
If NDia(2) = "8" Then TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(8)
If NDia(2) = "9" Then TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(9)

TopMenu.f3.Picture = TopMenu.BigClip.GraphicCell(13)

'display the month
If NMes(1) = "0" Then TopMenu.f4.Picture = TopMenu.BigClip.GraphicCell(0)
If NMes(1) = "1" Then TopMenu.f4.Picture = TopMenu.BigClip.GraphicCell(1)

If NMes(2) = "0" Then TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(0)
If NMes(2) = "1" Then TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(1)
If NMes(2) = "2" Then TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(2)
If NMes(2) = "3" Then TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(3)
If NMes(2) = "4" Then TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(4)
If NMes(2) = "5" Then TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(5)
If NMes(2) = "6" Then TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(6)
If NMes(2) = "7" Then TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(7)
If NMes(2) = "8" Then TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(8)
If NMes(2) = "9" Then TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(9)

TopMenu.f6.Picture = TopMenu.BigClip.GraphicCell(13)

'display the year
If NAno(1) = "1" Then TopMenu.f7.Picture = TopMenu.BigClip.GraphicCell(1)
If NAno(1) = "2" Then TopMenu.f7.Picture = TopMenu.BigClip.GraphicCell(2)
If NAno(1) = "3" Then TopMenu.f7.Picture = TopMenu.BigClip.GraphicCell(3)

If NAno(2) = "0" Then TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(0)
If NAno(2) = "1" Then TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(1)
If NAno(2) = "2" Then TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(2)
If NAno(2) = "3" Then TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(3)
If NAno(2) = "4" Then TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(4)
If NAno(2) = "5" Then TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(5)
If NAno(2) = "6" Then TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(6)
If NAno(2) = "7" Then TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(7)
If NAno(2) = "8" Then TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(8)
If NAno(2) = "9" Then TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(9)

If NAno(3) = "0" Then TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(0)
If NAno(3) = "1" Then TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(1)
If NAno(3) = "2" Then TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(2)
If NAno(3) = "3" Then TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(3)
If NAno(3) = "4" Then TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(4)
If NAno(3) = "5" Then TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(5)
If NAno(3) = "6" Then TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(6)
If NAno(3) = "7" Then TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(7)
If NAno(3) = "8" Then TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(8)
If NAno(3) = "9" Then TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(9)

If NAno(4) = "0" Then TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(0)
If NAno(4) = "1" Then TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(1)
If NAno(4) = "2" Then TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(2)
If NAno(4) = "3" Then TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(3)
If NAno(4) = "4" Then TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(4)
If NAno(4) = "5" Then TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(5)
If NAno(4) = "6" Then TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(6)
If NAno(4) = "7" Then TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(7)
If NAno(4) = "8" Then TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(8)
If NAno(4) = "9" Then TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(9)

Exit Sub

NoClock:
'display non day
TopMenu.f1.Picture = TopMenu.BigClip.GraphicCell(0)
TopMenu.f2.Picture = TopMenu.BigClip.GraphicCell(0)

TopMenu.f3.Picture = TopMenu.BigClip.GraphicCell(13)

'display non month
TopMenu.f4.Picture = TopMenu.BigClip.GraphicCell(0)
TopMenu.f5.Picture = TopMenu.BigClip.GraphicCell(0)

TopMenu.f6.Picture = TopMenu.BigClip.GraphicCell(13)

'display non year
TopMenu.f7.Picture = TopMenu.BigClip.GraphicCell(0)
TopMenu.f8.Picture = TopMenu.BigClip.GraphicCell(0)
TopMenu.f9.Picture = TopMenu.BigClip.GraphicCell(0)
TopMenu.f10.Picture = TopMenu.BigClip.GraphicCell(0)
End Sub

'---------------------------------------------------------------------
'Funcion para establecer la hora del dia en el reloj principal
'del topmenu.
'---------------------------------------------------------------------
Sub TopClock(WNumber As String)

On Error GoTo NoClock
LHora = Left$(WNumber, 2)
LMinutos = Mid$(WNumber, 4, 2)
LSegundos = Right$(WNumber, 2)

'hora
NHora(1) = Left$(LHora, 1)
NHora(2) = Right$(LHora, 1)
'minutos
NMinutos(1) = Left$(LMinutos, 1)
NMinutos(2) = Right$(LMinutos, 1)
'segundos
NSegundos(1) = Left$(LSegundos, 1)
NSegundos(2) = Right$(LSegundos, 1)

'display the time
If NHora(1) = "0" Then TopMenu.t1.Picture = TopMenu.BigClip.GraphicCell(0)
If NHora(1) = "1" Then TopMenu.t1.Picture = TopMenu.BigClip.GraphicCell(1)
If NHora(1) = "2" Then TopMenu.t1.Picture = TopMenu.BigClip.GraphicCell(2)

If NHora(2) = "0" Then TopMenu.t2.Picture = TopMenu.BigClip.GraphicCell(0)
If NHora(2) = "1" Then TopMenu.t2.Picture = TopMenu.BigClip.GraphicCell(1)
If NHora(2) = "2" Then TopMenu.t2.Picture = TopMenu.BigClip.GraphicCell(2)
If NHora(2) = "3" Then TopMenu.t2.Picture = TopMenu.BigClip.GraphicCell(3)
If NHora(2) = "4" Then TopMenu.t2.Picture = TopMenu.BigClip.GraphicCell(4)
If NHora(2) = "5" Then TopMenu.t2.Picture = TopMenu.BigClip.GraphicCell(5)
If NHora(2) = "6" Then TopMenu.t2.Picture = TopMenu.BigClip.GraphicCell(6)
If NHora(2) = "7" Then TopMenu.t2.Picture = TopMenu.BigClip.GraphicCell(7)
If NHora(2) = "8" Then TopMenu.t2.Picture = TopMenu.BigClip.GraphicCell(8)
If NHora(2) = "9" Then TopMenu.t2.Picture = TopMenu.BigClip.GraphicCell(9)

TopMenu.t3.Picture = TopMenu.BigClip.GraphicCell(11)

If NMinutos(1) = "0" Then TopMenu.t4.Picture = TopMenu.BigClip.GraphicCell(0)
If NMinutos(1) = "1" Then TopMenu.t4.Picture = TopMenu.BigClip.GraphicCell(1)
If NMinutos(1) = "2" Then TopMenu.t4.Picture = TopMenu.BigClip.GraphicCell(2)
If NMinutos(1) = "3" Then TopMenu.t4.Picture = TopMenu.BigClip.GraphicCell(3)
If NMinutos(1) = "4" Then TopMenu.t4.Picture = TopMenu.BigClip.GraphicCell(4)
If NMinutos(1) = "5" Then TopMenu.t4.Picture = TopMenu.BigClip.GraphicCell(5)

If NMinutos(2) = "0" Then TopMenu.t5.Picture = TopMenu.BigClip.GraphicCell(0)
If NMinutos(2) = "1" Then TopMenu.t5.Picture = TopMenu.BigClip.GraphicCell(1)
If NMinutos(2) = "2" Then TopMenu.t5.Picture = TopMenu.BigClip.GraphicCell(2)
If NMinutos(2) = "3" Then TopMenu.t5.Picture = TopMenu.BigClip.GraphicCell(3)
If NMinutos(2) = "4" Then TopMenu.t5.Picture = TopMenu.BigClip.GraphicCell(4)
If NMinutos(2) = "5" Then TopMenu.t5.Picture = TopMenu.BigClip.GraphicCell(5)
If NMinutos(2) = "6" Then TopMenu.t5.Picture = TopMenu.BigClip.GraphicCell(6)
If NMinutos(2) = "7" Then TopMenu.t5.Picture = TopMenu.BigClip.GraphicCell(7)
If NMinutos(2) = "8" Then TopMenu.t5.Picture = TopMenu.BigClip.GraphicCell(8)
If NMinutos(2) = "9" Then TopMenu.t5.Picture = TopMenu.BigClip.GraphicCell(9)

TopMenu.t6.Picture = TopMenu.BigClip.GraphicCell(11)

If NSegundos(1) = "0" Then TopMenu.t7.Picture = TopMenu.BigClip.GraphicCell(0)
If NSegundos(1) = "1" Then TopMenu.t7.Picture = TopMenu.BigClip.GraphicCell(1)
If NSegundos(1) = "2" Then TopMenu.t7.Picture = TopMenu.BigClip.GraphicCell(2)
If NSegundos(1) = "3" Then TopMenu.t7.Picture = TopMenu.BigClip.GraphicCell(3)
If NSegundos(1) = "4" Then TopMenu.t7.Picture = TopMenu.BigClip.GraphicCell(4)
If NSegundos(1) = "5" Then TopMenu.t7.Picture = TopMenu.BigClip.GraphicCell(5)

If NSegundos(2) = "0" Then TopMenu.t8.Picture = TopMenu.BigClip.GraphicCell(0)
If NSegundos(2) = "1" Then TopMenu.t8.Picture = TopMenu.BigClip.GraphicCell(1)
If NSegundos(2) = "2" Then TopMenu.t8.Picture = TopMenu.BigClip.GraphicCell(2)
If NSegundos(2) = "3" Then TopMenu.t8.Picture = TopMenu.BigClip.GraphicCell(3)
If NSegundos(2) = "4" Then TopMenu.t8.Picture = TopMenu.BigClip.GraphicCell(4)
If NSegundos(2) = "5" Then TopMenu.t8.Picture = TopMenu.BigClip.GraphicCell(5)
If NSegundos(2) = "6" Then TopMenu.t8.Picture = TopMenu.BigClip.GraphicCell(6)
If NSegundos(2) = "7" Then TopMenu.t8.Picture = TopMenu.BigClip.GraphicCell(7)
If NSegundos(2) = "8" Then TopMenu.t8.Picture = TopMenu.BigClip.GraphicCell(8)
If NSegundos(2) = "9" Then TopMenu.t8.Picture = TopMenu.BigClip.GraphicCell(9)

Exit Sub

NoClock:
End Sub

