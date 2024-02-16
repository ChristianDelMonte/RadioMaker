Attribute VB_Name = "MNUm"
'********************* RM100 *********************
'     RADIO MAKER DIGITAL DISPLAYS MODULE
'COPYRIGHT (C) 1987-2002 ONLY development inc.
'Christian A. Del Monte
'*************************************************

Option Explicit

Function FormatSegs(WUnfTime As String) As String

'Funcion para formatear el tiempo devuelto por el RMM control
'a segundos.
'ej: 120,38478378437 =>> 120

Dim UnfTime, ForTime
Dim GetComma As String

'chequeos de procesamiento
GetComma = Left$(WUnfTime, 1)   '=.0000000
If GetComma = "." Then
    ForTime = 0
Else
    GetComma = Mid$(WUnfTime, 2, 1) '=0,00000000
    If GetComma = "." Then
        ForTime = Left$(WUnfTime, 1)
    Else
        GetComma = Mid$(WUnfTime, 3, 1) '=00,0000000
        If GetComma = "." Then
            ForTime = Left$(WUnfTime, 2)
        Else
            GetComma = Mid$(WUnfTime, 4, 1) '=000,0000000
            If GetComma = "." Then
                ForTime = Left$(WUnfTime, 3)
            Else
                GetComma = Mid$(WUnfTime, 5, 1) '=0000,0000000
                If GetComma = "." Then
                    ForTime = Left$(WUnfTime, 4)
                Else
                    ForTime = Left$(WUnfTime, 4)
                End If
            End If
        End If
    End If
End If

FormatSegs = ForTime

End Function
Public Sub OrderTndTime(WOrderMode As String)

Dim nIndex As Integer
Dim ItmX As ListItem
Dim DataA(0 To 9) As String, DataB(0 To 9) As String, DataKa As String
Dim HoldIT As Integer, nCount As Integer, i As Integer
Dim DTema As String, HTema As String, HNew As String
Dim Time1 As Double, Time2 As Double, RTime As Double
Dim TMint As Integer

Select Case WOrderMode
    
    '///// RESETEA A LA HORA ACTUAL TODA LA LISTA DESDE EL COMIENZO
    Case "ResetAll"
        nIndex = 1
        Tanda01.T1View.ListItems.Item(nIndex).Selected = True
        'extraemos los datos del tema seleccionado
        DataA(0) = Tanda01.T1View.SelectedItem.text    'file & path
        'DataA(1) = Tanda01.T1View.SelectedItem.ListSubItems(1).Text     'filetype
        'DataA(2) = Tanda01.T1View.SelectedItem.ListSubItems(2).Text     'filename
        'DataA(3) = Tanda01.T1View.SelectedItem.ListSubItems(3).Text     'duracion
        'DataA(4) = Tanda01.T1View.SelectedItem.ListSubItems(4).Text     'hora de lanz
        'MIXER FILE
        'DataA(5) = Tanda01.T1View.SelectedItem.ListSubItems(5).Text     'file & path
        'DataA(6) = Tanda01.T1View.SelectedItem.ListSubItems(6).Text     'filetype
        'DataA(7) = Tanda01.T1View.SelectedItem.ListSubItems(7).Text     'filename
        'DataA(8) = Tanda01.T1View.SelectedItem.ListSubItems(8).Text     'duracion
        'DataA(9) = Tanda01.T1View.SelectedItem.ListSubItems(9).Text     'hora de lanz
        'eliminamos el stream de la lista
        Tanda01.T1View.ListItems.Remove (nIndex)
        DataKa = "r" & Str(nIndex)
        'ponemos los nuevos datos
        'Set ItmX = Tanda01.T1View.ListItems.Add(nIndex, DataKa, DataA(0)) 'path & file
        'ItmX.SubItems(1) = DataA(1)
        'ItmX.SubItems(2) = DataA(2)
        'ItmX.SubItems(3) = DataA(3)
        'ItmX.SubItems(4) = Trim(time$)
        'ItmX.SubItems(5) = DataA(5)
        'ItmX.SubItems(6) = DataA(6)
        'ItmX.SubItems(7) = DataA(7)
        'ItmX.SubItems(8) = DataA(8)
        'ItmX.SubItems(9) = DataA(9)
        'seleccionamos el primer stream nuevamente
        Tanda01.T1View.ListItems.Item(nIndex).Selected = True
        'actualizamos los datos de los demas temas
        Call OrderTndTime("Selected")
        Exit Sub
    
    '///// RESETEA A LA HORA ACTUAL DE LO SELECCIONADO PARA ABAJO
    Case "ResetSelected"
        nIndex = Tanda01.T1View.SelectedItem.index   'numero de index
        Tanda01.T1View.ListItems.Item(nIndex).Selected = True
        'extraemos los datos del tema seleccionado
        'DataA(0) = Tanda01.T1View.SelectedItem.Text    'file & path
        'DataA(1) = Tanda01.T1View.SelectedItem.ListSubItems(1).Text     'filetype
        'DataA(2) = Tanda01.T1View.SelectedItem.ListSubItems(2).Text     'filename
        'DataA(3) = Tanda01.T1View.SelectedItem.ListSubItems(3).Text     'duracion
        'DataA(4) = Tanda01.T1View.SelectedItem.ListSubItems(4).Text     'hora de lanz
        'MIXER FILE
        'DataA(5) = Tanda01.T1View.SelectedItem.ListSubItems(5).Text     'file & path
        'DataA(6) = Tanda01.T1View.SelectedItem.ListSubItems(6).Text     'filetype
        'DataA(7) = Tanda01.T1View.SelectedItem.ListSubItems(7).Text     'filename
        'DataA(8) = Tanda01.T1View.SelectedItem.ListSubItems(8).Text     'duracion
        'DataA(9) = Tanda01.T1View.SelectedItem.ListSubItems(9).Text     'hora de lanz
        'eliminamos el stream de la lista
        Tanda01.T1View.ListItems.Remove (nIndex)
        DataKa = "r" & Str(nIndex)
        'ponemos los nuevos datos
        'Set ItmX = Tanda01.T1View.ListItems.Add(nIndex, DataKa, DataA(0)) 'path & file
        'ItmX.SubItems(1) = DataA(1)
        'ItmX.SubItems(2) = DataA(2)
        'ItmX.SubItems(3) = DataA(3)
        'ItmX.SubItems(4) = Trim(time$)
        'ItmX.SubItems(5) = DataA(5)
        'ItmX.SubItems(6) = DataA(6)
        'ItmX.SubItems(7) = DataA(7)
        'ItmX.SubItems(8) = DataA(8)
        'ItmX.SubItems(9) = DataA(9)
        'seleccionamos el primer stream nuevamente
        Tanda01.T1View.ListItems.Item(nIndex).Selected = True
        'actualizamos los datos de los demas temas
        Call OrderTndTime("Selected")
        Exit Sub
    
    '///// REORDENA LO SELECCIONADO Y HACIA ABAJO
    Case "Selected"
        nIndex = Tanda01.T1View.SelectedItem.index   'numero de index
        HoldIT = Tanda01.T1View.SelectedItem.index   'numero de index
        'NIndex = CInt(Trim(Lidx.Caption))           'numero de index
        nCount = Tanda01.T1View.ListItems.count      'total de streams en la lista
        For i = nIndex To nCount    'start from stream selected in the list
            If nIndex >= nCount Then
                Exit For
            End If
            Tanda01.T1View.ListItems.Item(i).Selected = True
            'extraemos los datos del tema seleccionado
            DataA(0) = Tanda01.T1View.SelectedItem.text    'file & path
            'DataA(1) = Tanda01.T1View.SelectedItem.ListSubItems(1).Text     'filetype
            'DataA(2) = Tanda01.T1View.SelectedItem.ListSubItems(2).Text     'filename
            'DataA(3) = Tanda01.T1View.SelectedItem.ListSubItems(3).Text     'duracion
            'DataA(4) = Tanda01.T1View.SelectedItem.ListSubItems(4).Text     'hora de lanz
            'MIXER FILE
            'DataA(5) = Tanda01.T1View.SelectedItem.ListSubItems(5).Text     'file & path
            'DataA(6) = Tanda01.T1View.SelectedItem.ListSubItems(6).Text     'filetype
            'DataA(7) = Tanda01.T1View.SelectedItem.ListSubItems(7).Text     'filename
            'DataA(8) = Tanda01.T1View.SelectedItem.ListSubItems(8).Text     'duracion
            'DataA(9) = Tanda01.T1View.SelectedItem.ListSubItems(9).Text     'hora de lanz
            'duracion del tema seleccionado
            'DTema = Trim(Tanda01.T1View.SelectedItem.ListSubItems(3).Text)    'duracion
            'hora de lanzamiento del tema
            'HTema = Trim(Tanda01.T1View.SelectedItem.ListSubItems(4).Text)    'hora de lanz
            'procesamos los datos
            Time1 = ConvMinToSec(DTema) 'duracion
            Time2 = ConvMinToSec(HTema) 'hora de lanzamiento
            'extraemos el tiempo de mixado intermedio
            TMint = CInt(Trim(Tanda01.Intr.text))
            RTime = Time2 + Time1 'sumamos los tiempos
            RTime = (RTime - TMint) + 1
            HNew = ConvSecToMin(RTime)
            'seleccionamos el siguiente stream de la lista
            nIndex = i + 1
            Tanda01.T1View.ListItems.Item(nIndex).Selected = True
            'extraemos los datos del tema seleccionado
            DataB(0) = Tanda01.T1View.SelectedItem.text    'file & path
            'DataB(1) = Tanda01.T1View.SelectedItem.ListSubItems(1).Text     'filetype
            'DataB(2) = Tanda01.T1View.SelectedItem.ListSubItems(2).Text     'filename
            'DataB(3) = Tanda01.T1View.SelectedItem.ListSubItems(3).Text     'duracion
            'DataB(4) = Tanda01.T1View.SelectedItem.ListSubItems(4).Text     'hora de lanz
            'MIXER FILE
            'DataB(5) = Tanda01.T1View.SelectedItem.ListSubItems(5).Text     'file & path
            'DataB(6) = Tanda01.T1View.SelectedItem.ListSubItems(6).Text     'filetype
            'DataB(7) = Tanda01.T1View.SelectedItem.ListSubItems(7).Text     'filename
            'DataB(8) = Tanda01.T1View.SelectedItem.ListSubItems(8).Text     'duracion
            'DataB(9) = Tanda01.T1View.SelectedItem.ListSubItems(9).Text     'hora de lanz
            'eliminamos el stream de la lista
            Tanda01.T1View.ListItems.Remove (nIndex)
            DataKa = "r" & Str(nIndex)
            'ponemos los nuevos datos
            'Set ItmX = Tanda01.T1View.ListItems.Add(nIndex, DataKa, DataB(0)) 'path & file
            'ItmX.SubItems(1) = DataB(1)
            'ItmX.SubItems(2) = DataB(2)
            'ItmX.SubItems(3) = DataB(3)
            'ItmX.SubItems(4) = Trim(HNew)
            'ItmX.SubItems(5) = DataB(5)
            'ItmX.SubItems(6) = DataB(6)
            'ItmX.SubItems(7) = DataB(7)
            'ItmX.SubItems(8) = DataB(8)
            'ItmX.SubItems(9) = DataB(9)
        Next i
End Select

'//// seleccionamos el primer stream nuevamente
Tanda01.T1View.ListItems.Item(HoldIT).Selected = True

End Sub

Public Sub SetAudioTime(WDisplay As String, DNum As String)

'SOLO PARA propiedades de audio en TANDA
'el dnum debe estar en minutos = 00:00

Dim N1, N2, N3, N4, N5, N6
Dim LenNum

LenNum = Len(DNum)

If LenNum > 5 Then
    Exit Sub
End If

Select Case WDisplay
    Case 1, "1"
        If LenNum = 5 Then
            DNum = LTrim(RTrim(DNum))
            N1 = Left$(DNum, 1)
            N2 = Mid$(DNum, 2, 1)
            N3 = Mid$(DNum, 4, 1)
            N4 = Right$(DNum, 1)
            
            If N1 = 0 Then AudioProp.T1T1.Picture = TopMenu.SmallClip.GraphicCell(0)
            If N1 = 1 Then AudioProp.T1T1.Picture = TopMenu.SmallClip.GraphicCell(1)
            If N1 = 2 Then AudioProp.T1T1.Picture = TopMenu.SmallClip.GraphicCell(2)
            If N1 = 3 Then AudioProp.T1T1.Picture = TopMenu.SmallClip.GraphicCell(3)
            If N1 = 4 Then AudioProp.T1T1.Picture = TopMenu.SmallClip.GraphicCell(4)
            If N1 = 5 Then AudioProp.T1T1.Picture = TopMenu.SmallClip.GraphicCell(5)
            If N1 = 6 Then AudioProp.T1T1.Picture = TopMenu.SmallClip.GraphicCell(6)
            If N1 = 7 Then AudioProp.T1T1.Picture = TopMenu.SmallClip.GraphicCell(7)
            If N1 = 8 Then AudioProp.T1T1.Picture = TopMenu.SmallClip.GraphicCell(8)
            If N1 = 9 Then AudioProp.T1T1.Picture = TopMenu.SmallClip.GraphicCell(9)
            
            If N2 = 0 Then AudioProp.T1T2.Picture = TopMenu.SmallClip.GraphicCell(0)
            If N2 = 1 Then AudioProp.T1T2.Picture = TopMenu.SmallClip.GraphicCell(1)
            If N2 = 2 Then AudioProp.T1T2.Picture = TopMenu.SmallClip.GraphicCell(2)
            If N2 = 3 Then AudioProp.T1T2.Picture = TopMenu.SmallClip.GraphicCell(3)
            If N2 = 4 Then AudioProp.T1T2.Picture = TopMenu.SmallClip.GraphicCell(4)
            If N2 = 5 Then AudioProp.T1T2.Picture = TopMenu.SmallClip.GraphicCell(5)
            If N2 = 6 Then AudioProp.T1T2.Picture = TopMenu.SmallClip.GraphicCell(6)
            If N2 = 7 Then AudioProp.T1T2.Picture = TopMenu.SmallClip.GraphicCell(7)
            If N2 = 8 Then AudioProp.T1T2.Picture = TopMenu.SmallClip.GraphicCell(8)
            If N2 = 9 Then AudioProp.T1T2.Picture = TopMenu.SmallClip.GraphicCell(9)

            AudioProp.T1T3.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
        
            If N3 = 0 Then AudioProp.T1T4.Picture = TopMenu.SmallClip.GraphicCell(0)
            If N3 = 1 Then AudioProp.T1T4.Picture = TopMenu.SmallClip.GraphicCell(1)
            If N3 = 2 Then AudioProp.T1T4.Picture = TopMenu.SmallClip.GraphicCell(2)
            If N3 = 3 Then AudioProp.T1T4.Picture = TopMenu.SmallClip.GraphicCell(3)
            If N3 = 4 Then AudioProp.T1T4.Picture = TopMenu.SmallClip.GraphicCell(4)
            If N3 = 5 Then AudioProp.T1T4.Picture = TopMenu.SmallClip.GraphicCell(5)
            If N3 = 6 Then AudioProp.T1T4.Picture = TopMenu.SmallClip.GraphicCell(6)
            If N3 = 7 Then AudioProp.T1T4.Picture = TopMenu.SmallClip.GraphicCell(7)
            If N3 = 8 Then AudioProp.T1T4.Picture = TopMenu.SmallClip.GraphicCell(8)
            If N3 = 9 Then AudioProp.T1T4.Picture = TopMenu.SmallClip.GraphicCell(9)
    
            If N4 = 0 Then AudioProp.T1T5.Picture = TopMenu.SmallClip.GraphicCell(0)
            If N4 = 1 Then AudioProp.T1T5.Picture = TopMenu.SmallClip.GraphicCell(1)
            If N4 = 2 Then AudioProp.T1T5.Picture = TopMenu.SmallClip.GraphicCell(2)
            If N4 = 3 Then AudioProp.T1T5.Picture = TopMenu.SmallClip.GraphicCell(3)
            If N4 = 4 Then AudioProp.T1T5.Picture = TopMenu.SmallClip.GraphicCell(4)
            If N4 = 5 Then AudioProp.T1T5.Picture = TopMenu.SmallClip.GraphicCell(5)
            If N4 = 6 Then AudioProp.T1T5.Picture = TopMenu.SmallClip.GraphicCell(6)
            If N4 = 7 Then AudioProp.T1T5.Picture = TopMenu.SmallClip.GraphicCell(7)
            If N4 = 8 Then AudioProp.T1T5.Picture = TopMenu.SmallClip.GraphicCell(8)
            If N4 = 9 Then AudioProp.T1T5.Picture = TopMenu.SmallClip.GraphicCell(9)
        End If
    Case 2, "2"
        If LenNum = 5 Then
            DNum = LTrim(RTrim(DNum))
            N1 = Left$(DNum, 1)
            N2 = Mid$(DNum, 2, 1)
            N3 = Mid$(DNum, 4, 1)
            N4 = Right$(DNum, 1)
            
            If N1 = 0 Then AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(0)
            If N1 = 1 Then AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(1)
            If N1 = 2 Then AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(2)
            If N1 = 3 Then AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(3)
            If N1 = 4 Then AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(4)
            If N1 = 5 Then AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(5)
            If N1 = 6 Then AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(6)
            If N1 = 7 Then AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(7)
            If N1 = 8 Then AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(8)
            If N1 = 9 Then AudioProp.T1M1.Picture = TopMenu.SmallClip.GraphicCell(9)
            
            If N2 = 0 Then AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(0)
            If N2 = 1 Then AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(1)
            If N2 = 2 Then AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(2)
            If N2 = 3 Then AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(3)
            If N2 = 4 Then AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(4)
            If N2 = 5 Then AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(5)
            If N2 = 6 Then AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(6)
            If N2 = 7 Then AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(7)
            If N2 = 8 Then AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(8)
            If N2 = 9 Then AudioProp.T1M2.Picture = TopMenu.SmallClip.GraphicCell(9)

            AudioProp.T1M3.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
        
            If N3 = 0 Then AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(0)
            If N3 = 1 Then AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(1)
            If N3 = 2 Then AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(2)
            If N3 = 3 Then AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(3)
            If N3 = 4 Then AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(4)
            If N3 = 5 Then AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(5)
            If N3 = 6 Then AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(6)
            If N3 = 7 Then AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(7)
            If N3 = 8 Then AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(8)
            If N3 = 9 Then AudioProp.T1M4.Picture = TopMenu.SmallClip.GraphicCell(9)

            If N4 = 0 Then AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(0)
            If N4 = 1 Then AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(1)
            If N4 = 2 Then AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(2)
            If N4 = 3 Then AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(3)
            If N4 = 4 Then AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(4)
            If N4 = 5 Then AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(5)
            If N4 = 6 Then AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(6)
            If N4 = 7 Then AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(7)
            If N4 = 8 Then AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(8)
            If N4 = 9 Then AudioProp.T1M5.Picture = TopMenu.SmallClip.GraphicCell(9)
        End If
End Select

End Sub

Public Sub SetDigClock(WTime As String, WEstNum As String, WType As String)

'formatea el tiempo de salida de los temas
'para mostrarlos en el display
'WTime debe ser: 00:00:00 y el resultado es: 00:00 or -00:00
'WEstNum debe ser: 1 or 2 or other is there is more (numero de est. donde display)
'WType debe ser: Normal or Restante

Dim Minutos As String
Dim M1, M2 As Integer
Dim Segundos As String
Dim s1, s2 As Integer
Dim Resultado As String

'formateamos el tiempo y separamos los minutos de los segundos
Resultado = Trim(Right$(WTime, 5))
Minutos = Left$(Resultado, 2)
Segundos = Right$(Resultado, 2)
M1 = CInt(Left$(Minutos, 1)): M2 = CInt(Right$(Minutos, 1))
s1 = CInt(Left$(Segundos, 1)): s2 = CInt(Right$(Segundos, 1))

'seteamos el display con los numeros correspondientes
Select Case WEstNum
    Case 1, "1"     '****************** ESTACION 01
        Select Case WType
            Case "Normal"
                Est01.E1p6.Picture = TopMenu.SmallClip.GraphicCell(10) '= nada
                Est01.E1p7.Picture = TopMenu.SmallClip.GraphicCell(M1)
                Est01.E1p8.Picture = TopMenu.SmallClip.GraphicCell(M2)
                Est01.E1p9.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
                Est01.E1p10.Picture = TopMenu.SmallClip.GraphicCell(s1)
                Est01.E1p11.Picture = TopMenu.SmallClip.GraphicCell(s2)
            Case "Restante"
                Est01.E1p6.Picture = TopMenu.SmallClip.GraphicCell(13) '= signo menos
                Est01.E1p7.Picture = TopMenu.SmallClip.GraphicCell(M1)
                Est01.E1p8.Picture = TopMenu.SmallClip.GraphicCell(M2)
                Est01.E1p9.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
                Est01.E1p10.Picture = TopMenu.SmallClip.GraphicCell(s1)
                Est01.E1p11.Picture = TopMenu.SmallClip.GraphicCell(s2)
        End Select
    Case 2, "2"     '****************** ESTACION 02
        Select Case WType
            Case "Normal"
                Est02.E2p6.Picture = TopMenu.SmallClip.GraphicCell(10) '= nada
                Est02.E2p7.Picture = TopMenu.SmallClip.GraphicCell(M1)
                Est02.E2p8.Picture = TopMenu.SmallClip.GraphicCell(M2)
                Est02.E2p9.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
                Est02.E2p10.Picture = TopMenu.SmallClip.GraphicCell(s1)
                Est02.E2p11.Picture = TopMenu.SmallClip.GraphicCell(s2)
            Case "Restante"
                Est02.E2p6.Picture = TopMenu.SmallClip.GraphicCell(13) '= signo menos
                Est02.E2p7.Picture = TopMenu.SmallClip.GraphicCell(M1)
                Est02.E2p8.Picture = TopMenu.SmallClip.GraphicCell(M2)
                Est02.E2p9.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
                Est02.E2p10.Picture = TopMenu.SmallClip.GraphicCell(s1)
                Est02.E2p11.Picture = TopMenu.SmallClip.GraphicCell(s2)
        End Select
    Case 3, "3"     '****************** TANDA 01
        Select Case WType
            Case "Normal"
                Tanda01.T1p0.Picture = TopMenu.SmallClip.GraphicCell(10) '= nada
                Tanda01.T1p1.Picture = TopMenu.SmallClip.GraphicCell(M1)
                Tanda01.T1p2.Picture = TopMenu.SmallClip.GraphicCell(M2)
                Tanda01.T1p3.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
                Tanda01.T1p4.Picture = TopMenu.SmallClip.GraphicCell(s1)
                Tanda01.T1p5.Picture = TopMenu.SmallClip.GraphicCell(s2)
            Case "Restante"
                Tanda01.T1p0.Picture = TopMenu.SmallClip.GraphicCell(13) '= signo menos
                Tanda01.T1p1.Picture = TopMenu.SmallClip.GraphicCell(M1)
                Tanda01.T1p2.Picture = TopMenu.SmallClip.GraphicCell(M2)
                Tanda01.T1p3.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
                Tanda01.T1p4.Picture = TopMenu.SmallClip.GraphicCell(s1)
                Tanda01.T1p5.Picture = TopMenu.SmallClip.GraphicCell(s2)
        End Select
    Case 4, "4"     '****************** TANDA 02
        Select Case WType
            Case "Normal"
                Tanda01.T1p6.Picture = TopMenu.SmallClip.GraphicCell(10)  '= nada
                Tanda01.T1p7.Picture = TopMenu.SmallClip.GraphicCell(M1)
                Tanda01.T1p8.Picture = TopMenu.SmallClip.GraphicCell(M2)
                Tanda01.T1p9.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
                Tanda01.T1p10.Picture = TopMenu.SmallClip.GraphicCell(s1)
                Tanda01.T1p11.Picture = TopMenu.SmallClip.GraphicCell(s2)
            Case "Restante"
                Tanda01.T1p6.Picture = TopMenu.SmallClip.GraphicCell(13)  '= signo menos
                Tanda01.T1p7.Picture = TopMenu.SmallClip.GraphicCell(M1)
                Tanda01.T1p8.Picture = TopMenu.SmallClip.GraphicCell(M2)
                Tanda01.T1p9.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
                Tanda01.T1p10.Picture = TopMenu.SmallClip.GraphicCell(s1)
                Tanda01.T1p11.Picture = TopMenu.SmallClip.GraphicCell(s2)
        End Select
End Select

End Sub

Public Sub SetDigNum(Wnum As String, WEstNum As String, WType As String)

'formatea el tiempo de salida de los temas
'para mostrarlos en el display
'WTime debe ser de hasta: 00000000 (8bytes) y el resultado es: -00000 or 00000 (5bytes)
'WEstNum debe ser: 1 or 2 or other is there is more (numero de est. donde display)
'WType debe ser: Normal or Restante

Dim NumLen As Long
Dim NumRes As String
Dim N1, N2, N3, N4, N5 As Integer

'formateamos el tiempo
Wnum = Trim(Wnum)
NumLen = Len(Wnum)
Select Case NumLen
    Case Is <= 3
        N1 = 10: N2 = 10: N3 = 10: N4 = 10: N5 = 10
    Case 4
        Wnum = Left$(Wnum, 1)
        N1 = 10: N2 = 10: N3 = 10: N4 = 10
        N5 = CInt(Wnum)
    Case 5
        Wnum = Left$(Wnum, 2)
        N1 = 10: N2 = 10: N3 = 10
        N4 = CInt(Left$(Wnum, 1))
        N5 = CInt(Right$(Wnum, 1))
    Case 6
        Wnum = Left$(Wnum, 3)
        N1 = 10: N2 = 10
        N3 = CInt(Left$(Wnum, 1))
        N4 = CInt(Mid$(Wnum, 2, 1))
        N5 = CInt(Right$(Wnum, 1))
    Case 7
        Wnum = Left$(Wnum, 4)
        N1 = 10
        N2 = CInt(Left$(Wnum, 1))
        N3 = CInt(Mid$(Wnum, 2, 1))
        N4 = CInt(Mid$(Wnum, 3, 1))
        N5 = CInt(Right$(Wnum, 1))
    Case Is >= 8
        Wnum = Left$(Wnum, 5)
        N1 = CInt(Left$(Wnum, 1))
        N2 = CInt(Mid$(Wnum, 2, 1))
        N3 = CInt(Mid$(Wnum, 3, 1))
        N4 = CInt(Mid$(Wnum, 4, 1))
        N5 = CInt(Right$(Wnum, 1))
End Select

'Seteamos los displays con los numeros correspondientes
Select Case WEstNum
    Case 1, "1"     '****************** ESTACION 01
        Select Case WType
            Case "Normal"
                Est01.E1p0.Picture = TopMenu.SmallClip.GraphicCell(10) '= nada
                Est01.E1p1.Picture = TopMenu.SmallClip.GraphicCell(N1)
                Est01.E1p2.Picture = TopMenu.SmallClip.GraphicCell(N2)
                Est01.E1p3.Picture = TopMenu.SmallClip.GraphicCell(N3)
                Est01.E1p4.Picture = TopMenu.SmallClip.GraphicCell(N4)
                Est01.E1p5.Picture = TopMenu.SmallClip.GraphicCell(N5)
            Case "Restante"
                Est01.E1p0.Picture = TopMenu.SmallClip.GraphicCell(13) '= signo menos
                Est01.E1p1.Picture = TopMenu.SmallClip.GraphicCell(N1)
                Est01.E1p2.Picture = TopMenu.SmallClip.GraphicCell(N2)
                Est01.E1p3.Picture = TopMenu.SmallClip.GraphicCell(N3)
                Est01.E1p4.Picture = TopMenu.SmallClip.GraphicCell(N4)
                Est01.E1p5.Picture = TopMenu.SmallClip.GraphicCell(N5)
        End Select
    Case 2, "2"     '****************** ESTACION 02
        Select Case WType
            Case "Normal"
                Est02.E2p0.Picture = TopMenu.SmallClip.GraphicCell(10) '= nada
                Est02.E2p1.Picture = TopMenu.SmallClip.GraphicCell(N1)
                Est02.E2p2.Picture = TopMenu.SmallClip.GraphicCell(N2)
                Est02.E2p3.Picture = TopMenu.SmallClip.GraphicCell(N3)
                Est02.E2p4.Picture = TopMenu.SmallClip.GraphicCell(N4)
                Est02.E2p5.Picture = TopMenu.SmallClip.GraphicCell(N5)
            Case "Restante"
                Est02.E2p0.Picture = TopMenu.SmallClip.GraphicCell(13) '= signo menos
                Est02.E2p1.Picture = TopMenu.SmallClip.GraphicCell(N1)
                Est02.E2p2.Picture = TopMenu.SmallClip.GraphicCell(N2)
                Est02.E2p3.Picture = TopMenu.SmallClip.GraphicCell(N3)
                Est02.E2p4.Picture = TopMenu.SmallClip.GraphicCell(N4)
                Est02.E2p5.Picture = TopMenu.SmallClip.GraphicCell(N5)
        End Select
End Select

End Sub

Public Sub SetStartTime()

Dim Atime As String, NTime As String
Dim IDX As Integer
Dim Data1 As String
Dim Data2 As String
Dim Sum1 As Long, Sum2 As Long, Rst As Long
Dim OldIndex As Integer

Atime = time$       'hora actual

'extraemos el index del tema actual
OldIndex = Tanda01.T1View.SelectedItem.index

'extraemos los datos del ultimo tema en la Tanda
On Error GoTo nop
IDX = Tanda01.T1View.ListItems.count
Tanda01.T1View.ListItems.Item(IDX).Selected = True
'Data1 = Tanda01.T1View.SelectedItem.ListSubItems(3).Text     'duracion
'Data2 = Tanda01.T1View.SelectedItem.ListSubItems(4).Text     'hora de lanz

Sum1 = ConvMinToSec(Data1)
Sum2 = ConvMinToSec(Data2)
Rst = Sum2 + Sum1
NTime = ConvSecToMin(Rst)

'seteamos los relojes digitales
SetSumTime Atime, 2     'hora de comiento = hora actual
SetSumTime NTime, 3     'hora de finalizacion = op1+op2

'seteamos el label para el tiempo
Tanda01.FTime.Caption = NTime

'restauramos la seleccion al item del usuario
Tanda01.T1View.ListItems.Item(OldIndex).Selected = True
Exit Sub

nop:
Tanda01.T1View.ListItems.Item(OldIndex).Selected = True
End Sub
Public Sub SetTOPTime(DNum As String)

'SOLO PARA PHTIMER
'el dnum debe estar en minutos = 00:00:00

Dim N1, N2, N3, N4, N5, N6
Dim LenNum

LenNum = Len(DNum)

If LenNum > 8 Then
    Exit Sub
End If

If LenNum = 8 Then
    DNum = LTrim(RTrim(DNum))
    N1 = Left$(DNum, 1)
    N2 = Mid$(DNum, 2, 1)
    N3 = Mid$(DNum, 4, 1)
    N4 = Mid$(DNum, 5, 1)
    N5 = Mid$(DNum, 7, 1)
    N6 = Right$(DNum, 1)
        
    If N1 = 0 Then TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(0)
    If N1 = 1 Then TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(1)
    If N1 = 2 Then TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(2)
    If N1 = 3 Then TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(3)
    If N1 = 4 Then TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(4)
    If N1 = 5 Then TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(5)
    If N1 = 6 Then TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(6)
    If N1 = 7 Then TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(7)
    If N1 = 8 Then TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(8)
    If N1 = 9 Then TopMenu.Pht1.Picture = TopMenu.SmallClip.GraphicCell(9)
            
    If N2 = 0 Then TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(0)
    If N2 = 1 Then TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(1)
    If N2 = 2 Then TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(2)
    If N2 = 3 Then TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(3)
    If N2 = 4 Then TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(4)
    If N2 = 5 Then TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(5)
    If N2 = 6 Then TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(6)
    If N2 = 7 Then TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(7)
    If N2 = 8 Then TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(8)
    If N2 = 9 Then TopMenu.Pht2.Picture = TopMenu.SmallClip.GraphicCell(9)

    TopMenu.Pht3.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
        
    If N3 = 0 Then TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(0)
    If N3 = 1 Then TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(1)
    If N3 = 2 Then TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(2)
    If N3 = 3 Then TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(3)
    If N3 = 4 Then TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(4)
    If N3 = 5 Then TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(5)
    If N3 = 6 Then TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(6)
    If N3 = 7 Then TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(7)
    If N3 = 8 Then TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(8)
    If N3 = 9 Then TopMenu.Pht4.Picture = TopMenu.SmallClip.GraphicCell(9)
    
    If N4 = 0 Then TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(0)
    If N4 = 1 Then TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(1)
    If N4 = 2 Then TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(2)
    If N4 = 3 Then TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(3)
    If N4 = 4 Then TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(4)
    If N4 = 5 Then TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(5)
    If N4 = 6 Then TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(6)
    If N4 = 7 Then TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(7)
    If N4 = 8 Then TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(8)
    If N4 = 9 Then TopMenu.Pht5.Picture = TopMenu.SmallClip.GraphicCell(9)
        
    TopMenu.Pht6.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
    
    If N5 = 0 Then TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(0)
    If N5 = 1 Then TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(1)
    If N5 = 2 Then TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(2)
    If N5 = 3 Then TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(3)
    If N5 = 4 Then TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(4)
    If N5 = 5 Then TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(5)
    If N5 = 6 Then TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(6)
    If N5 = 7 Then TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(7)
    If N5 = 8 Then TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(8)
    If N5 = 9 Then TopMenu.Pht7.Picture = TopMenu.SmallClip.GraphicCell(9)
    
    If N6 = 0 Then TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(0)
    If N6 = 1 Then TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(1)
    If N6 = 2 Then TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(2)
    If N6 = 3 Then TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(3)
    If N6 = 4 Then TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(4)
    If N6 = 5 Then TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(5)
    If N6 = 6 Then TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(6)
    If N6 = 7 Then TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(7)
    If N6 = 8 Then TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(8)
    If N6 = 9 Then TopMenu.Pht8.Picture = TopMenu.SmallClip.GraphicCell(9)
End If

End Sub

Public Sub SetSumTime(DNum As String, ByVal WClock As Long)

'SOLO PARA TANDA
'el dnum debe estar en minutos = 00:00:00

Dim N1, N2, N3, N4, N5, N6
Dim LenNum

LenNum = Len(DNum)

If LenNum <> 8 Then
    Exit Sub
Else
    DNum = LTrim(RTrim(DNum))
    N1 = Left$(DNum, 1)
    N2 = Mid$(DNum, 2, 1)
    N3 = Mid$(DNum, 4, 1)
    N4 = Mid$(DNum, 5, 1)
    N5 = Mid$(DNum, 7, 1)
    N6 = Right$(DNum, 1)
End If

Select Case WClock
    Case 1
        If N1 = 0 Then Tanda01.T1T1.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N1 = 1 Then Tanda01.T1T1.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N1 = 2 Then Tanda01.T1T1.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N1 = 3 Then Tanda01.T1T1.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N1 = 4 Then Tanda01.T1T1.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N1 = 5 Then Tanda01.T1T1.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N1 = 6 Then Tanda01.T1T1.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N1 = 7 Then Tanda01.T1T1.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N1 = 8 Then Tanda01.T1T1.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N1 = 9 Then Tanda01.T1T1.Picture = TopMenu.SmallClip.GraphicCell(9)
            
        If N2 = 0 Then Tanda01.T1T2.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N2 = 1 Then Tanda01.T1T2.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N2 = 2 Then Tanda01.T1T2.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N2 = 3 Then Tanda01.T1T2.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N2 = 4 Then Tanda01.T1T2.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N2 = 5 Then Tanda01.T1T2.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N2 = 6 Then Tanda01.T1T2.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N2 = 7 Then Tanda01.T1T2.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N2 = 8 Then Tanda01.T1T2.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N2 = 9 Then Tanda01.T1T2.Picture = TopMenu.SmallClip.GraphicCell(9)

        Tanda01.T1T3.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
        
        If N3 = 0 Then Tanda01.T1T4.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N3 = 1 Then Tanda01.T1T4.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N3 = 2 Then Tanda01.T1T4.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N3 = 3 Then Tanda01.T1T4.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N3 = 4 Then Tanda01.T1T4.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N3 = 5 Then Tanda01.T1T4.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N3 = 6 Then Tanda01.T1T4.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N3 = 7 Then Tanda01.T1T4.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N3 = 8 Then Tanda01.T1T4.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N3 = 9 Then Tanda01.T1T4.Picture = TopMenu.SmallClip.GraphicCell(9)
    
        If N4 = 0 Then Tanda01.T1T5.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N4 = 1 Then Tanda01.T1T5.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N4 = 2 Then Tanda01.T1T5.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N4 = 3 Then Tanda01.T1T5.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N4 = 4 Then Tanda01.T1T5.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N4 = 5 Then Tanda01.T1T5.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N4 = 6 Then Tanda01.T1T5.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N4 = 7 Then Tanda01.T1T5.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N4 = 8 Then Tanda01.T1T5.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N4 = 9 Then Tanda01.T1T5.Picture = TopMenu.SmallClip.GraphicCell(9)
        
        Tanda01.T1t6.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
        
        If N5 = 0 Then Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N5 = 1 Then Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N5 = 2 Then Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N5 = 3 Then Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N5 = 4 Then Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N5 = 5 Then Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N5 = 6 Then Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N5 = 7 Then Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N5 = 8 Then Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N5 = 9 Then Tanda01.T1t7.Picture = TopMenu.SmallClip.GraphicCell(9)
        
        If N6 = 0 Then Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N6 = 1 Then Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N6 = 2 Then Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N6 = 3 Then Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N6 = 4 Then Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N6 = 5 Then Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N6 = 6 Then Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N6 = 7 Then Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N6 = 8 Then Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N6 = 9 Then Tanda01.T1t8.Picture = TopMenu.SmallClip.GraphicCell(9)
    
    Case 2
        If N1 = 0 Then Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N1 = 1 Then Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N1 = 2 Then Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N1 = 3 Then Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N1 = 4 Then Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N1 = 5 Then Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N1 = 6 Then Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N1 = 7 Then Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N1 = 8 Then Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N1 = 9 Then Tanda01.T1I1.Picture = TopMenu.SmallClip.GraphicCell(9)
                
        If N2 = 0 Then Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N2 = 1 Then Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N2 = 2 Then Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N2 = 3 Then Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N2 = 4 Then Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N2 = 5 Then Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N2 = 6 Then Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N2 = 7 Then Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N2 = 8 Then Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N2 = 9 Then Tanda01.T1I2.Picture = TopMenu.SmallClip.GraphicCell(9)
    
        Tanda01.T1I3.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
            
        If N3 = 0 Then Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N3 = 1 Then Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N3 = 2 Then Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N3 = 3 Then Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N3 = 4 Then Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N3 = 5 Then Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N3 = 6 Then Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N3 = 7 Then Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N3 = 8 Then Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N3 = 9 Then Tanda01.T1I4.Picture = TopMenu.SmallClip.GraphicCell(9)
        
        If N4 = 0 Then Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N4 = 1 Then Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N4 = 2 Then Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N4 = 3 Then Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N4 = 4 Then Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N4 = 5 Then Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N4 = 6 Then Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N4 = 7 Then Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N4 = 8 Then Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N4 = 9 Then Tanda01.T1I5.Picture = TopMenu.SmallClip.GraphicCell(9)
            
        Tanda01.T1I6.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
        
        If N5 = 0 Then Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N5 = 1 Then Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N5 = 2 Then Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N5 = 3 Then Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N5 = 4 Then Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N5 = 5 Then Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N5 = 6 Then Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N5 = 7 Then Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N5 = 8 Then Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N5 = 9 Then Tanda01.T1I7.Picture = TopMenu.SmallClip.GraphicCell(9)
        
        If N6 = 0 Then Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N6 = 1 Then Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N6 = 2 Then Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N6 = 3 Then Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N6 = 4 Then Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N6 = 5 Then Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N6 = 6 Then Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N6 = 7 Then Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N6 = 8 Then Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N6 = 9 Then Tanda01.T1I8.Picture = TopMenu.SmallClip.GraphicCell(9)

    Case 3
        If N1 = 0 Then Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N1 = 1 Then Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N1 = 2 Then Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N1 = 3 Then Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N1 = 4 Then Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N1 = 5 Then Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N1 = 6 Then Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N1 = 7 Then Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N1 = 8 Then Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N1 = 9 Then Tanda01.T1F1.Picture = TopMenu.SmallClip.GraphicCell(9)
                
        If N2 = 0 Then Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N2 = 1 Then Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N2 = 2 Then Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N2 = 3 Then Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N2 = 4 Then Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N2 = 5 Then Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N2 = 6 Then Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N2 = 7 Then Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N2 = 8 Then Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N2 = 9 Then Tanda01.T1F2.Picture = TopMenu.SmallClip.GraphicCell(9)
    
        Tanda01.T1F3.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
            
        If N3 = 0 Then Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N3 = 1 Then Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N3 = 2 Then Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N3 = 3 Then Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N3 = 4 Then Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N3 = 5 Then Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N3 = 6 Then Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N3 = 7 Then Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N3 = 8 Then Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N3 = 9 Then Tanda01.T1F4.Picture = TopMenu.SmallClip.GraphicCell(9)
        
        If N4 = 0 Then Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N4 = 1 Then Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N4 = 2 Then Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N4 = 3 Then Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N4 = 4 Then Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N4 = 5 Then Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N4 = 6 Then Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N4 = 7 Then Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N4 = 8 Then Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N4 = 9 Then Tanda01.T1F5.Picture = TopMenu.SmallClip.GraphicCell(9)
            
        Tanda01.T1F6.Picture = TopMenu.SmallClip.GraphicCell(11) '= :
        
        If N5 = 0 Then Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N5 = 1 Then Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N5 = 2 Then Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N5 = 3 Then Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N5 = 4 Then Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N5 = 5 Then Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N5 = 6 Then Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N5 = 7 Then Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N5 = 8 Then Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N5 = 9 Then Tanda01.T1F7.Picture = TopMenu.SmallClip.GraphicCell(9)
        
        If N6 = 0 Then Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(0)
        If N6 = 1 Then Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(1)
        If N6 = 2 Then Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(2)
        If N6 = 3 Then Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(3)
        If N6 = 4 Then Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(4)
        If N6 = 5 Then Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(5)
        If N6 = 6 Then Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(6)
        If N6 = 7 Then Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(7)
        If N6 = 8 Then Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(8)
        If N6 = 9 Then Tanda01.T1F8.Picture = TopMenu.SmallClip.GraphicCell(9)
            
End Select

End Sub

Public Function SetTimerMilisec(ByVal WSegIn As Long, ByVal WSegOut As Long) As Long

Dim CTimeIn As Long
Dim CTimeOut As Long
Dim Result As Long

'chequeos varios
If WSegOut = 0 Then
    MsgBox LoadResString(152), vbInformation
    SetTimerMilisec = 0
    Exit Function
End If
If WSegIn > WSegOut Then
    SetTimerMilisec = 0
    MsgBox LoadResString(153), vbInformation
    Exit Function
End If

'extraemos la diferencia entre ellos
Result = WSegOut - WSegIn

'convertimos los segundos en milisegundos para setear el timer
Result = Result * 1000

'finalizamos
SetTimerMilisec = Result

End Function
