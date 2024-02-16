Attribute VB_Name = "ChipherMod"
'/////////////////////////////////////////
'
' Cab Module Managger.
' Copyright (c) 2002 ONLY development inc.
'           reservados todos los derechos.
' Christian A. Del Monte
'/////////////////////////////////////////

Option Explicit

'///////////////////////////////////////////////
'* Password:   encryption password
'* TxtOrigen:  encrypted text
'* return:     text
'///////////////////////////////////////////////

Function Desencriptar(ByVal Password As String, ByVal TxtOrigen As String) As String

Const MIN_ASC = 1
Const MAX_ASC = 255
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer
Dim TxtDestino As String

    'Inicializar el generador de numeros aleatorios
    offset = ClaveNumerica(Password)
    Rnd -1
    Randomize offset
    
    'Desencriptar el texto
    str_len = Len(TxtOrigen)
    For i = 1 To str_len
        ch = Asc(Mid$(TxtOrigen, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            TxtDestino = TxtDestino & Chr$(ch)
        End If
    Next i

Desencriptar = TxtDestino

End Function

'///////////////////////////////////////////////
'* Password:   encryption password
'* TxtOrigen:  text to encrypt
'* return:     encrypted text
'///////////////////////////////////////////////

Function Encriptar(ByVal Password As String, ByVal TxtOrigen As String) As String

Const MIN_ASC = 1
Const MAX_ASC = 255
Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer
Dim TxtDestino As String

    'Inicializar el generador de numeros aleatorios
    offset = ClaveNumerica(Password)
    Rnd -1
    Randomize offset
    
    'Encriptar el texto
    str_len = Len(TxtOrigen)
    For i = 1 To str_len
        ch = Asc(Mid$(TxtOrigen, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            TxtDestino = TxtDestino & Chr$(ch)
        End If
    Next i

Encriptar = TxtDestino

End Function

Function ClaveNumerica(ByVal Password As String) As Long

Dim value As Long
Dim ch As Long
Dim shift1 As Long
Dim shift2 As Long
Dim i As Integer
Dim str_len As Integer

    str_len = Len(Password)
    For i = 1 To str_len
        ch = Asc(Mid$(Password, i, 1))
        value = value Xor (ch * 2 ^ shift1)
        value = value Xor (ch * 2 ^ shift2)
        
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    ClaveNumerica = value

End Function

