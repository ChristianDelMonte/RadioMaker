Attribute VB_Name = "RMD"

'********************* RMD ***********************
'            RADIO MAKER DATA MODULE
'COPYRIGHT (C) 1987-2002 CREACIONES DIGITALES INC.
'*************************************************

Option Explicit

Private Const MIN_ASC = 1
Private Const MAX_ASC = 255
Private Const NUM_ASC = MAX_ASC - MIN_ASC + 1

Function DecipherData(WPass As String, InText As String) As String

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer
Dim OutText As String

    'chequeos necesarios
    If WPass = "" Or WPass = " " Then
        DecipherData = ""
        Exit Function
    End If
    If InText = "" Or InText = " " Then
        DecipherData = ""
        Exit Function
    End If
    
    'Inicializar el generador de numeros aleatorios
    OutText = ""
    offset = CN(WPass)
    Rnd -1
    Randomize offset
    
    'Desencriptar el texto
    str_len = Len(InText)
    For i = 1 To str_len
        ch = Asc(Mid$(InText, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch - offset) Mod NUM_ASC)
            If ch < 0 Then ch = ch + NUM_ASC
            ch = ch + MIN_ASC
            OutText = OutText & Chr$(ch)
        End If
    Next i

    DecipherData = OutText

End Function

Function CipherData(WPass As String, InText As String) As String

Dim offset As Long
Dim str_len As Integer
Dim i As Integer
Dim ch As Integer
Dim OutText As String

    'chequeos necesarios
    If WPass = "" Or WPass = " " Then
        CipherData = ""
        Exit Function
    End If
    If InText = "" Or InText = " " Then
        CipherData = ""
        Exit Function
    End If

    'Inicializar el generador de numeros aleatorios
    OutText = ""
    offset = CN(WPass)
    Rnd -1
    Randomize offset
    
    'Encriptar el texto
    str_len = Len(InText)
    For i = 1 To str_len
        ch = Asc(Mid$(InText, i, 1))
        If ch >= MIN_ASC And ch <= MAX_ASC Then
            ch = ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            ch = ((ch + offset) Mod NUM_ASC)
            ch = ch + MIN_ASC
            OutText = OutText & Chr$(ch)
        End If
    Next i

    CipherData = OutText
    
End Function

Private Function CN(ByVal WnPass As String) As Long

Dim value As Long
Dim ch As Long
Dim shift1 As Long
Dim shift2 As Long
Dim i As Integer
Dim str_len As Integer

    str_len = Len(WnPass)
    For i = 1 To str_len
        ch = Asc(Mid$(WnPass, i, 1))
        value = value Xor (ch * 2 ^ shift1)
        value = value Xor (ch * 2 ^ shift2)
        
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    
    CN = value

End Function
