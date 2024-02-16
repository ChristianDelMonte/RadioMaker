Attribute VB_Name = "BlockModule"
'////////////////////////////////////////////////////
'*
'*  // BLOCK managger module for Vb.6+ //
'*  ** module for Radiomaker 1.0 only **
'*  Copyright (c) 1987-2002 Only development Inc.
'*  Christian A. Del Monte
'///////////////////////////////////////////////////

Option Explicit

Public Const AppBlockDir = "\Blocks"            'Directorio de bloques
Public Const AppBlockFileExt = ".blk"     'extension de archivos de bloques

'*** Type para archivos de bloques.
Type BlockRecord
    id As Integer
    FFilePath As String * 255           'path del archivo
    FFileName As String * 255           'nombre del archivo
    FFileDur As String * 8              'duracion del archivo
    FPrefH(0 To 2) As Integer           'horario pref. (enum: BlockPrefH)
    FPrefD(0 To 2) As Integer           'dia pref.     (enum: BlockPrefD)
    FCantV(0 To 2) As Integer           'cantidad de veces a reproducir?
    FPubInit As String * 10             'fecha de inicio de reproduccion
    FPubFin As String * 10              'fecha de fin de reproduccion
End Type

Type BlockRecordSearch
    FPrefH As Integer                   'horario pref. (enum: BlockPrefH)
    FPrefD As Integer                   'dia pref.     (enum: BlockPrefD)
    FPubInit As String * 10             'fecha de inicio de reproduccion
    FPubFin As String * 10              'fecha de fin de reproduccion
End Type

Public Enum BlockPrefH
    All = 0                             'cualquier horario
    d1a2 = 1
    d2a3 = 2
    d3a4 = 3
    d4a5 = 4
    d5a6 = 5
    d6a7 = 6
    d7a8 = 7
    d8a9 = 8
    d9a10 = 9
    d10a11 = 10
    d11a12 = 11
    d12a13 = 12
    d13a14 = 13
    d14a15 = 14
    d15a16 = 15
    d16a17 = 16
    d17a18 = 17
    d18a19 = 18
    d19a20 = 19
    d20a21 = 20
    d21a22 = 21
    d22a23 = 22
    d23a0 = 23
    d0a1 = 24
    Vacio = 99                          'nada especificado
End Enum

Public Enum BlockPrefD
    All = 0                             'cualquier dia ó todos los días
    Dom = 1
    Lun = 2
    Mar = 3
    Mie = 4
    Jue = 5
    Vie = 6
    Sab = 7
    Vacio = 99                          'nada especificado
End Enum

Public BlockData As BlockRecord         'registros de archivos de bloques
Public BlockSearch As BlockRecordSearch 'busqueda de registros dentro de bloques

Private Function BlockIsOutOfDate(WStartDate As String, WEndDate As String) As Boolean

Dim SDay As Long, SMonth As Long, sYear As Long
Dim FDay As Long, FMonth As Long, FYear As Long
Dim ADay As Long, AMonth As Long, AYear As Long

'/// ERROR check
If IsDate(WStartDate) = False Or IsDate(WEndDate) = False Then
    '/// error fecha incorrecta
    BlockIsOutOfDate = True
    Exit Function
End If

If Len(Trim(WStartDate)) <> 10 Or Len(Trim(WEndDate)) <> 10 Then
    '/// error fecha incorrecta
    BlockIsOutOfDate = True
    Exit Function
End If

'/// process the user date
SDay = CLng(Left$(WStartDate, 2))
SMonth = CLng(Mid$(WStartDate, 4, 2))
sYear = CLng(Right$(WStartDate, 4))
'-----
FDay = CLng(Left$(WEndDate, 2))
FMonth = CLng(Mid$(WEndDate, 4, 2))
FYear = CLng(Right$(WEndDate, 4))

'/// process the system date
ADay = CLng(Left$(Date, 2))
AMonth = CLng(Mid$(Date, 4, 2))
AYear = CLng(Right$(Date, 4))

'/// lets compare the date start
If AYear = sYear Then
    If AMonth = SMonth Then
        If ADay = SDay Then
            GoTo CheckEnd
        Else
            If ADay > SDay Then
                GoTo CheckEnd
            Else
                BlockIsOutOfDate = True
            End If
        End If
    Else
        If AMonth > SMonth Then
            GoTo CheckEnd
        Else
            BlockIsOutOfDate = True
        End If
    End If
Else
    If AYear > sYear Then
        GoTo CheckEnd
    Else
        BlockIsOutOfDate = True
    End If
End If
Exit Function

CheckEnd:
'/// lets compare the date end
If AYear = FYear Then
    If AMonth = FMonth Then
        If ADay = FDay Then
            BlockIsOutOfDate = False
        Else
            If ADay < FDay Then
                BlockIsOutOfDate = False
            Else
                BlockIsOutOfDate = True
            End If
        End If
    Else
        If AMonth < FMonth Then
            BlockIsOutOfDate = False
        Else
            BlockIsOutOfDate = True
        End If
    End If
Else
    If AYear < FYear Then
        BlockIsOutOfDate = False
    Else
        BlockIsOutOfDate = True
    End If
End If

End Function

Private Sub OrderTandaBlocks(WNewData As BlockRecord)

Dim DataA(0 To 9) As String
Dim DataKa As String

Dim DataB(0 To 9) As String
Dim DataKb As String
Dim nIndex As Integer
Dim ItmX As ListItem

Dim ONum As Integer
Dim nCount As Integer
Dim NNum As Integer

On Error GoTo Continue

'/// chequeos necesarios
nCount = Tanda01.T1View.ListItems.count
ONum = Tanda01.T1View.SelectedItem.index + 1
NNum = ONum + 1

nIndex = ONum
Tanda01.T1View.ListItems.Item(nIndex).Selected = True

'/// extraemos los datos del item seleccionado
DataA(0) = Tanda01.T1View.SelectedItem.text                     'file & path
'DataA(1) = Tanda01.T1View.SelectedItem.ListSubItems(1).Text     'filetype
'DataA(2) = Tanda01.T1View.SelectedItem.ListSubItems(2).Text     'filename
'DataA(3) = Tanda01.T1View.SelectedItem.ListSubItems(3).Text     'duracion del tema
'DataA(4) = Tanda01.T1View.SelectedItem.ListSubItems(4).Text     'hora de lanzamiento
'DataA(5) = Tanda01.T1View.SelectedItem.ListSubItems(5).Text     'file & path (mix)
'DataA(6) = Tanda01.T1View.SelectedItem.ListSubItems(6).Text     'filetype (mix)
'DataA(7) = Tanda01.T1View.SelectedItem.ListSubItems(7).Text     'filename (mix)
'DataA(8) = Tanda01.T1View.SelectedItem.ListSubItems(8).Text     'duracion del mix
'DataA(9) = Tanda01.T1View.SelectedItem.ListSubItems(9).Text     'hora de lanzamiento (mix)
DataKa = Tanda01.T1View.SelectedItem.Key

'/// borramos el item seleccionado de la lista
Tanda01.T1View.ListItems.Remove (nIndex)

    'id As Integer                 'identificador o numero de registro
    'Name As String * 255          'nombre del tema o nombre del bloque
                                  ' ((si es un bloque debe comienzar con BLOCK:xx))
    'FNType As String * 10         'tipo de archivo (stream or music)
    'Direccion As String * 255     'path del tema
    'Duracion As String * 8        'duracion del tema '00:00:00
                                  ' ((o duracion del bloque si es BLOCK:))
    'Hora As String * 8            'hora de lanzamiento del tema '00:00:00
                                  ' ((u hora de lanzamiento del bloque si es BLOCK:))
    'NameX As String * 255         'nombre del tema mixado
                                  ' ((= BLOCK: si es un bloque))
    'FNTypeX As String * 10        'tipo de archivo (stream or music)
                                  ' ((=vacio si es BLOCK:))
    'DireccionX As String * 255    'path del tema de mixado
                                  ' ((=vacio si es BLOCK:))
    'DuracionX As String * 8       'duracion del mixado '00:00:00
                                  ' ((=vacio si es BLOCK:))
    'HoraX As String * 5           'hora de lanzamiento del mixado '00:00
                                  ' ((u hora predeterminada de lanz. si es BLOCK:))

'/// ponemos los nuevos datos (los datos del block)
'Set ItmX = Tanda01.T1View.ListItems.Add(nIndex, DataKb, Trim(WNewData.FFilePath)) 'path & file
'ItmX.SubItems(1) = "Stream"
'ItmX.SubItems(2) = "BLOCK: " & WNewData.FFileName
'ItmX.SubItems(3) = WNewData.FFileDur
'ItmX.SubItems(4) = "00:00:00"
'ItmX.SubItems(5) = "BLOCK"
'ItmX.SubItems(6) = ""
'ItmX.SubItems(7) = ""
'ItmX.SubItems(8) = ""
'ItmX.SubItems(9) = "00:00"

'///////////////////////////////////////////////////
'///////////////////// FIX /////////////////////////
'///////////////////////////////////////////////////

'seleccionamos el siguiente item hacia abajo
nIndex = NNum
Tanda01.T1View.ListItems.Item(nIndex).Selected = True

'extraemos los datos del item
DataB(0) = Tanda01.T1View.SelectedItem.text    'file & path
'DataB(1) = Tanda01.T1View.SelectedItem.ListSubItems(1).Text   'filetype
'DataB(2) = Tanda01.T1View.SelectedItem.ListSubItems(2).Text  'filename
'DataB(3) = Tanda01.T1View.SelectedItem.ListSubItems(3).Text
'DataB(4) = Tanda01.T1View.SelectedItem.ListSubItems(4).Text
'DataB(5) = Tanda01.T1View.SelectedItem.ListSubItems(5).Text
'DataB(6) = Tanda01.T1View.SelectedItem.ListSubItems(6).Text
'DataB(7) = Tanda01.T1View.SelectedItem.ListSubItems(7).Text
'DataB(8) = Tanda01.T1View.SelectedItem.ListSubItems(8).Text
'DataB(9) = Tanda01.T1View.SelectedItem.ListSubItems(9).Text
DataKb = Tanda01.T1View.SelectedItem.Key

'ponemos los nuevos datos
'Set ItmX = Tanda01.T1View.ListItems.Add(nIndex, DataKb, DataA(0)) 'path & file
'ItmX.SubItems(1) = DataA(1)
'ItmX.SubItems(2) = DataA(2)
'ItmX.SubItems(3) = DataA(3)
'ItmX.SubItems(4) = DataA(4)
'ItmX.SubItems(5) = DataA(5)
'ItmX.SubItems(6) = DataA(6)
'ItmX.SubItems(7) = DataA(7)
'ItmX.SubItems(8) = DataA(8)
'ItmX.SubItems(9) = DataA(9)

'seleccionamos el index anterior
nIndex = nIndex - 1
Tanda01.T1View.ListItems.Item(nIndex).Selected = True

'ponemos los nuevos datos
Tanda01.T1View.ListItems.Remove (nIndex)
'Set ItmX = Tanda01.T1View.ListItems.Add(nIndex, DataKa, DataB(0)) 'path & file
'ItmX.SubItems(1) = DataB(1)
'ItmX.SubItems(2) = DataB(2)
'ItmX.SubItems(3) = DataB(3)
'ItmX.SubItems(4) = DataB(4)
'ItmX.SubItems(5) = DataB(5)
'ItmX.SubItems(6) = DataB(6)
'ItmX.SubItems(7) = DataB(7)
'ItmX.SubItems(8) = DataB(8)
'ItmX.SubItems(9) = DataB(9)

'una vez finalizado. seleccionamos el item
nIndex = nIndex + 1
Tanda01.T1View.ListItems.Item(nIndex).Selected = True
Exit Sub

Continue:
    'nothing to do....

End Sub

Function SetTandaBlocks() As Boolean

Dim i As Integer, X As Integer, Z As Integer, Y As Integer
Dim NMReg As Integer, TTReg As Integer
Dim BlName As String, AATime As String
Dim SrchTime As Double, RstH As Integer, RstD As Integer
Dim RstData As BlockRecord, nIndex As Integer, nCount As Integer
Dim Result As Boolean

'//// ahora chequeamos por publicidad de bloques
If Tanda01.LBlk.Caption = "/ Auto" Then
    '/// procesamos el archivo de bloques publicitarios
    BlName = Trim(Tanda01.BlkFn.Caption)
    If FileExist(BlName) = False Then
        '/// error!!! archivo de bloques no se encuentra
        SetTandaBlocks = False
        Exit Function
    End If
    '/// get the selected file data
    nIndex = Tanda01.T1View.SelectedItem.index   'numero de index
    nCount = Tanda01.T1View.ListItems.count      'total de streams en la lista
    For i = nIndex To nCount    'start from stream selected in the list
        If nIndex >= nCount Then
            Exit For
        End If
        AATime = Tanda01.T1View.SelectedItem.ListSubItems(4).text   'hora de lanz
        SrchTime = ConvMinToSec(AATime)
        RstH = GetBLockTime(SrchTime)
        RstD = Weekday(Date, vbUseSystemDayOfWeek)
        '/// block file process
        TTReg = GetBlockLastReg(BlName)     'get total of regs in file
        For X = 1 To TTReg
            RstData = OpenBlockFile(BlName, "", X)
            '///compare the data in the file
            For Z = 0 To 2
                'check for time
                If RstData.FPrefH(Z) = RstH Then
                    For Y = 0 To 2
                        'check for date
                        If RstData.FPrefD(Y) = RstD Then
                            'check for out of date (vencimiento)
                            If BlockIsOutOfDate(Trim(RstData.FPubInit), Trim(RstData.FPubFin)) = False Then
                                If X >= TTReg Then
                                    '/////////////////////////////////////
                                    '///// PUBLICATE THE FILE IN THE TANDA
                                    '/////////////////////////////////////
                                    Call OrderTandaBlocks(RstData)
                                    
                                Else
                                    '/////////////////////////////////////
                                    '///// PUBLICATE THE FILE IN THE TANDA
                                    '/////////////////////////////////////
                                    Call OrderTandaBlocks(RstData)
                                End If
                            End If
                        End If
                    Next Y
                End If
            Next Z
        Next X
    Next i
Else
    '/// sin bloques publicitarios
    'xxxxx
End If

End Function

'//////////////////////////////////////////////////
'*
'* SAVEBLOCKFILE usage:
'* BlockFileName = nombre de archivo de bloque
'* Wdata = datos a guardar (formato BlockRecord)
'* OptionalID = ID de los datos a guardar (opcio-
'*              nales). Si no se especifica el ID
'*              los datos se guardan en el ultimo
'*              registro valido dentro del archivo
'* Return = false if there is an error
'*
'//////////////////////////////////////////////////

Function SaveBlockFile(BlockFileName As String, WData As BlockRecord, OptionalID As Integer) As Boolean

Dim LastReg As Integer

'/// start
'/// check the file for correct extension
If LCase(StripExtFromFile(BlockFileName)) = AppBlockFileExt Then
    BlockFileName = BlockFileName
Else
    BlockFileName = StripFileFromExt(BlockFileName) & AppBlockFileExt
End If

'/// abrimos el archivo de bloques
On Error GoTo err
Open BlockFileName For Random As #38 Len = Len(BlockData)

'/// check for the ID to save
If OptionalID = 0 Then
    LastReg = LOF(38) \ Len(BlockData)
    LastReg = LastReg + 1
Else
    LastReg = OptionalID
End If

'/// seteamos los datos de la configuracion a guardar
BlockData.id = LastReg
BlockData.FFilePath = WData.FFilePath           'path
BlockData.FFileName = WData.FFileName           'filename
BlockData.FFileDur = WData.FFileDur             'duracion
BlockData.FPrefH(0) = WData.FPrefH(0)           'horario predefinido
BlockData.FPrefH(1) = WData.FPrefH(1)           '
BlockData.FPrefH(2) = WData.FPrefH(2)
BlockData.FPrefD(0) = WData.FPrefD(0)           'dias predefinidos
BlockData.FPrefD(1) = WData.FPrefD(1)           '
BlockData.FPrefD(2) = WData.FPrefD(2)
BlockData.FCantV(0) = WData.FCantV(0)           'cantidad de veces
BlockData.FCantV(1) = WData.FCantV(1)           '
BlockData.FCantV(2) = WData.FCantV(2)
BlockData.FPubInit = WData.FPubInit             'dia de inicio
BlockData.FPubFin = WData.FPubFin               'dia de finalizacion

Put #38, LastReg, BlockData
Close #38
SaveBlockFile = True
Exit Function

'/// if the is an error ------------------------------------------
err:
Close #38
SaveBlockFile = False
End Function

'////////////////////////////////////////////////
'*
'* OPENBLOCKFILE usage:
'* BlockFileName= nombre de archivo de bloque
'* BlockSearchID= nombre de archivo de busqueda
'*                dentro del archivo que bloque.
'* OptionalID = ID opcional de datos. (numero
'*              de registro de busqueda).
'*
'* Si no se especifica un OptionalID, se usa
'* BlockSearchID como punto de busqueda.
'*
'* Return = datos resultado de la busqueda del
'*          registro en formato BlockRecord.
'*
'///////////////////////////////////////////////

Function OpenBlockFile(BlockFileName As String, BlockSearchID As String, OptionalID As Integer) As BlockRecord

Dim LastReg As Integer, i As Integer

If FileExist(BlockFileName) = False Or Trim(BlockSearchID) = "" Then
    Exit Function
End If

'/// abrimos el archivo de bloques
On Error GoTo err
Open BlockFileName For Random As #38 Len = Len(BlockData)

'/// get the nums of regs in the file
LastReg = LOF(38) \ Len(BlockData)
LastReg = LastReg

'/// Atención: cuando hay mas de un registro con el mismo nombre
'/// pero con diferente numero de registro. Solo se obtendrá como
'/// resultado el registro de menor valor o el que se encuentre
'/// primero. Siendo el resultado siempre 1(uno). No se pueden
'/// guardar registros iguales o con el mismo nombre de archivo.
If OptionalID = 0 Then
    For i = 1 To LastReg
        Get #38, i, BlockData
        If Trim(BlockData.FFileName) = Trim(BlockSearchID) Then
            OpenBlockFile = BlockData
            Close #38: Exit For: Exit Function
        End If
    Next i
Else
    Get #38, OptionalID, BlockData
    OpenBlockFile = BlockData
End If

Close #38
Exit Function

'/// if the is an error ------------------------------------------
err:
Close #38
End Function

'/////////////////////////////////////////////////
'*
'* GETBLOCKLASTREG: (function usage)
'* BlockFileName = Nombre de archivo de bloques.
'* Return: numero total de registros en el ar-
'*         chivo. ó 0 si ocurre un error.
'*
'////////////////////////////////////////////////

Private Function GetBlockLastReg(BlockFileName As String) As Integer

Dim LastReg As Integer

If FileExist(BlockFileName) = False Then
    GetBlockLastReg = 0
    Exit Function
End If

'/// abrimos el archivo de bloques
On Error GoTo err
Open BlockFileName For Random As #39 Len = Len(BlockData)

'/// get the nums of regs in the file
LastReg = LOF(39) \ Len(BlockData)

'/// return the number of regs in the file
Close #38
GetBlockLastReg = LastReg

Exit Function

'/// if there is an error ------------------------------------------
err:
Close #38
GetBlockLastReg = 0
End Function

Private Function GetBLockTime(WTime As Double) As Integer

Select Case WTime
    Case Is >= 0 < 3600
        GetBLockTime = BlockPrefH.d0a1
    Case Is >= 3600 < 7200
        GetBLockTime = BlockPrefH.d1a2
    Case Is >= 7200 < 10800
        GetBLockTime = BlockPrefH.d2a3
    Case Is >= 10800 < 14400
        GetBLockTime = BlockPrefH.d3a4
    Case Is >= 14400 < 18000
        GetBLockTime = BlockPrefH.d4a5
    Case Is >= 18000 < 21600
        GetBLockTime = BlockPrefH.d5a6
    Case Is >= 21600 < 25200
        GetBLockTime = BlockPrefH.d6a7
    Case Is >= 25200 < 28800
        GetBLockTime = BlockPrefH.d7a8
    Case Is >= 28800 < 32400
        GetBLockTime = BlockPrefH.d8a9
    Case Is >= 32400 < 36000
        GetBLockTime = BlockPrefH.d9a10
    Case Is >= 36000 < 39600
        GetBLockTime = BlockPrefH.d10a11
    Case Is >= 39600 < 43200
        GetBLockTime = BlockPrefH.d11a12
    Case Is >= 43200 < 46800
        GetBLockTime = BlockPrefH.d12a13
    Case Is >= 46800 < 50400
        GetBLockTime = BlockPrefH.d13a14
    Case Is >= 50400 < 54000
        GetBLockTime = BlockPrefH.d14a15
    Case Is >= 54000 < 57600
        GetBLockTime = BlockPrefH.d15a16
    Case Is >= 57600 < 61200
        GetBLockTime = BlockPrefH.d16a17
    Case Is >= 61200 < 64800
        GetBLockTime = BlockPrefH.d17a18
    Case Is >= 64800 < 68400
        GetBLockTime = BlockPrefH.d18a19
    Case Is >= 68400 < 72000
        GetBLockTime = BlockPrefH.d19a20
    Case Is >= 72000 < 75600
        GetBLockTime = BlockPrefH.d20a21
    Case Is >= 75600 < 79200
        GetBLockTime = BlockPrefH.d21a22
    Case Is >= 79200 < 82800
        GetBLockTime = BlockPrefH.d22a23
    Case Is >= 82800 < 86400
        GetBLockTime = BlockPrefH.d23a0
    Case Else
        GetBLockTime = BlockPrefH.All
End Select

End Function

