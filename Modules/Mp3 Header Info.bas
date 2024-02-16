Attribute VB_Name = "MP3Header"

'### Módulo para el manejo de archivos Mp3.
'### Proporcionado por:
'###
'### Copyright (c) Shannon Harmon
'### Sharmon@microtechcomputers.com

Option Explicit
Public sGenreMatrix

Type info
    sTitle As String * 30
    sArtist As String * 30
    sAlbum As String * 30
    sComment As String * 30
    sYear As String * 4
    sGenre As String * 21 ' NEW
End Type

Type HeaderInfo
    Layer As String
    Frequency As String
    Bitrate As String
    mode As String
    MpegVersion As String
    Emphasis As String
    FPlayTime As String 'Formatted playing time - 04:32
    mFileSize As String
End Type

Public MP3Info As info      'informacion del mp3 - artista, genero, etc.
Public MP3HInfo As HeaderInfo   'informacion de compresion y version del mp3

Sub AddGenList()

'agregamos los generos MP3 al
'combo box de generos en el Startup.

'Startup.Mgenero.AddItem "Blues"
'Startup.Mgenero.AddItem "Classic Rock"
'Startup.Mgenero.AddItem "Country"
'Startup.Mgenero.AddItem "Dance"
'Startup.Mgenero.AddItem "Disco"
'Startup.Mgenero.AddItem "Funk"
'Startup.Mgenero.AddItem "Grunge"
'Startup.Mgenero.AddItem "Hip-Hop"
'Startup.Mgenero.AddItem "Jazz"
'Startup.Mgenero.AddItem "Metal"
'Startup.Mgenero.AddItem "New Age"
'Startup.Mgenero.AddItem "Oldies"
'Startup.Mgenero.AddItem "Other"
'Startup.Mgenero.AddItem "Pop"
'Startup.Mgenero.AddItem "R&B"
'Startup.Mgenero.AddItem "Rap"
'Startup.Mgenero.AddItem "Reggae"
'Startup.Mgenero.AddItem "Rock"
'Startup.Mgenero.AddItem "Techno"
'Startup.Mgenero.AddItem "Industrial"
'Startup.Mgenero.AddItem "Alternative"
'Startup.Mgenero.AddItem "Ska"
'Startup.Mgenero.AddItem "Death Metal"
'Startup.Mgenero.AddItem "Pranks"
'Startup.Mgenero.AddItem "Soundtrack"
'Startup.Mgenero.AddItem "Euro-Techno"
'Startup.Mgenero.AddItem "Ambient"
'Startup.Mgenero.AddItem "Trip Hop"
'Startup.Mgenero.AddItem "Vocal"
'Startup.Mgenero.AddItem "Jazz+Funk"
'Startup.Mgenero.AddItem "Fusion"
'Startup.Mgenero.AddItem "Trance"
'Startup.Mgenero.AddItem "Classical"
'Startup.Mgenero.AddItem "Instrumental"
'Startup.Mgenero.AddItem "Acid"
'Startup.Mgenero.AddItem "House"
'Startup.Mgenero.AddItem "Game"
'Startup.Mgenero.AddItem "Sound Clip"
'Startup.Mgenero.AddItem "Gospel"
'Startup.Mgenero.AddItem "Noise"
'Startup.Mgenero.AddItem "Alt. Rock"
'Startup.Mgenero.AddItem "Bass"
'Startup.Mgenero.AddItem "Soul"
'Startup.Mgenero.AddItem "Punk"
'Startup.Mgenero.AddItem "Space"
'Startup.Mgenero.AddItem "Meditative"
'Startup.Mgenero.AddItem "Instrumental Pop"
'Startup.Mgenero.AddItem "Instrumental Rock"
'Startup.Mgenero.AddItem "Ethnic"
'Startup.Mgenero.AddItem "Gothic"
'Startup.Mgenero.AddItem "Darkwave"
'Startup.Mgenero.AddItem "Techno-Industrial"
'Startup.Mgenero.AddItem "Electronic"
'Startup.Mgenero.AddItem "Pop-Folk"
'Startup.Mgenero.AddItem "Eurodance"
'Startup.Mgenero.AddItem "Dream"
'Startup.Mgenero.AddItem "Southern Rock"
'Startup.Mgenero.AddItem "Comedy"
'Startup.Mgenero.AddItem "Cult"
'Startup.Mgenero.AddItem "Gangsta Rap"
'Startup.Mgenero.AddItem "Top 40"
'Startup.Mgenero.AddItem "Christian Rap"
'Startup.Mgenero.AddItem "Pop/Punk"
'Startup.Mgenero.AddItem "Jungle"
'Startup.Mgenero.AddItem "Native American"
'Startup.Mgenero.AddItem "Cabaret"
'Startup.Mgenero.AddItem "New Wave"
'Startup.Mgenero.AddItem "Phychedelic"
'Startup.Mgenero.AddItem "Rave"
'Startup.Mgenero.AddItem "Showtunes"
'Startup.Mgenero.AddItem "Trailer"
'Startup.Mgenero.AddItem "Lo-Fi"
'Startup.Mgenero.AddItem "Tribal"
'Startup.Mgenero.AddItem "Acid Punk"
'Startup.Mgenero.AddItem "Acid Jazz"
'Startup.Mgenero.AddItem "Polka"
'Startup.Mgenero.AddItem "Retro"
'Startup.Mgenero.AddItem "Musical"
'Startup.Mgenero.AddItem "Rock & Roll"
'Startup.Mgenero.AddItem "Hard Rock"
'Startup.Mgenero.AddItem "Folk"
'Startup.Mgenero.AddItem "Folk/Rock"
'Startup.Mgenero.AddItem "National Folk"
'Startup.Mgenero.AddItem "Swing"
'Startup.Mgenero.AddItem "Fast-Fusion"
'Startup.Mgenero.AddItem "Bebob"
'Startup.Mgenero.AddItem "Latin"
'Startup.Mgenero.AddItem "Revival"
'Startup.Mgenero.AddItem "Celtic"
'Startup.Mgenero.AddItem "Blue Grass"
'Startup.Mgenero.AddItem "Avantegarde"
'Startup.Mgenero.AddItem "Gothic Rock"
'Startup.Mgenero.AddItem "Progressive Rock"
'Startup.Mgenero.AddItem "Psychedelic Rock"
'Startup.Mgenero.AddItem "Symphonic Rock"
'Startup.Mgenero.AddItem "Slow Rock"
'Startup.Mgenero.AddItem "Big Band"
'Startup.Mgenero.AddItem "Chorus"
'Startup.Mgenero.AddItem "Easy Listening"
'Startup.Mgenero.AddItem "Acoustic"
'Startup.Mgenero.AddItem "Humour"
'Startup.Mgenero.AddItem "Speech"
'Startup.Mgenero.AddItem "Chanson"
'Startup.Mgenero.AddItem "Opera"
'Startup.Mgenero.AddItem "Chamber Music"
'Startup.Mgenero.AddItem "Sonata"
'Startup.Mgenero.AddItem "Symphony"
'Startup.Mgenero.AddItem "Booty Bass"
'Startup.Mgenero.AddItem "Primus"
'Startup.Mgenero.AddItem "Porn Groove"
'Startup.Mgenero.AddItem "Satire"
'Startup.Mgenero.AddItem "Slow Jam"
'Startup.Mgenero.AddItem "Club"
'Startup.Mgenero.AddItem "Tango"
'Startup.Mgenero.AddItem "Samba"
'Startup.Mgenero.AddItem "Folklore"
'Startup.Mgenero.AddItem "Ballad"
'Startup.Mgenero.AddItem "Power Ballad"
'Startup.Mgenero.AddItem "Rhythmic Soul"
'Startup.Mgenero.AddItem "Freestyle"
'Startup.Mgenero.AddItem "Duet"
'Startup.Mgenero.AddItem "Punk Rock"
'Startup.Mgenero.AddItem "Drum Solo"
'Startup.Mgenero.AddItem "A Capella"
'Startup.Mgenero.AddItem "Euro-House"
'Startup.Mgenero.AddItem "Dance Hall"
'Startup.Mgenero.AddItem "Goa"
'Startup.Mgenero.AddItem "Drum & Bass"
'Startup.Mgenero.AddItem "Club-House"
'Startup.Mgenero.AddItem "Hardcore"
'Startup.Mgenero.AddItem "Terror"
'Startup.Mgenero.AddItem "indie"
'Startup.Mgenero.AddItem "Brit Pop"
'Startup.Mgenero.AddItem "Negerpunk"
'Startup.Mgenero.AddItem "Polsk Punk"
'Startup.Mgenero.AddItem "Beat"
'Startup.Mgenero.AddItem "Christian Gangsta Rap"
'Startup.Mgenero.AddItem "Heavy Metal"
'Startup.Mgenero.AddItem "Black Metal"
'Startup.Mgenero.AddItem "Crossover"
'Startup.Mgenero.AddItem "Comteporary Christian"
'Startup.Mgenero.AddItem "Christian Rock"
'Startup.Mgenero.AddItem "Merengue"
'Startup.Mgenero.AddItem "Salsa"
'Startup.Mgenero.AddItem "Trash Metal"
'Startup.Mgenero.AddItem "Anime"
'Startup.Mgenero.AddItem "JPop"
'Startup.Mgenero.AddItem "Synth Pop"

'Startup.Mgenero.Text = Startup.Mgenero.List(0)    ' Display first item.

End Sub

Public Function GetMP3Tag(sPassFileName As String) As Boolean

Dim iFreefile As Integer
Dim LFilePos As Long
Dim sData As String * 128
Dim sGenreMatrix As String
Dim sGenre() As String
Dim Test1
Dim Test2

' Genre          'TAG Alternative = sGenre 20 / Blues = sGenre 0 / etc.
sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
"Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
"Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
"Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
"House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
"Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
"Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
"Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
"Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
"Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
"Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
"Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
"Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
"Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
"Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
"Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
"Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
' Build the Genre array (VB6+ only)
sGenre = Split(sGenreMatrix, "|")
    
' Clear the info variables
MP3Info.sTitle = ""
MP3Info.sArtist = ""
MP3Info.sAlbum = ""
MP3Info.sYear = ""
MP3Info.sComment = ""
' Ensure the MP3 file exists
If Dir(sPassFileName) = "" Then
    GetMP3Tag = False
    GoTo CloseMe
End If
' Retrieve the info data from the MP3
GetMP3Tag = True
iFreefile = FreeFile
LFilePos = fileLen(sPassFileName) - 127
    
On Error GoTo Mp3TagErr
Open sPassFileName For Binary As #iFreefile
Get #iFreefile, LFilePos, sData
Close #iFreefile
    
' Populate the info variables
If Left(sData, 3) = "TAG" Then
    MP3Info.sTitle = RTrim(Mid(sData, 4, 30))
    MP3Info.sArtist = RTrim(Mid(sData, 34, 30))
    MP3Info.sAlbum = RTrim(Mid(sData, 64, 30))
    MP3Info.sYear = RTrim(Mid(sData, 94, 4))
    MP3Info.sComment = RTrim(Mid(sData, 98, 30))
    MP3Info.sGenre = RTrim(sGenre(Asc(Mid(sData, 128, 1))))
End If
Exit Function
    
CloseMe:
Close #iFreefile
Exit Function

Mp3TagErr:
Resume CloseMe
Exit Function

End Function

''''''Read MP3 Header BEGIN''''''
Public Function ReadMP3Header(sPassFileName As String)

Dim z, i
Dim BinaryString As String
Dim byteArray(4) As Byte    'array that store first four bytes
Dim bin As String           'string that store binary number converted from readed bytes
Dim BinString As String     'containing binary string
Dim DecString As Integer  'containing decimal extracted from BinString
Dim FreeMp3
'''''''''''''''end of declarations'''''''

FreeMp3 = FreeFile
On Error GoTo Mp3FileErr
Open sPassFileName For Binary Access Read As #FreeMp3  'open file for read
For z = 1 To 4                           'step through four bytes
    Get #FreeMp3, z, byteArray(z)                  'store every(z)byte  in array position z
Next z                                   'back for next byte
Close #FreeMp3                                   'close file

bin = ""                                   'reset and build the desired binary number in this string
For z = 1 To 4                           'convert all bytes to binary
    For i = 0 To 7 Step 1                  'Here comes the decimal=>binary conversion
        If byteArray(z) And (2 ^ i) Then   'Use the logical "AND" operator.
            bin = bin + "1"
        Else
            bin = bin + "0"
        End If
    Next i                             'End of binary conversion
Next z
BinaryString = bin

'''''''''check MP3HeaderInfo.Frequency''''
DecString = 0
BinString = Mid(bin, 19, 2)         'take 19 to 21

For i = 1 To Len(BinString)         'convert to decimal
    If Mid(BinString, i, 1) = 1 Then
        DecString = DecString + 2 ^ (Len(BinString) - i)
    End If
Next i

Select Case DecString
    Case 0
        MP3HInfo.Frequency = 44100
    Case 1
        MP3HInfo.Frequency = 32000
    Case 2
        MP3HInfo.Frequency = 48000
    Case 3
        'xxxx
End Select

'''''check MP3HeaderInfo.Layer''''
DecString = 0
BinString = Mid(bin, 10, 2)

For i = 1 To Len(BinString)
    If Mid(BinString, i, 1) = 1 Then
        DecString = DecString + 2 ^ (Len(BinString) - i)
    End If
Next i

Select Case DecString
    Case 0
        MP3HInfo.Layer = ""
    Case 1
        MP3HInfo.Layer = 2
    Case 2
        MP3HInfo.Layer = 3
    Case 3
        MP3HInfo.Layer = 1
End Select

''''check MP3HeaderInfo.Mode''''
DecString = 0
BinString = Mid(bin, 31, 2)

For i = 1 To Len(BinString)
    If Mid(BinString, i, 1) = 1 Then
        DecString = DecString + 2 ^ (Len(BinString) - i)
    End If
Next i

Select Case DecString
    Case 0
        MP3HInfo.mode = "Stereo"
    Case 1
        MP3HInfo.mode = "Dual Channel"
    Case 2
        MP3HInfo.mode = "Joint stereo"
    Case 3
        MP3HInfo.mode = "Mono"
End Select

''''check MP3HeaderInfo.MpegVersion
If Mid(bin, 12, 1) = 0 Then
    MP3HInfo.MpegVersion = 2
Else
    MP3HInfo.MpegVersion = 1
End If

'''''check MP3HeaderInfo.Bitrate''''
DecString = 0
BinString = Mid(bin, 21, 4)

For i = 1 To Len(BinString)
    If Mid(BinString, i, 1) = 1 Then
        DecString = DecString + 2 ^ (Len(BinString) - i)
    End If
Next i

Select Case DecString
    Case 0
        MP3HInfo.Bitrate = 0
    Case 1
        MP3HInfo.Bitrate = 112
    Case 2
        MP3HInfo.Bitrate = 56
    Case 3
        MP3HInfo.Bitrate = 224
    Case 4
        MP3HInfo.Bitrate = 40
    Case 5
        MP3HInfo.Bitrate = 160
    Case 6
        MP3HInfo.Bitrate = 80
    Case 7
        MP3HInfo.Bitrate = 320
    Case 8
        MP3HInfo.Bitrate = 32
    Case 9
        MP3HInfo.Bitrate = 128
    Case 10
        MP3HInfo.Bitrate = 64
    Case 11
        MP3HInfo.Bitrate = 256
    Case 12
        MP3HInfo.Bitrate = 48
    Case 13
        MP3HInfo.Bitrate = 192
    Case 14
        MP3HInfo.Bitrate = 96
    Case 15
        MP3HInfo.Bitrate = 0
        If MP3HInfo.Layer = 1 Then
            Select Case DecString
                Case 0
                    MP3HInfo.Bitrate = 0
                Case 1
                    MP3HInfo.Bitrate = 128
                Case 2
                    MP3HInfo.Bitrate = 64
                Case 3
                    MP3HInfo.Bitrate = 256
                Case 4
                    MP3HInfo.Bitrate = 48
                Case 5
                    MP3HInfo.Bitrate = 192
                Case 6
                    MP3HInfo.Bitrate = 96
                Case 7
                    MP3HInfo.Bitrate = 384
                Case 8
                    MP3HInfo.Bitrate = 32
                Case 9
                    MP3HInfo.Bitrate = 160
                Case 10
                    MP3HInfo.Bitrate = 80
                Case 11
                    MP3HInfo.Bitrate = 320
                Case 12
                    MP3HInfo.Bitrate = 56
                Case 13
                    MP3HInfo.Bitrate = 224
                Case 14
                    MP3HInfo.Bitrate = 112
                Case 15
                    MP3HInfo.Bitrate = 0
            End Select
        End If
End Select

'''''MP3HeaderInfo.Emphasis''''
DecString = 0
BinString = Mid(bin, 25, 2)

For i = 1 To Len(BinString)        'go from first
    If Mid(BinString, i, 1) = 1 Then
        DecString = DecString + 2 ^ (Len(BinString) - i)
    End If
Next i

Select Case DecString
    Case 0
        MP3HInfo.Emphasis = "No"
    Case 1
        MP3HInfo.Emphasis = "-?-"
    Case 2
        MP3HInfo.Emphasis = "50/15"
    Case 3
        MP3HInfo.Emphasis = "CITT j. 17"
End Select

With MP3HInfo
    Dim min, sec
    .Bitrate = Int(.Bitrate)
    .mFileSize = FileSizeMP3(sPassFileName)
    .FPlayTime = ((.mFileSize * 8) / (.Bitrate * 1000))
    min = .FPlayTime \ 60         'minutes
    sec = .FPlayTime - (min * 60) 'seconds
    .FPlayTime = Format(min, "#0#") & ":" & Format(sec, "0#") 'format time to 00:00
End With
Exit Function

Mp3FileErr:
MP3HInfo.Bitrate = ""
MP3HInfo.Emphasis = ""
MP3HInfo.FPlayTime = ""
MP3HInfo.Frequency = ""
MP3HInfo.Layer = ""
MP3HInfo.mFileSize = ""
MP3HInfo.mode = ""
MP3HInfo.MpegVersion = ""
Resume CloseMp3
Exit Function

CloseMp3:
End Function
''''''Read MP3 Header END''''''

''''''Remove Tag BEGIN''''''
Public Function RemoveMP3Tag(sPassFileName As String) As Boolean

Dim blank
Dim FreeMp3
Dim LFilePos As Long

blank = String$(127, 0)                     'assign string "blank" 127 blank
FreeMp3 = FreeFile
LFilePos = fileLen(sPassFileName) - 127

On Error GoTo Errorcheck
Open sPassFileName For Binary Access Write As #FreeMp3 'open file
Seek #FreeMp3, LFilePos       'seek position
Put #FreeMp3, , blank         'write string
Close #FreeMp3                'close file

RemoveMP3Tag = True
Exit Function

Errorcheck:
RemoveMP3Tag = False
Exit Function

End Function

Public Function WriteMP3Tag(sPassFileName As String, TAG As String, SongName As String, Artist As String, Album As String, Year As String, Comment As String, ByVal Genre As Long) As Boolean
 
 Dim FreeMp3
 Dim wTag As String * 3     ' First 3 Chars of 128 byte Tag Info - 'TAG'
 Dim wSongname As String * 30
 Dim wArtist As String * 30
 Dim wAlbum As String * 30
 Dim wYear As String * 4
 Dim wComment As String * 30
 Dim wGenre As String * 1
 Dim LFilePos As Long
 
    LFilePos = fileLen(sPassFileName) - 127
    
    wTag = TAG
    wSongname = Trim(Left(SongName, 30))
    wArtist = Trim(Left(Artist, 30))
    wAlbum = Trim(Left(Album, 30))
    wYear = Trim(Left(Year, 4))
    wComment = Trim(Left(Comment, 30))
    wGenre = Chr(Genre)
    
    FreeMp3 = FreeFile
    On Error GoTo Errorcheck
    Open sPassFileName For Binary Access Write As #FreeMp3
    Seek #FreeMp3, LFilePos
    Put #FreeMp3, , wTag
    Put #FreeMp3, , wSongname
    Put #FreeMp3, , wArtist
    Put #FreeMp3, , wAlbum
    Put #FreeMp3, , wYear
    Put #FreeMp3, , wComment
    Put #FreeMp3, , wGenre
    WriteMP3Tag = True
    Close #FreeMp3
Exit Function

Errorcheck:
WriteMP3Tag = False
Close #FreeMp3
Exit Function

End Function

Private Function FileSizeMP3(File As String) As String
    
    Dim LSize As String
    
    If File = "" Then
        FileSizeMP3 = ""
        Exit Function
    End If
    
    LSize = fileLen(File)
    FileSizeMP3 = LSize 'Size in bytes
    
End Function
