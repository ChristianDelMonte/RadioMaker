Attribute VB_Name = "Compress"
'/////////////////////////////////////////////////////
'* ZLib.bas                                          *
'* By: W-Buffer (Carlos Daniel Ruvalcaba Valenzuela) *
'*     Iridium Studios.                              *
'* Web: http://istudios.virtualave.net               *
'* Mail: chadruva@hotmail.com                        *
'* Thanks to: the ZLib.dll guys! :)                  *
'*                                                   *
'* NOTES: - You need to have ZLib.dll in             *
'*        your System Folder.                        *
'*        - You need to have the ZLib.dll            *
'*        Version 1.1.3.1                            *
'*        - Fell Free to do with this bas whatever   *
'*        you want (Steal, Copy, etc.)               *
'/////////////////////////////////////////////////////

'You can use this silly enumeration for Compresion Level
Public Enum CompressionLevel
    None = 0
    Poor = 1
    Fair = 2
    Average = 3
    Normal = 4
    Good = 5
    VeryGood = 6
    Best = 7
    SuperCompressed = 8
    MaxCompression = 9
End Enum

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function compress2 Lib "zlib.dll" (Dest As Any, destLen As Any, Src As Any, ByVal srcLen As Long, ByVal Level As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (Dest As Any, destLen As Any, Src As Any, ByVal srcLen As Long) As Long

'/////////////////////////////////////////
'* CompressBytes - Compress a Bytes Buffer
'* IN  - Bytes - Bytes Array
'*      Level - Compression Level to use
'* OUT - Nothing
'/////////////////////////////////////////

Public Sub CompressBytes(Bytes() As Byte, Level As Integer)
    Dim BuffSize As Long
    Dim TBuff() As Byte
    
    BuffSize = UBound(Bytes) + 1
    BuffSize = BuffSize + (BuffSize * 1.01) + 12
    ReDim TBuff(BuffSize)
    
    compress2 TBuff(0), BuffSize, Bytes(0), UBound(Bytes) + 1, Level
    
    ReDim Bytes(BuffSize - 1)
    
    CopyMemory Bytes(0), TBuff(0), BuffSize
End Sub

'////////////////////////////////////////////////////////////////
'* UnCompressBytes - Uncompresses a Byte Buffer to original size
'* IN  - Bytes - Compressed Bytes Array
'*      OriginalSize - Uncompressed size of Bytes Buffer
'* Out - Nothing
'////////////////////////////////////////////////////////////////

Public Sub UnCompressBytes(Bytes() As Byte, OriginalSize As Long)
    Dim BuffSize As Long
    Dim TBuff() As Byte
    
    BuffSize = OriginalSize
    BuffSize = BuffSize + (BuffSize * 1.01) + 12
    ReDim TBuff(BuffSize)
    
    uncompress TBuff(0), BuffSize, Bytes(0), UBound(Bytes) + 1
    
    ReDim Bytes(BuffSize - 1)
    
    CopyMemory Bytes(0), TBuff(0), BuffSize
End Sub

'////////////////////////////////////////////////////////
'* CompressFile - Compresses a File using CompressBytes
'* IN  - Src - Source File to compress
'*       Dest - Compressed Destination File
'*       Level - Compression Level To Use
'* OUT - Nothing
'///////////////////////////////////////////////////////

Public Sub CompressFile(Src As String, Dest As String, Level As Integer)
      
    Open Src For Binary Access Read As 15
    Open Dest For Binary Access Read Write As 25
    
    Dim Srcs As Long
    Srcs = LOF(15)
    ReDim buff(Srcs - 1) As Byte
    Get 15, , buff
    
    CompressBytes buff, 9
        
    Put 25, , Srcs
    Put 25, , buff
    Close
    
End Sub

'//////////////////////////////////////////////////////////////////////////////
'* UnCompressFile - UnCompresses a Compressed File (duh!) using UnCompressBytes
'* IN  - Src - Source File to UnCompress
'*       Dest - UnCompressed Destination File
'* OUT - Nothing
'//////////////////////////////////////////////////////////////////////////////

Public Sub UnCompressFile(Src As String, Dest As String)
       
    On Error GoTo Nop
    Open Src For Binary Access Read As 15
    Open Dest For Binary Access Write As 25
    
    Dim Srcs As Long
    Get 15, , Srcs
    ReDim buff(Srcs - 1) As Byte
    Get 15, , buff
        
    UnCompressBytes buff, Srcs
        
    Put 25, , buff
    
    Close 15
    Close 25
    
    Exit Sub
    
Nop:
MsgBox "Archivo Corrupto.", vbCritical
End Sub
