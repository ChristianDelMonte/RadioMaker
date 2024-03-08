Attribute VB_Name = "BASSWMA"
'////////////////////////////////////////////////////////////////
' BASSWMA.BAS - Visual Basic API Header File
'       Copyright (c) 2002 JOBnik! [Arthur Aminov, ISRAEL]
'                          e-mail: jobnik2k@hotmail.com
'
' Originally translated from - basswma.h - C/C++ Header file
'
' Requires BASS.DLL & BASS.BAS 1.6 - available @ www.un4seen.com
' --------------------------------------------------------------
' See the BASSWMA.CHM file for more complete documentation
'////////////////////////////////////////////////////////////////

' Additional error codes returned by BASS_WMA_GetErrorCode
Global Const BASS_ERROR_WMA_LICENSE = 1000     ' the file is protected

' Additional flags for use with BASS_WMA_EncodeOpenFile/Network
Global Const BASS_WMA_ENCODE_TAGS = &H10000    ' set tags in the WMA encoding


Declare Function BASS_WMA_ErrorGetCode Lib "basswma.dll" () As Long
'Get the BASS_ERROR_xxx error code. Use this function to get the
'reason for an error.

Declare Function BASS_WMA_StreamCreateFile Lib "basswma.dll" (ByVal mem As Long, ByVal file As Any, ByVal offset As Long, ByVal length As Long, ByVal flags As Long) As Long
'Create a sample stream from a WMA file (or URL).
'mem    : TRUE = Stream file from memory
'file   : Filename (mem=FALSE) or memory location (mem=TRUE)
'offset : ignored (set to 0)
'length : Data length (only used if mem=TRUE)
'flags  : flags (BASS_SAMPLE_LOOP / BASS_SAMPLE_3D / BASS_SAMPLE_FX / BASS_STREAM_DECODE)
'RETURN : The created stream's handle (0=error)

Declare Sub BASS_WMA_StreamFree Lib "basswma.dll" (ByVal handle As Long)
'Free a WMA stream's resources.
'handle : Stream handle

Declare Function BASS_WMA_StreamGetLength Lib "basswma.dll" (ByVal handle As Long) As Long
'Retrieve the playback length (in bytes) of a WMA stream.
'handle : stream handle
'RETURN : The length (0xffffffff=error)

Declare Function BASS_WMA_StreamGetTags Lib "basswma.dll" (ByVal handle As Long, ByVal tags As Long) As Long
'Retrieve the WMA tags, if available.
'handle : stream handle
'tags   : ignored
'RETURN : Pointer to the tags (0=error)

Declare Function BASS_WMA_StreamPlay Lib "basswma.dll" (ByVal handle As Long, ByVal flush As Long, ByVal flags As Long) As Long
'Play a WMA stream.
'handle : Handle of stream to play
'flush  : TRUE=restart from the beginning.
'flags  : BASS_SAMPLE_LOOP flag

Declare Function BASS_WMA_ChannelSetPosition Lib "basswma.dll" (ByVal handle As Long, ByVal pos As Long) As Long
'Set the current playback position of a WMA channel.
'handle : channel handle
'pos    : the position (in bytes)

Declare Function BASS_WMA_ChannelGetPosition Lib "basswma.dll" (ByVal handle As Long) As Long
'Get the current playback position of a WMA channel.
'handle : channel handle
'RETURN : the position in bytes (0xffffffff=error)

Declare Function BASS_WMA_ChannelSetSync Lib "basswma.dll" (ByVal handle As Long, ByVal atype As Long, ByVal param As Long, ByVal proc As Long, ByVal user As Long) As Long
'Setup a sync on a WMA channel. Multiple syncs may be used per channel.
'handle : Channel handle
'atype  : Sync type (BASS_SYNC_xxx type & flags)
'param  : Sync parameters (see the BASS_SYNC_xxx type description)
'proc   : User defined callback function
'user   : The 'user' value passed to the callback function
'RETURN : Sync handle (0=error)

Declare Function BASS_WMA_ChannelRemoveSync Lib "basswma.dll" (ByVal handle As Long, ByVal sync As Long) As Long
'Remove a sync from a WMA channel
'handle : Channel handle
'sync   : Handle of sync to remove

Declare Function BASS_WMA_GetIWMReader Lib "basswma.dll" (ByVal handle As Long) As Long
'Retrieve the IWMReader interface of a WMA stream. This allows direct
'access to the WMFSDK functions.
'handle : channel handle
'RETURN : Pointer to the IWMReader object interface (0=error)


Declare Function BASS_WMA_EncodeGetRates Lib "basswma.dll" (ByVal freq As Long, ByVal flags As Long) As Long
'Retrieve a list of the encoding bitrates available for a
'specified input sample format.
'freq   : Sampling rate
'flags  : BASS_SAMPLE_MONO flag
'RETURN : Pointer to an array of bitrates, terminated by a 0 (0=error)

Declare Function BASS_WMA_EncodeOpenFile Lib "basswma.dll" (ByVal freq As Long, ByVal flags As Long, ByVal bitrate As Long, ByVal file As Any) As Long
'Initialize WMA encoding to a file.
'freq   : Sampling rate
'flags  : BASS_SAMPLE_MONO/BASS_WMA_ENCODE_TAGS flags
'bitrate: Encoding bitrate
'file   : Filename
'RETURN : The created encoder's handle (0=error)

Declare Function BASS_WMA_EncodeOpenNetwork Lib "basswma.dll" (ByVal freq As Long, ByVal flags As Long, ByVal bitrate As Long, ByVal port As Long, ByVal clients As Long) As Long
'Initialize WMA encoding to the network.
'freq   : Sampling rate
'flags  : BASS_SAMPLE_MONO/BASS_WMA_ENCODE_TAGS flags
'bitrate: Encoding bitrate
'port   : Port number for clients to conenct to (0=let system choose)
'clients: Maximum number of clients that can connect
'RETURN : The created encoder's handle (0=error)

Declare Function BASS_WMA_EncodeGetPort Lib "basswma.dll" (ByVal handle As Long) As Long
'Retrieve the port for clients to connect to a network encoder.
'handle : Encoder handle
'RETURN : The port number for clients to connect to (0=error)

Declare Function BASS_WMA_EncodeGetClients Lib "basswma.dll" (ByVal handle As Long) As Long
'Retrieve the number of clients connected.
'handle : Encoder handle
'RETURN : The number of clients (-1=error)

Declare Function BASS_WMA_EncodeSetTag Lib "basswma.dll" (ByVal handle As Long, ByVal tag As String, ByVal text As String) As Long
'Set a tag. Requires that the BASS_WMA_ENCODE_TAGS flag was used in
'the BASS_WMA_EncodeOpenFile/Network call.
'handle : Encoder handle
'tag    : The tag (vbNullString=no more tags)
'text   : The tag's text

Declare Function BASS_WMA_EncodeWrite Lib "basswma.dll" (ByVal handle As Long, ByVal buffer As Long, ByVal length As Long) As Long
'Encode sample data and write it to the file or network.
'handle : Encoder handle
'buffer : Buffer containing the sample data
'length : Number of bytes in the buffer

Declare Sub BASS_WMA_EncodeClose Lib "basswma.dll" (ByVal handle As Long)
'Finish encoding and close the file or network port.
'handle : Encoder handle
