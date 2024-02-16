Attribute VB_Name = "ModSum"
'////////////////////////////////////////////////////
'*
'*  // TIME managger/calculation module for Vb.6+ //
'*  ******** module for Radiomaker 1.0 only ********
'*  Copyright (c) 1987-2002 Only development Inc.
'*  Christian A. Del Monte
'///////////////////////////////////////////////////

Option Explicit

'////////////////////////////////////////////////////
'* Convert and format seconds into hh:mm:ss
'///////////////////////////////////////////////////
Public Function ConvSecToMin(ByVal Seconds As Long) As String
    
Dim Tmp As String
Dim hour As Single, Min As Single, sec As Single

On Error GoTo errGetMinutes
Seconds = IIf(Seconds < 0, 0, Seconds)
  
Tmp = Seconds / 60 / 60
hour = Int(Tmp)
Tmp = Seconds - ((hour * 60) * 60)
Min = Int(Tmp / 60)
sec = Tmp - (Min * 60)
  
If hour >= 24 Then hour = hour - 24
    
ConvSecToMin = Format(hour, "00") & ":" & Format(Min, "00") & ":" & Format(sec, "00")
Exit Function

errGetMinutes:
  ConvSecToMin = "00:00:00"

End Function

'////////////////////////////////////////////////////
'* Convert and deformat hh:mm:ss into seconds
'///////////////////////////////////////////////////
Public Function ConvMinToSec(WNumber As String) As Single

Dim Hours, Mins, Segs, Milisegs As Long
Dim Total As Single
Dim FFormat As String
Dim LnMins As Integer

On Error GoTo errGetSeconds

LnMins = Len(WNumber)
Select Case LnMins
    Case 10 'hh:mm:ss.s
        Hours = CLng(Left$(WNumber, 2))
        Mins = CLng(Mid$(WNumber, 4, 2))
        Segs = CLng(Mid$(WNumber, 7, 2))
        Milisegs = CLng(Right$(WNumber, 1))
        If Hours = 0 Then
            Hours = Hours
        Else
            Hours = Hours * 60
        End If
        Mins = Mins + Hours
        If Mins = 0 Then
            Mins = Mins
        Else
            Mins = Mins * 60
        End If
        
        Total = Mins + Segs
        FFormat = Str$(Total) & "." & Str$(Milisegs)
        Total = CSng(FFormat)
        ConvMinToSec = Total
    
    Case 8  'hh:mm:ss
        Hours = CLng(Left$(WNumber, 2))
        Mins = CLng(Mid$(WNumber, 4, 2))
        Segs = CLng(Right$(WNumber, 2))
        If Hours = 0 Then
            Hours = Hours
        Else
            Hours = Hours * 60
        End If
        Mins = Mins + Hours
        If Mins = 0 Then
            Mins = Mins
        Else
            Mins = Mins * 60
        End If
        Total = Mins + Segs
        ConvMinToSec = Total
        
    Case 5  'hh:mm
        Mins = CLng(Left$(WNumber, 2))
        Segs = CLng(Right$(WNumber, 2))
        If Mins = 0 Then
            Mins = Mins
        Else
            Mins = Mins * 60
        End If
        Total = Mins + Segs
        ConvMinToSec = Total
    
    Case Else
        Total = 0
        ConvMinToSec = Total
End Select

Exit Function

errGetSeconds:
    ConvMinToSec = 0
    
End Function
