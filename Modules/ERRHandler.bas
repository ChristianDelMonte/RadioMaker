Attribute VB_Name = "ERRHandler"

'////////////////////////////////////////////////////////
'*
'*  ////////// ERRHANDLER module for Vb. 6+ ///////////
'*  * this module is for error handler localization  *
'*  ********* and is for Radiomaker 1.0 only ********
'*
'*    Copyright (c) 1987-2008 Only development Inc.
'*    Christian A. Del Monte
'////////////////////////////////////////////////////////

Private ERRFilename As String
Private ERRtoWrite As String

Private Function WriteErrors(ERRDescription As String, ERRDetails As String)

ERRFilename = App.path & "\" & App.EXEName & ".error.log"
ERRtoWrite = Date$ & " - " & time$ & " /> Desc: " & ERRDescription & " /> Detalle: " & ERRDetails

Open ERRFilename For Append As #13
Write #13, ERRtoWrite
Close #13

End Function

Public Function DisplayMsg(ERRDescription As String, ERRDetails As String, ERRnum As String, ShowDisplay As Boolean)

'Display error dialogues
If ShowDisplay = True Then
    MsgBox ERRDescription & vbCrLf & vbCrLf & "Codigo N: " & ERRnum & vbCrLf & vbCrLf & "Detalles:" & vbCrLf & ERRDetails, vbCritical, "ONLY Radiomaker Error"
End If

WriteErrors ERRDescription, ERRDetails

End Function

