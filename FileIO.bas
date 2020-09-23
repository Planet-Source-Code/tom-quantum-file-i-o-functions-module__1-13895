Attribute VB_Name = "FileIO"
'****************************************************************
'* MODULE:      FileIO.bas                                      *
'* PURPOSE:     Contains various file i/o functions             *
'* REQUIRES:    Microsoft Common Dialog Control                 *
'* DESCRIPTION: A collection of file i/o functions including    *
'*              reading/writing binary and random access files, *
'*              and functions that simplify the use of the Open *
'*              and Save common dialog boxes.                   *
'****************************************************************

'FUNCTION: FileContents
'PURPOSE:  Reads the contents of a file
Public Function FileContents(Filename As String)
    Dim ff As Integer
    Dim txt As String
    
    On Error Resume Next
    ff = FreeFile
    Open Filename For Binary As #ff
    txt = Input(LOF(ff), #ff)
    Close ff
    FileContents = txt
End Function

'FUNCTION: WriteFile
'PURPOSE:  Writes text to a file
Public Sub WriteFile(Filename As String, Text As String)
    Dim ff As Integer
    
    On Error Resume Next
    ff = FreeFile
    Open Filename For Binary As #ff
    Put #ff, 1, Text
    Close ff
End Sub

'FUNCTION: ReadRAM
'PURPOSE:  Reads the content of a record in a random-access file
Public Function ReadRAM(Filename As String, Record As Long) As String
    Dim ff As Integer
    Dim strRes As String
    
    On Error Resume Next
    ff = FreeFile
    Open Filename For Random As #ff
    Get #ff, Record, strRes
    Close
    ReadRAM = strRes
End Function

'FUNCTION: WriteRAM
'PURPOSE:  Writes text into a record in a random-access file
Public Sub WriteRAM(Filename As String, Record As Long, Text As String)
    Dim ff2 As Integer
    Dim strWrite As String
    
    On Error Resume Next
    ff2 = FreeFile
    strWrite = Text
    Open Filename For Random As #ff2
    Put #ff2, Record, strWrite
    Close
End Sub

'FUNCTION: CDBFileOpen
'PURPOSE:  Gets a file from a common dialog box, and reads the contents of it
Public Function CDBFileOpen(Dialog As CommonDialog, Filter As String)
    Dim fn As String
    
    On Error Resume Next
    With Dialog
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = Filter
        .Filename = ""
        .ShowOpen
        fn = .Filename
        CDBFileOpen = FileContents(fn)
    End With
End Function

'FUNCTION: CDBFileSave
'PURPOSE:  Gets a file from a common dialog box, and writes text into it.
Public Sub CDBFileSave(Dialog As CommonDialog, Filter As String, Text As String)
    Dim fn As String
    
    On Error Resume Next
    With Dialog
        .CancelError = True
        .DialogTitle = "Save"
        .Filter = Filter
        .Filename = ""
        .ShowOpen
        fn = .Filename
        WriteFile fn, Text
    End With
End Sub
