Attribute VB_Name = "modDebug"
'Simple method now designed to write out to a static log file
Public Sub writeLog(log As String)
    Dim LogFile As Integer
    LogFile = FreeFile
    
    Open "debug.log" For Input As LogFile
    
    Print LogFile, log
    
    Close LogFile
End Sub
