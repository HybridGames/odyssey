Attribute VB_Name = "modDebug"
'Simple method now designed to write out to a static log file
Public Sub WriteLog(log As String)
    Open "debug.log" For Output As #1
    
    Print #1, log
    
    Close #1
End Sub

'Allows for logging to display, as well as writing to the log
Public Sub DisplayLog(log As String)
    WriteLog log
    PrintChat log, YELLOW
End Sub
