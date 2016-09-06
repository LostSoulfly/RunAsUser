Attribute VB_Name = "modTaskSchd"
Option Explicit

Public Function CheckTaskSched(Arguments As String) As Boolean
Dim strXML As String
Dim xmlPath As String
Dim StartDate As Date
Dim strDate As String

StartDate = DateAdd("s", 1, Now)
strDate = Format$(StartDate, "yyyy-mm-ddThh:mm:ss")
'Arguments = TrimQuotes(Arguments)

AddLog "Generating XML for task scheduler.."
strXML = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-16" & Chr(34) & "?>" & vbNewLine
strXML = strXML & "<Task version=" & Chr(34) & "1.2" & Chr(34) & " xmlns=" & Chr(34) & "http://schemas.microsoft.com/windows/2004/02/mit/task" & Chr(34) & ">" & vbNewLine
strXML = strXML & "  <RegistrationInfo>" & vbNewLine
strXML = strXML & "    <Date>" & strDate & "</Date>" & vbNewLine
strXML = strXML & "    <Author>User\RNGesus</Author>" & vbNewLine
strXML = strXML & "    <URI>\" & App.EXEName & "</URI>" & vbNewLine
strXML = strXML & "  </RegistrationInfo>" & vbNewLine
strXML = strXML & "  <Triggers>" & vbNewLine
strXML = strXML & "    <BootTrigger>" & vbNewLine
strXML = strXML & "      <StartBoundary>" & strDate & "</StartBoundary>" & vbNewLine
strXML = strXML & "      <Enabled>true</Enabled>" & vbNewLine
strXML = strXML & "    </BootTrigger>" & vbNewLine
strXML = strXML & "  </Triggers>" & vbNewLine
strXML = strXML & "  <Principals>" & vbNewLine
strXML = strXML & "    <Principal id=" & Chr(34) & "Author" & Chr(34) & ">" & vbNewLine
strXML = strXML & "      <UserId>S-1-5-18</UserId>" & vbNewLine
strXML = strXML & "      <RunLevel>HighestAvailable</RunLevel>" & vbNewLine
strXML = strXML & "    </Principal>" & vbNewLine
strXML = strXML & "  </Principals>" & vbNewLine
strXML = strXML & "  <Settings>" & vbNewLine
strXML = strXML & "    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>" & vbNewLine
strXML = strXML & "    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>" & vbNewLine
strXML = strXML & "    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>" & vbNewLine
strXML = strXML & "    <AllowHardTerminate>false</AllowHardTerminate>" & vbNewLine
strXML = strXML & "    <StartWhenAvailable>true</StartWhenAvailable>" & vbNewLine
strXML = strXML & "    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>" & vbNewLine
strXML = strXML & "    <IdleSettings>" & vbNewLine
strXML = strXML & "      <StopOnIdleEnd>false</StopOnIdleEnd>" & vbNewLine
strXML = strXML & "      <RestartOnIdle>false</RestartOnIdle>" & vbNewLine
strXML = strXML & "    </IdleSettings>" & vbNewLine
strXML = strXML & "    <AllowStartOnDemand>true</AllowStartOnDemand>" & vbNewLine
strXML = strXML & "    <Enabled>true</Enabled>" & vbNewLine
strXML = strXML & "    <Hidden>false</Hidden>" & vbNewLine
strXML = strXML & "    <RunOnlyIfIdle>false</RunOnlyIfIdle>" & vbNewLine
strXML = strXML & "    <WakeToRun>false</WakeToRun>" & vbNewLine
strXML = strXML & "    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>" & vbNewLine
strXML = strXML & "    <Priority>7</Priority>" & vbNewLine
strXML = strXML & "    <RestartOnFailure>" & vbNewLine
strXML = strXML & "      <Interval>PT10M</Interval>" & vbNewLine
strXML = strXML & "      <Count>99</Count>" & vbNewLine
strXML = strXML & "    </RestartOnFailure>" & vbNewLine
strXML = strXML & "  </Settings>" & vbNewLine
strXML = strXML & "  <Actions Context=" & Chr(34) & "Author" & Chr(34) & ">" & vbNewLine
strXML = strXML & "    <Exec>" & vbNewLine

strXML = strXML & "      <Command>" & App.Path & "\" & App.EXEName & ".exe" & "</Command>" & vbNewLine 'should be encapped in quotes
strXML = strXML & "      <Arguments>" & Arguments & "</Arguments>" & vbNewLine

strXML = strXML & "    </Exec>" & vbNewLine
strXML = strXML & "  </Actions>" & vbNewLine
strXML = strXML & "</Task>"

AddLog "Path: " & App.Path & "\" & App.EXEName & ".exe", True
AddLog "Args: " & Arguments, True

xmlPath = App.Path & "\" & Int(Rnd * 9999999999#) & ".xml"

AddLog "XML output: " & xmlPath, True

WriteFile strXML, xmlPath

If Not FileExists(xmlPath) Then
    xmlPath = Environ("TEMP") & "\" & Int(Rnd * 9999999999#) & ".xml"
    WriteFile strXML, xmlPath
    AddLog "XML 2nd try output: " & xmlPath, True
End If

    ExecFile "schtasks.exe", "/CREATE /RU SYSTEM /TN " & App.EXEName & " /XML " & Chr(34) & xmlPath & Chr(34) & " /F"
    AddLog "Running: schtasks.exe" & " /CREATE /RU SYSTEM /TN " & App.EXEName & " /XML " & Chr(34) & xmlPath & Chr(34) & " /F", True
    
    If Not RunTaskSched Then
        AddLog "Unable to elevate. Try again.."
        CheckTaskSched = False
        'If RunTaskSched Then AddLog "Process elevated to system successfully (2nd try?)!"    'try again?
    Else
        AddLog "Process elevated to system successfully!"
        CheckTaskSched = True
    End If
    
    AddLog "Delete " & xmlPath, True
    DoEvents
    Kill xmlPath
    


End Function

Private Function RunTaskSched() As Boolean
    Dim fColor As Long, bColor As Long, i As Long
    Dim blSuccess As Boolean
    
'AddLog "RunTaskSched"
   
If ExecFile("schtasks.exe", "/RUN /TN " & App.EXEName) > 0 Then

    Con.CurrentX = 0
    Con.CursorVisible = False
    'Con.ForeColor = conYellowHi
    
    'bColor = Con.BackColor
    fColor = Con.ForeColor
    
    Con.ForeColor = conYellowHi
    
    Con.WriteLine "Listening for SysMutex.", False

    
    For i = 1 To 100
        If i Mod 3 = 0 Then Con.WriteLine ".", False
        If i Mod 10 = 0 Then If IsSysMutexHeld Then blSuccess = True: Exit For
    Sleep 10
    Next i
    
    Con.WriteLine ""
    Con.CursorVisible = True
    
    'For i = 1 To 5000
    '' Do a small part of a large task.
    '    If i Mod 2 = 0 Then
    '        n = n + 1
    '        Con.WriteLine Mid$(Twirl, (n Mod 4) + 1, 1), False
    '        Con.CurrentX = Con.CurrentX - 1
    '            If i Mod 3 = 0 Then
    '            Con.WriteLine ".", False
    '            End If
    '            If i Mod 50 = 0 Then
    '
    '            End If
    '    End If
    '    If Con.Break Then
    '        Con.WriteLine vbLf
    '        Con.CursorVisible = True
    '        Exit For
    '    End If
    '    Sleep 1
    'Next i
Else
    'fail?
    AddLog "Problem scheduling the task!?"
End If

Con.ForeColor = fColor  'reset the fore color
If blSuccess Then AddLog "System mutex reserved!"

RemoveTaskSched

RunTaskSched = blSuccess

End Function

Private Sub RemoveTaskSched()

    AddLog "Removing scheduled task.."
    ExecFile "schtasks.exe", "/delete /TN " & App.EXEName & " /F"
    
End Sub
