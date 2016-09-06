Attribute VB_Name = "modMain"
Option Explicit

'Basic rogram process:

'Check if running as SYSTEM.
'If not, check if running elevated (UAC elevation)
'If not, log to file and exit.
'Attempt to run self as SYSTEM using task scheduler and pass the command line args
'New instance (hopefully) starts and captures a global mutex
'Old instance is monitoring the new mutex, and when it's taken it closes itself after removing the scheduled task
'If old instance doesn't detect a new mutex for the new, elevated instance, it tries again
'If it doesn't work after a few tries, it logs this to a file and exits.
'New instance verifies its command line args and runs them.

'Cmd args:
'-cu Copy as the user, this would allow you to utilize their network permissions when copying from network shares

Sub Main()

'setup the console
If Con.Initialize = conLaunchExplorer And LenB(Command$) = 0 Then
    'display a form or messagebox
    Con.Visible = False
    MsgBox "You must run this program from a command line."
    End
End If
Con.WriteLine "RunAsActiveUser - (" & getCurrentUser & ") v" & App.Major & "." & App.Minor & "." & App.Revision & " by Bradley S."

If LenB(Command$) = 0 Then DoShowHelp: End

MutexBaseName = "Global\" & App.EXEName 'To have more than one Mutex, the simplest way I could think of was to use the App's EXEname.
                                        'Basically, change the name if you want to use multiple instances on the same computers.
                                        'Alternatively, we could generate one and pass it along the chain, but that's more work!
                                        '(feel free to add that if you want to!)

ParseCommandArgs    'Parse/read our args, determine if we're doing a msgbox, copying a file, running a file..

AddLog "MutexBaseName: " & MutexBaseName, True
AddLog "Command line args: " & Command$

If TerminateRunAsInstances Then AddLog "Removing All Instances with TaskKill..": DoEvents: ExecFile "taskkill.exe", "/F /IM " & App.EXEName & "*": End

If Launch Then
    If ShowMessageBox Then
        DoMessageBox
    Else
        ProcessCommands2
    End If
    DoEvents
    End
End If

If isSystem Then
    'Attempt to grab the System mutex
    'pause for a second or two to kill previous instance
    If Not IsSysMutexHeld Then GetSysMutex Else AddLog "SysMutex held. Exiting..": End
    DoEvents
    Sleep 100
    ProcessCommands2
    
ElseIf IsAdmin Then
    'attempt to grab the admin mutex
    If Not IsAdminMutexHeld Then GetAdminMutex Else AddLog "AdminMutex held. Exiting..": End
    
    CheckRunAsSystem 'attempt to schedule a task to elevate to system privs
Else
    If Not IsUserMutexHeld Then GetUserMutex Else AddLog "UserMutex held. Exiting..": End
    'attempt to grab the User mutex
    If ShowMessageBox Then
        DoMessageBox
    Else
        AddLog "Not running elevated. No commands to process."
    End If
    DoEvents
    End
End If


CloseAllMutex   'mainly do this if the program is run from the IDE! Otherwise, VB holds the mutex until restarted.

End Sub

Private Sub CheckSafeToProceed()

    If Not SafeToProceed Then AddLog "Unable to proceed, there is likely an issue with your command arguments." _
    & vbCrLf & "To bypass this message, use -b": End

End Sub

Private Sub ProcessCommands2()
On Error GoTo oops
Dim lngSession As Long
Dim blStopProcessing As Boolean
    AddLog "Processing Commands 2.."
    
    CheckSafeToProceed
    
    AddLog "Checking Copy/Download..", True
    'check for Download/Copying file to destination
    If LenB(CopyTo) > 0 And LenB(FileToRun) > 0 Then DoCopy FileToRun, CopyTo, DownloadURL

    AddLog "Checking if waiting for specific user..", True
If PermanentlyRun And Not isSystem Then AddLog "Not SYSTEM. Can't -p.": End

If Not Launch Then
    If PermanentlyRun Then
        If WaitForUser Or LenB(RunForSpecificUser) > 0 Then
            lngSession = DoWaitForUser(RunForSpecificUser)
            
            Do While True
                If getSession <> lngSession Then
                    DoWaitForUser RunForSpecificUser, lngSession

                End If
                Sleep 5000
            Loop
            
        Else
            lngSession = 0      'set the default session to the Windows session
            Do While True
                If getSession <> lngSession Then
                    'Session has changed!
                    lngSession = getSession
                    If ShowMessageBox = True Then   'we need to run ourself and pass the old command$
                        AddLog "ShowMessageBox: True"
                        RunAsSessionId App.Path & "\" & App.EXEName & ".exe", AddToArgs("-l", Command$), lngSession
                    End If
                    
                    If LenB(FileToRun) > 0 And Not ExecuteAsCommand Then
                        AddLog "Running " & FileToRun & " " & FileArguments
                        RunAsSessionId FileToRun, FileArguments, lngSession, HideFile
                        DoEvents
                    Else
                        AddLog "Running command " & FileToRun & " " & FileArguments
                        RunAsSessionId App.Path & "\" & App.EXEName & ".exe", AddToArgs("-l", Command$), lngSession
                    End If
                    
                End If
                Sleep 10000
            Loop
        End If
    End If
    
    If LenB(RunForSpecificUser) > 0 Or WaitForUser Then DoWaitForUser RunForSpecificUser: End

    AddLog "Checking for MsgBox", True
    If ShowMessageBox = True Then   'we need to run ourself and pass the old command$
        AddLog "ShowMessageBox: True"
        RunAsSessionId App.Path & "\" & App.EXEName & ".exe", AddToArgs("-l", Command$), getSession
    End If
    
    AddLog "Checking RunForAllSessions..", True
    If RunForAllSessions Then
    Dim strLong As Variant
        For Each strLong In getSessionArray
            If ExecuteAsCommand Then
                AddLog "Running command " & FileToRun & " " & FileArguments & " user: " & GetSessionName(CLng(strLong))
                RunAsSessionId App.Path & "\" & App.EXEName & ".exe", AddToArgs("-l", Command$), CLng(strLong)
            Else
                AddLog "Running " & FileToRun & " " & FileArguments & " user: " & GetSessionName(CLng(strLong))
                RunAsSessionId FileToRun, FileArguments, CLng(strLong), HideFile
                DoEvents
            End If
        Next
        End
    End If
    
    AddLog "Checking for ExecAsCmd", True
    If ExecuteAsCommand Then
        AddLog "Running command " & FileToRun & " " & FileArguments
        RunAsSessionId App.Path & "\" & App.EXEName & ".exe", AddToArgs("-l", Command$), getSession
    Else
        AddLog "Running " & FileToRun & " " & FileArguments
        RunAsSessionId FileToRun, FileArguments, getSession, HideFile
        DoEvents
    End If
    
Else
    '-e was passed and we're not system. Exec the file!
    AddLog "Checking for ELSE Cmd", True
    If LenB(FileToRun) > 0 Then
        AddLog "Running command " & FileToRun & " " & FileArguments
        ExecFile FileToRun, FileArguments, , , CLng(Abs(HideFile))
    End If

End If

Exit Sub
oops:

AddLog "ProcessCommands2: " & Err.Number & " " & Err.Description
Resume Next

End Sub

Private Function DoWaitForUser(UserName As String, Optional LastSession As Long) As Long
Dim lngSession As Long

AddLog "Starting DoWaitForUser", True

If RunForAllSessions Then AddLog "Error: Can't -all with -p, -u, -w.": End

If LenB(UserName) = 0 And WaitForUser Then
    'wait for any user
    lngSession = getActiveSession
    
    If lngSession <> LastSession Then
        'do stuff
        AddLog "Active session detected: " & lngSession & "!"
        RunAsSessionId FileToRun, FileArguments, lngSession, HideFile
        DoEvents
    ElseIf lngSession = 0 Then AddLog "Waiting for any active session.."
        Do While lngSession = 0
            AddLog "No active session. Sleeping..", True
            Sleep 5000
            lngSession = getActiveSession
        Loop
        
        AddLog "Active session detected: " & lngSession & "!"
        RunAsSessionId FileToRun, FileArguments, lngSession, HideFile
        DoEvents
    End If
    
   
ElseIf LenB(UserName) > 0 And WaitForUser Then
    'wait for a specific user
    lngSession = getUserSession(UserName)
    If lngSession = -1 Then AddLog "Waiting for " & UserName & ".."
    Do While lngSession = -1
        AddLog "User not found. Sleeping..", True
        Sleep 5000
        lngSession = getUserSession(UserName)
    Loop
    
    AddLog "Active session of " & UserName & " detected: " & lngSession & "!"
    RunAsSessionId FileToRun, FileArguments, lngSession, HideFile
    DoEvents
    
ElseIf LenB(UserName) > 0 Then
    'wait for a specific user
    lngSession = getUserSession(UserName)
    If lngSession = -1 Then AddLog UserName & " not found, not waiting..": End
    AddLog "Active session of " & UserName & " detected: " & lngSession & "!"
    RunAsSessionId FileToRun, FileArguments, lngSession, HideFile
    DoEvents

Else
    AddLog "DoWaitForUser Else", True
    
End If

DoWaitForUser = lngSession

End Function

Private Function DoCopy(Src As String, Dest As String, Optional URL As String) As Boolean
On Error GoTo oops
Dim CopyDirectory As String

If LenB(URL) = 0 Then
    CopyDirectory = PreviousDir(Dest) 'get the path we're writing the file to
    AddLog "CopyDirectory set to: " & CopyDirectory, True
    If Not DirExists(CopyDirectory) Then CreateFolder (CopyDirectory)
    Dest = Dest & FileFromDir(Src)
    
    If FileExists(Dest) Then
        'the file already exists?
        AddLog "File already exists! Attempting to kill file.."
        Kill Dest
        If FileExists(Dest) And KillBeforeCopy Then
            ExecFile "taskkill.exe", "/F /IM " & LCase$(FileFromDir(Src))
            AddLog "Attempting to kill processes of " & LCase$(FileFromDir(Src)), True
            DoEvents
            Sleep 250
            Kill Dest
        End If
    End If
    
    If DirExists(CopyDirectory) And Not FileExists(Dest) Then
        AddLog "Copying " & Src & " to " & Dest & ".."
        FileCopy Src, Dest
        If FileExists(Dest) Then
            FileToRun = Dest      'set the new filetorun as the copy destination
            AddLog "File copy successfull!"
            DoCopy = True
            Exit Function
        Else
            AddLog "File didn't copy correctly. Permissions issue?"
            DoEvents
            End
        End If
    Else
        AddLog "Unable to copy (Dir doesn't exist/File still exists): " & Src & " to " & Dest
        DoEvents
        End
    End If
Else
    CopyDirectory = PreviousDir(Dest)
    'download URL exists. Download the sucker!
    FileToRun = Dest
    AddLog "Attempting to download " & URL
    AddLog "To: " & Dest
        If DownloadFile(URL, Dest) Then
            AddLog "Download successfull!"
            DoCopy = True
        Else
            AddLog "Unable to download file."
            DoEvents
            End
        End If
End If

Exit Function
oops:

'select case when more errors are found, not really tested yet!
If Err.Number = 76 Then
    AddLog "DoCopy: Source or destination file not found!"
ElseIf Err.Number = 70 Then
    AddLog "DoCopy: Destination file exists and is in use!"
ElseIf Err.Number = 75 Then
    AddLog "DoCopy: Error 75: Path/File access error!"
    Resume Next
Else
    AddLog "DoCopy: " & Err.Number & " " & Err.Description
End If

End
End Function

Private Sub CheckRunAsSystem()
Dim i As Integer

CheckSafeToProceed

    For i = 1 To 5
        If CheckTaskSched(Command$) Then Exit For
        Sleep 1000
    Next i

If i = 10 Then
    Con.ForeColor = conRedHi
    AddLog "Unable to elevate to system. May need a new elevation method? Try PSEXEC."
End If

End Sub

Private Sub DoMessageBox()

    AddLog "Displaying messagebox.."
    
    If isSystem Then AddLog "Running as SYSTEM, not showing MessageBox.": DoEvents: End
    MsgBox MessageBoxText, MessageBoxStyle, MessageBoxTitle
    

End Sub

Public Sub DoShowHelp()
    Dim fColor As Long
    Dim sCaption As String

    fColor = Con.ForeColor

Con.WriteLine "RunAsActiveUser Help Info" & vbCrLf

Con.WriteLine "<File Path>" & vbTab & "First arg; path to file. Supports UNC paths."
Con.WriteLine "-f <file>" & vbTab & "<F>ile name or path to file."
Con.WriteLine "-c <dir>" & vbTab & "Copy <file> to <dir> before running."
Con.WriteLine "-k" & vbTab & vbTab & "<K>ill running processes that match the destination filename"
Con.WriteLine vbTab & vbTab & "if the destination file exists and is not writable."
Con.WriteLine "-i ""URL""" & vbTab & "URL to download from, saved to -c\-f"
Con.WriteLine "-a ""args""" & vbTab & "File <A>rguments passed to <file>. If you must use quotes,"
Con.WriteLine vbTab & vbTab & "then you must use -a2 instead."
Con.WriteLine "-a2 <args>" & vbTab & "EVERYTHING after this switch is passed. This must be last!"
Con.WriteLine "-e " & vbTab & vbTab & "<E>xecute as command instead of a file."
Con.WriteLine "-h " & vbTab & vbTab & "Attempt to <H>ide the launched file (minimized)."
Con.WriteLine "-w " & vbTab & vbTab & "<W>ait until there is an active session before proceeding."
Con.WriteLine "-u <user>" & vbTab & "Only execute the command for specified <U>ser."
Con.WriteLine "-p " & vbTab & vbTab & "<P>ermanently run as SYSTEM, executing the desired"
Con.WriteLine vbTab & vbTab & "command when session changes. Does not survive PC restarts!"
Con.WriteLine "-all" & vbTab & vbTab & "Run the command for <ALL> current sessions."
Con.WriteLine "-t" & vbTab & vbTab & "<t>erminate all RunAsUser*.exe currently running processes."
Con.WriteLine
Con.WriteLine "-m ""Text""" & vbTab & "Display a <M>essageBox with supplied Text."
Con.WriteLine "-mt ""Title""" & vbTab & "Sets the Title of the MessageBox."
Con.WriteLine "-ms <####>" & vbTab & "Set the MessageBox's Style. Numerical."
Con.WriteLine vbTab & vbTab & "OKOnly = 0"
Con.WriteLine vbTab & vbTab & "VBCritical = 16"
Con.WriteLine vbTab & vbTab & "VBQuestion = 32"
Con.WriteLine vbTab & vbTab & "VBExclamation = 48"
Con.WriteLine vbTab & vbTab & "VBInformation = 64"
Con.WriteLine vbTab & vbTab & "VBSystemModal = 4096 (Add 4096 to any of the above numbers as well)"
Con.WriteLine ""
Con.WriteLine "Paths are formatted, available strings:"
Con.WriteLine "Press any key to continue.."
Con.ReadChar
Con.WriteLine "(Keep in mind, these will be formatted by current user, likely SYSTEM)"
'and this should be changed to use the destination user, but there's no easy way
'that I can think to do this. These paths only really matter when copying a file
'which you'd normally want to be done under the SYSTEM account anyway
Con.WriteLine "%TEMP%" & vbTab & "Temp directory of the current user"
Con.WriteLine "%APPDATA%" & vbTab & "User's AppData Directory"
Con.WriteLine "%CD%" & vbTab & "Current Directory"
Con.WriteLine "%COMSPEC%" & vbTab & ""
Con.WriteLine "%CMDCMDLINE%" & vbTab & ""
Con.WriteLine "%LOCALAPPDATA%" & vbTab & "Local App Data directory"
Con.WriteLine "%APPDATA%" & vbTab & "User's AppData directory"
Con.WriteLine "%ALLUSERSPROFILE%" & vbTab & "All User's profile"
Con.WriteLine "%HOMEDRIVE%" & vbTab & "User's home drive (C: in most cases)"
Con.WriteLine "%HOMEPATH%" & vbTab & "User's home path"
Con.WriteLine "%PATH%" & vbTab & "PATH system variable"
Con.WriteLine "%PROGRAMDATA%" & vbTab & "ProgramData Directory"
Con.WriteLine "%PROGRAMFILES%" & vbTab & "Program Files Directory"
Con.WriteLine "%PROGRAMFILES(x86)%" & vbTab & "64-bit Program Files Directory"
Con.WriteLine "%RANDOM%" & vbTab & "A random number, ex: 651"
Con.WriteLine "%SYSTEMDRIVE%" & vbTab & "Local system drive, ex: C:"
Con.WriteLine "%SYSTEMROOT%" & vbTab & "Windows directory root directory"
Con.WriteLine "%USERDOMAIN%" & vbTab & "User's DOMAIN"
Con.WriteLine "%USERNAME%" & vbTab & "User's name/account"

Con.WriteLine "-d" & vbTab & vbTab & "Enable Debug output to file"
Con.WriteLine "-v" & vbTab & vbTab & "Enable Verbose output to console"
Con.WriteLine ""
Con.WriteLine "Examples:"
Con.WriteLine "Run ""calc.exe"" whenever the active session changes, until PC is restarted:"
Con.WriteLine "RunAsUser.exe ""calc.exe"" -p"
Con.WriteLine "Run ""calc.exe"" in only ""Bradley""'s session, if it exists:"
Con.WriteLine "RunAsUser.exe ""calc.exe"" -u ""bradley"""
Con.WriteLine "Run ""calc.exe"" on the first session found, waiting for someone to log in:"
Con.WriteLine "RunAsUser.exe ""calc.exe"" -w"

End Sub

