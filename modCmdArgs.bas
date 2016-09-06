Attribute VB_Name = "modCmdArgs"
Option Explicit

'=======================================
'     ============================
' GetCommandArgs - © Nik Keso 2009
'----------------------------------
'The function returns an array with the
'     command line arguments,
'contained in the command$, like cmd.exe


Public Function GetCommandArgs() As String()
Dim CountQ As Integer 'chr(34) counter
Dim OpenQ As Boolean ' left open string indicator (ex:"c:\bbb ccc.bat ) OpenQ=true, (ex:"c:\bbb ccc.bat" ) OpenQ=false
Dim ArgIndex As Integer
Dim tmpSTR As String
Dim strIndx As Integer
Dim TmpArr() As String
Dim comSTR As String
AddLog "Reading and verifying command line args..", True
    GetCommandArgs = Split("", " ")
    TmpArr = Split("", " ")
    comSTR = Trim$(Command$) 'remove front and back spaces
    If Len(comSTR) = 0 Then Exit Function
    CountQ = UBound(Split(comSTR, """"))
    If CountQ Mod 2 = 1 Then Exit Function 'like cmd.exe , command$ must contain even number of chr(34)=(")
    strIndx = 1

    Do
        If Mid$(comSTR, strIndx, 1) = """" Then OpenQ = Not OpenQ
        If Mid$(comSTR, strIndx, 1) = " " And OpenQ = False Then
            If tmpSTR <> "" Then 'don't include the spaces between args as args!!!!!
                ReDim Preserve TmpArr(ArgIndex)
                TmpArr(ArgIndex) = tmpSTR
                ArgIndex = ArgIndex + 1
            End If
            tmpSTR = ""
        Else
            tmpSTR = tmpSTR & Mid$(comSTR, strIndx, 1)
        End If
        strIndx = strIndx + 1
    Loop Until strIndx = Len(comSTR) + 1
    
    ReDim Preserve TmpArr(ArgIndex)
    TmpArr(ArgIndex) = tmpSTR
    GetCommandArgs = TmpArr
End Function
        

Public Sub ParseCommandArgs()
Dim i As Long
Dim strArray() As String

strArray = GetCommandArgs()
'AddLog "Parsing command line args..", True

    For i = 0 To UBound(strArray)
        If Left(strArray(i), 1) = "-" Then
        
            Select Case strArray(i)

                Case Is = "-b"
                    SafeToProceed = True

                Case Is = "-d"
                    blDebug = True
                    If (i + 1) <= UBound(strArray) Then
                        LogFileDirectory = TrimQuotes(strArray(i + 1))
                        LogFileDirectory = FormatPath(LogFileDirectory)
                        If LenB(LogFileDirectory) < 3 Then LogFileDirectory = App.Path & "\"
                    Else
                        Con.WriteLine "No log directory specified with -d!"
                        LogFileDirectory = App.Path & "\"
                    End If
                    
                Case Is = "-v"
                    blVerbose = True
                    
                Case Is = "-?"
                    DoShowHelp
                    End
                    
                Case Is = "-c"
                    CopyTo = TrimQuotes(strArray(i + 1))
                    CopyTo = FormatPath(CopyTo)
                    i = i + 1
                    
                Case Is = "-f"
                    FileToRun = TrimQuotes(strArray(i + 1))
                    SafeToProceed = True
                    i = i + 1
                
                Case Is = "-a"
                    FileArguments = TrimQuotes(strArray(i + 1), True)
                    i = i + 1
                    
                Case Is = "-a2"
                    'this switch must be at the end of the Command$ string
                    'Basically, everything after it is take (except the quotation marks, if any)
                    'I know.. super lazy to do it this way.. but it works so well. :/
                    FileArguments = Split(Command$, " -a2 ")(UBound(Split(Command$, " -a2 ")))
                    FileArguments = Trim$(FileArguments)
                    
                Case Is = "-h"
                    HideFile = True
                
                Case Is = "-m"
                    MessageBoxText = TrimQuotes(strArray(i + 1), True)
                    MessageBoxText = FormatPath(MessageBoxText, True)
                    ShowMessageBox = True
                    SafeToProceed = True
                    i = i + 1
                    
                Case Is = "-mt"
                    MessageBoxTitle = TrimQuotes(strArray(i + 1), True)
                    MessageBoxTitle = FormatPath(MessageBoxTitle, True)
                    SafeToProceed = True
                    i = i + 1
                
                Case Is = "-ms"
                    MessageBoxStyle = CLng(strArray(i + 1))
                    i = i + 1
                
                Case Is = "-l"
                    Launch = True
                    SafeToProceed = True
                
                Case Is = "-w"
                    WaitForUser = True
                    
                Case Is = "-i"
                    DownloadURL = TrimQuotes(strArray(i + 1), True)
                    i = i + 1
                
                Case Is = "-p"
                    PermanentlyRun = True
                    
                Case Is = "-all"
                    RunForAllSessions = True
                    
                Case Is = "-u"
                    RunForSpecificUser = TrimQuotes(strArray(i + 1), True)
                    
                Case Is = "-e"
                    ExecuteAsCommand = True
                    
                Case Is = "-t"
                    TerminateRunAsInstances = True
                
                Case Is = "-k"
                    KillBeforeCopy = True
                
            End Select
        
        End If
    Next
    
    If Not Left(strArray(0), 1) = "-" Then
        FileToRun = TrimQuotes(strArray(0), True)
        AddLog "FileToRun: " & FileToRun, True
        SafeToProceed = True
    End If
    
    If LenB(RunForSpecificUser) > 0 And RunForAllSessions Then RunForSpecificUser = ""
    
    
AddLog "Successfully parsed command line args. Args: " & UBound(strArray) + 1, True
End Sub

Public Function AddToArgs(strSwitch As String, strCommand As String) As String
Dim lngPosition

lngPosition = InStr(1, strCommand, " -a ")
If lngPosition = 0 Then lngPosition = InStr(1, strCommand, " -a2 ")

If lngPosition > 0 Then
    AddToArgs = Mid(strCommand, 1, lngPosition) & strSwitch & Mid(strCommand, lngPosition, Len(strCommand) - lngPosition + 1)
    Exit Function
End If

AddToArgs = strCommand & " " & strSwitch
End Function
