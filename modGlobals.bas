Attribute VB_Name = "modGlobals"
Option Explicit

Public SafeToProceed As Boolean

'CMD Args
Public blVerbose As Boolean         '-v
Public blDebug As Boolean          'old -d
Public LogFileDirectory As String   '-d
Public CopyTo As String             '-c
Public FileToRun As String          '-f
Public DownloadURL As String        '-i
Public FileArguments As String      '-a
Public HideFile As Boolean          '-h
Public Launch As Boolean            '-l
Public WaitForUser As Boolean       '-w
Public PermanentlyRun As Boolean    '-p
Public RunForAllSessions As Boolean '-all
Public RunForSpecificUser As String '-u
Public ExecuteAsCommand As Boolean  '-e
Public TerminateRunAsInstances As Boolean  '-t
Public KillBeforeCopy As Boolean    '-k


'msgbox
Public MessageBoxText As String     '-m
Public MessageBoxTitle As String    '-mt
Public MessageBoxStyle As Long      '-ms
Public ShowMessageBox As Boolean

'Mutex
Public MutexBaseName As String
Public MutexAdminHandle As Long
Public MutexSystemHandle As Long
Public MutexUserHandle As Long
