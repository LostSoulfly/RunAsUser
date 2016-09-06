Attribute VB_Name = "modUtils"
Option Explicit
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Sub Sleep Lib "Kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function IsUserAnAdmin Lib "Shell32" Alias "#680" () As Integer

Public Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias _
    "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000

Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Sub AddLog(ByVal Text As String, Optional Verbose As Boolean = False)
    Dim FileName As String
    Dim f As Long
    Dim fColor As Long
    Text = "[" & Replace(Date, "-", "/") & Chr(32) & Time & "] - " & Text

    If Verbose And blVerbose Then
    
        fColor = Con.ForeColor
        Con.ForeColor = conGreen
        Con.WriteLine Text
        Con.ForeColor = fColor
        
    ElseIf Verbose = False Then
        Con.WriteLine Text
    End If

If blDebug Then
    If Not DirExists(LogFileDirectory) Then CreateFolder LogFileDirectory
    If Not IsAdmin Then
        FileName = LogFileDirectory & App.ThreadID & "-" & Environ("USERNAME") & ".txt"
    Else
        FileName = LogFileDirectory & App.ThreadID & "-" & Environ("USERNAME") & "-IsAdmin" & ".txt"
    End If

    If Not FileExists(FileName, True) Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If

    f = FreeFile
    Open FileName For Append As #f
    Print #f, Text
    Close #f
End If
    
End Sub

Public Function IsAdmin() As Boolean
    If IsUserAnAdmin = 1 Then IsAdmin = True
        'AddLog "IsAdmin: " & IsUserAnAdmin, True
End Function


Public Function getCurrentUser(Optional isSystem As Boolean = True) As String

     ' Dimension variables
     Dim lpBuff As String * 25
     Dim ret As Long, sUserName As String

     ' Get the user name minus any trailing spaces found in the name.
     ret = GetUserName(lpBuff, 25)
     sUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    
    'AddLog "User (before checking): " & sUserName
    'AddLog "sid: " & GetCurrentUserSid
    
    'side S-1-5-18 = SYSTEM
    If GetCurrentUserSid = "S-1-5-18" Then sUserName = "SYSTEM" 'this should help resolve localized OSes
    If sUserName = "" Then sUserName = Environ("USERNAME")
    
    If isSystem = False Then
        If UCase(sUserName) = "SYSTEM" Then
            sUserName = GetSessionName(getSession) & "*" 'the asterisk implies we're running as SYSTEM but are reporting the current user as ourself
            If sUserName = "*" Then sUserName = "SYSTEM"
        End If
    End If
    
    If Len(sUserName) <= 1 Then sUserName = "Unknown"
    
    getCurrentUser = sUserName
    
End Function

Public Function isSystem() As Boolean
    If getCurrentUser = "SYSTEM" Then isSystem = True
    AddLog "IsSystem: " & isSystem, True
End Function

Public Function DirExists(ByRef Path As String) As Boolean
    On Error GoTo errorhandler
    ' test the directory attribute
    
    If Not Right(Path, 1) = "\" Then Path = Path & "\"
    
    DirExists = GetAttr(Path) And vbDirectory
errorhandler:
    ' if an error occurs, this function returns False
End Function

Public Function CreateFolder(destDir As String) As Boolean
   Dim i As Long
   Dim prevDir As String
    
   On Error Resume Next
    
   For i = Len(destDir) To 1 Step -1
       If Mid(destDir, i, 1) = "\" Then
           prevDir = Left(destDir, i - 1)
           Exit For
       End If
   Next i
    
   If prevDir = "" Then CreateFolder = False: Exit Function
   If Not Len(Dir(prevDir & "\", vbDirectory)) > 0 Then
       If Not CreateFolder(prevDir) Then CreateFolder = False: Exit Function
   End If
    
   On Error GoTo errDirMake
   MkDir destDir
   CreateFolder = True
   Exit Function
    
errDirMake:
   CreateFolder = False

End Function


Public Function PreviousDir(Path As String) As String
Dim i As Long

    For i = (Len(Path)) To 1 Step -1
       If Mid(Path, i, 1) = "\" Then
           PreviousDir = Left(Path, i)
           Exit For
       End If
   Next i
   
End Function

Public Function FileFromDir(ByVal Path As String) As String
Dim i As Long

    For i = (Len(Path) - 1) To 1 Step -1
       If Mid(Path, i, 1) = "\" Then
           FileFromDir = Right(Path, Len(Path) - i)
           Exit Function
       End If
   Next i
   
   FileFromDir = Path
   
End Function

Public Function WriteFile(Data As String, FileName As String) As Boolean
On Error GoTo oops:
Dim InFile As Integer
 ' Get none used file handle number
 InFile = FreeFile
 ' Clear the file and recreate it
 Open FileName For Output As InFile
 Close InFile
 ' Open the file whit Binary, the best way!
 Open FileName For Binary Access Write As InFile
 ' Save data into the open file
 Put InFile, , Data
 Close InFile
 
Exit Function
 
oops:
WriteFile = False
 
End Function


Public Function FileExists(ByRef sFileName As String, Optional Derp As Boolean = False) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(sFileName) And vbDirectory) <> vbDirectory
    If Derp = False Then AddLog "FileExists?: " & sFileName & ": " & FileExists, True
End Function

Private Function ReplaceFast(ByRef Text As String, ByRef Value As String)
    Text = Replace(Text, "%" & Value & "%", Environ(Value))
End Function

Public Function FormatPath(ByVal Path As String, Optional SkipSlash As Boolean = False) As String
    If Left(Path, 1) = "." Then Path = App.Path & "\" & Right(Path, Len(Path) - 1)
    
    If InStr(1, Path, "%") > 0 Then
        ReplaceFast Path, "APPDATA"
        ReplaceFast Path, "CD"
        ReplaceFast Path, "COMSPEC"
        ReplaceFast Path, "CMDCMDLINE"
        ReplaceFast Path, "LOCALAPPDATA"
        ReplaceFast Path, "APPDATA"
        ReplaceFast Path, "ALLUSERSPROFILE"
        ReplaceFast Path, "HOMEDRIVE"
        ReplaceFast Path, "HOMEPATH"
        ReplaceFast Path, "PATH"
        ReplaceFast Path, "PROGRAMDATA"
        ReplaceFast Path, "PROGRAMFILES"
        ReplaceFast Path, "PROGRAMFILES(x86)"
        ReplaceFast Path, "RANDOM"
        ReplaceFast Path, "SYSTEMDRIVE"
        ReplaceFast Path, "SYSTEMROOT"
        ReplaceFast Path, "TEMP"
        ReplaceFast Path, "USERDOMAIN"
        ReplaceFast Path, "USERNAME"
        Path = Replace(Path, "%APPPATH%", App.Path)
    End If
    
    If SkipSlash = False Then If Not Right(Path, 1) = "\" Then Path = Path & "\"
    FormatPath = Replace(Path, "/", "\")
    
    Do While InStr(1, FormatPath, "\\") > 0
        FormatPath = Replace(FormatPath, "\\", "\")
    Loop
    
End Function

Public Function TrimQuotes(strToTrim As String, Optional TrimOnce As Boolean = False) As String

Do While Left$(strToTrim, 1) = Chr(34)
    strToTrim = Right$(strToTrim, Len(strToTrim) - 1): If TrimOnce Then Exit Do
Loop

Do While Right$(strToTrim, 1) = Chr(34)
    strToTrim = Left$(strToTrim, Len(strToTrim) - 1): If TrimOnce Then Exit Do
Loop

TrimQuotes = strToTrim

End Function

Public Function ExecFile(filePath As String, FileArgs As String, Optional Operation As String = "open", Optional Directory As String = vbNullString, Optional Visible As Long = 0) As Long
Dim ret As Long

    ret = ShellExecute(0, Operation, filePath, FileArgs, Directory, Visible)
    AddLog "execFile (shellex): (" & filePath & ") = " & CStr(ret), True
    DoEvents
    If Visible > 6 Then Visible = 0
    If ret <= 32 Then
        ret = Shell(filePath & FileArgs, Visible)
        AddLog "execFile (shell): (" & filePath & ") = " & CStr(ret), True
    End If
    
    ExecFile = ret
End Function


Public Function DownloadFile(sSourceUrl As String, _
                             sLocalFile As String) As Boolean

   DeleteUrlCacheEntry sSourceUrl

   DoEvents
   
   DownloadFile = URLDownloadToFile(0&, _
                                    sSourceUrl, _
                                    sLocalFile, _
                                    BINDF_GETNEWESTVERSION, _
                                    0&) = ERROR_SUCCESS

End Function

