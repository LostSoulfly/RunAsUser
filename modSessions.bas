Attribute VB_Name = "modSessions"

Option Explicit

Private Type SECURITY_ATTRIBUTES
    nLength                 As Long
    lpSecurityDescriptor    As Long
    bInheritHandle          As Long
End Type

Private Type STARTUPINFO
    cb                      As Long
    lpReserved              As Long
    lpDesktop               As Long
    lpTitle                 As Long
    dwX                     As Long
    dwY                     As Long
    dwXSize                 As Long
    dwYSize                 As Long
    dwXCountChars           As Long
    dwYCountChars           As Long
    dwFillAttribute         As Long
    dwFlags                 As Long
    wShowWindow             As Integer
    cbReserved2             As Integer
    lpReserved2             As Long
    hStdInput               As Long
    hStdOutput              As Long
    hStdError               As Long
End Type
      
Private Type PROCESS_INFORMATION
    hProcess                As Long
    hThread                 As Long
    dwProcessID             As Long
    dwThreadId              As Long
End Type

Private Const STARTF_USESTDHANDLES              As Long = &H100&
Private Const STARTF_USESHOWWINDOW              As Long = &H1

Public Const SW_MINIMIZE = 0
Public Const SW_SHOWNORMAL = 1

Private Const CREATE_DEFAULT_ERROR_MODE         As Long = &H4000000
Private Const CREATE_NEW_CONSOLE                As Long = &H10&
Private Const CREATE_NEW_PROCESS_GROUP          As Long = &H200&
Private Const CREATE_SEPARATE_WOW_VDM           As Long = &H800&
Private Const CREATE_SUSPENDED                  As Long = &H4&
Private Const CREATE_UNICODE_ENVIRONMENT        As Long = &H400&

Private Declare Function WTSGetActiveConsoleSessionId Lib "Kernel32.dll" () As Long
Private Declare Function WTSQueryUserToken Lib "wtsapi32.dll" (ByVal SessionID As Long, ByRef phToken As Long) As Long
Private Declare Function CreateEnvironmentBlock Lib "userenv.dll" (ByRef lpEnvironment As Long, ByVal hToken As Long, ByVal bInherit As Long) As Long
'Private Declare Function CreateProcessAsUser Lib "advapi32.dll" Alias "CreateProcessAsUserW" (ByVal hToken As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CreateProcessAsUser Lib "advapi32.dll" _
        Alias "CreateProcessAsUserA" _
        (ByVal hToken As Long, _
        ByVal lpApplicationName As Long, _
        ByVal lpCommandLine As String, _
        ByVal lpProcessAttributes As Long, _
        ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, _
        ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, _
        ByVal lpCurrentDirectory As String, _
        lpStartupInfo As STARTUPINFO, _
        lpProcessInformation As PROCESS_INFORMATION) As Long
        
Private Declare Function DestroyEnvironmentBlock Lib "userenv.dll" (ByVal lpEnvironment As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ProcessIdToSessionId Lib "Kernel32.dll" (ByVal dwProcessID As Long, ByRef pSessionId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'below is for current user sid
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function ConvertSidToStringSid Lib "advapi32.dll" Alias "ConvertSidToStringSidA" (ByVal lpSid As Long, lpString As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Public Const PROCESS_VM_READ As Long = (&H10)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
'Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
Public Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long


Private Const WTS_CURRENT_SERVER_HANDLE = 0&

Private Enum WTS_CONNECTSTATE_CLASS
    WTSActive       '0
    WTSConnected    '1
    WTSConnectQuery '2
    WTSShadow       '3
    WTSDisconnected '4
    WTSIdle         '5
    WTSListen       '6
    WTSReset        '7
    WTSDown         '8
    WTSInit         '9
End Enum

Private Type WTS_SESSION_INFO
    SessionID As Long
    pWinStationName As Long
    State As WTS_CONNECTSTATE_CLASS
End Type

Private Declare Function WTSEnumerateSessions _
    Lib "wtsapi32.dll" Alias "WTSEnumerateSessionsA" ( _
    ByVal hServer As Long, ByVal Reserved As Long, _
    ByVal Version As Long, ByRef ppSessionInfo As Long, _
    ByRef pCount As Long _
    ) As Long
    
Enum SECURITY_IMPERSONATION_LEVEL
    SecurityAnonymous
    SecurityIdentification
    SecurityImpersonation
    SecurityDelegation
End Enum

Enum TOKEN_TYPE
    TokenPrimary = 1
    TokenImpersonation
End Enum

Public Const GENERIC_WRITE             As Long = &H40000000
Public Const GENERIC_EXECUTE           As Long = &H20000000
Public Const GENERIC_ALL               As Long = &H10000000
Public Const MAXIMUM_ALLOWED As Long = &H2000000

Declare Function DuplicateTokenEx _
        Lib "advapi32" ( _
                   ByVal hExistingToken As Long, _
                   ByVal dwDesiredAccess As Long, _
                         lpTokenAttributes As Any, _
                   ByVal ImpersonationLevel As SECURITY_IMPERSONATION_LEVEL, _
                   ByVal TokenType As TOKEN_TYPE, _
                         phNewToken As Long _
                       ) As Long
    
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" ( _
    ByVal pMemory As Long)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function lstrlenA Lib "kernel32" ( _
    ByVal lpString As String) As Long

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Declare Function WTSQuerySessionInformation Lib "wtsapi32.dll" _
Alias "WTSQuerySessionInformationW" (ByVal hServer As Long, ByVal SessionID _
As Long, ByVal wtsInfoClass As Long, ByRef pBuffer As Long, ByRef dwSize As _
Long) As Long

Public Function GetSessionName(Optional Session As Long = -1) As String
    Dim pBuffer As Long
    Dim dwSize As Long
    If WTSQuerySessionInformation(0, Session, 5, pBuffer, dwSize) Then
        Dim clientName As String
        clientName = String(dwSize, 0)
        CopyMemory ByVal StrPtr(clientName), ByVal pBuffer, dwSize
        WTSFreeMemory pBuffer
        GetSessionName = Left(clientName, InStr(clientName, Chr(0)) - 1)
    End If
End Function

Private Function GetWTSSessions() As WTS_SESSION_INFO()
    Dim RetVal As Long
    Dim lpBuffer As Long
    Dim Count As Long
    Dim p As Long
    Dim arrSessionInfo() As WTS_SESSION_INFO

    RetVal = WTSEnumerateSessions(WTS_CURRENT_SERVER_HANDLE, _
                                   0&, _
                                   1, _
                                   lpBuffer, _
                                   Count)
    If RetVal Then
        ' WTSEnumerateProcesses was successful.

        p = lpBuffer
        ReDim arrSessionInfo(Count - 1)
        CopyMemory arrSessionInfo(0), ByVal p, _
           Count * LenB(arrSessionInfo(0))
        ' Free the memory buffer.
        WTSFreeMemory lpBuffer

     Else
    'probably winXP?
    WTSFreeMemory lpBuffer
    End If
    GetWTSSessions = arrSessionInfo
End Function

Public Function getSessionArray() As String()
    Dim i As Integer
    Dim arrWTSSessions() As WTS_SESSION_INFO
    arrWTSSessions = GetWTSSessions
    Dim myArr() As String
    ReDim myArr(0)
    For i = LBound(arrWTSSessions) To UBound(arrWTSSessions)
        If arrWTSSessions(i).State <> 6 Then

        If Not myArr(UBound(myArr)) = "" Then ReDim Preserve myArr(UBound(myArr) + 1)
            myArr(UBound(myArr)) = arrWTSSessions(i).SessionID

            'AddLog "Session ID: " & arrWTSSessions(i).SessionID, True
            'AddLog "Machine Name: " & _
              PointerToStringA(arrWTSSessions(i).pWinStationName), True
            'AddLog "Session Name: " & GetSessionName(arrWTSSessions(i).SessionID), True
            'AddLog "Connect State: " & arrWTSSessions(i).State, True
            'AddLog "***********", True
        End If
    Next i
    getSessionArray = myArr
End Function

Public Function getActiveSession() As Long
    Dim i As Integer
    Dim arrWTSSessions() As WTS_SESSION_INFO
    getActiveSession = -1
    arrWTSSessions = GetWTSSessions
    Dim myArr() As String
    ReDim myArr(0)
    For i = LBound(arrWTSSessions) To UBound(arrWTSSessions)
        If arrWTSSessions(i).State = 0 Then getActiveSession = arrWTSSessions(i).SessionID

        If Not myArr(UBound(myArr)) = "" Then ReDim Preserve myArr(UBound(myArr) + 1)
            myArr(UBound(myArr)) = arrWTSSessions(i).SessionID

            'AddLog "Session ID: " & arrWTSSessions(i).SessionID, True
            'AddLog "Machine Name: " & _
              PointerToStringA(arrWTSSessions(i).pWinStationName), True
            'AddLog "Session Name: " & GetSessionName(arrWTSSessions(i).SessionID), True
            'AddLog "Connect State: " & arrWTSSessions(i).State, True
            'AddLog "***********", True
    Next i
    
    AddLog "ActiveSession: " & getActiveSession
    AddLog "getSession would be: " & getSession
    
End Function


Public Function getUserSession(UserName As String) As Long
Dim Sess As Variant

getUserSession = -1

UserName = Trim$(UCase(UserName))
For Each Sess In getSessionArray
    If Trim$(UCase(GetSessionName(CLng(Sess)))) = UserName Then getUserSession = CLng(Sess)
    AddLog "Session: " & Sess & " (" & Trim$(UCase(GetSessionName(CLng(Sess)))) & ") vs " & UserName
Next
    
End Function

Public Function PointerToStringA(ByVal lpStringA As Long) As String
   Dim nLen As Long
   Dim sTemp As String ''
   If lpStringA Then
      nLen = lstrlenA(ByVal lpStringA)
      If nLen Then
         sTemp = String(nLen, vbNullChar)
         lstrcpy sTemp, ByVal lpStringA
         PointerToStringA = sTemp
      End If
   End If
End Function

'Public Function getProcessSessionID(ProcessID As Long) As Long
'    Dim lngSessionID As Long
'    getProcessSessionID = ProcessIdToSessionId(ProcessID, lngSessionID)
'End Function

Public Function getSession() As Long
    getSession = WTSGetActiveConsoleSessionId()
End Function

'Public Function getProcessID() As Long
'    getProcessID = GetCurrentProcessId
'End Function

Public Function RunAsSessionId(ApplicationName As String, ApplicationCommand As String, lngSessionID As Long, Optional blHide As Boolean = False) As Long
On Error GoTo oops
    Dim typProcess      As PROCESS_INFORMATION
    Dim typStartup      As STARTUPINFO
    Dim typSecurity     As SECURITY_ATTRIBUTES
    Dim lngToken  As Long
    Dim strApplication  As String
    Dim strDirectory    As String
    Dim lngEnvBlock     As Long
 
    With typSecurity
        .nLength = Len(typSecurity)
        .bInheritHandle = 1&
        .lpSecurityDescriptor = 0&
    End With
    
    ZeroMemory typStartup, Len(typStartup)  'unsure if necessary, but doesn't seem to hurt, either.
    With typStartup
        .cb = Len(typStartup)
        '.dwFlags = STARTF_USESHOWWINDOW
        '.wShowWindow = 6
    End With

    WTSQueryUserToken lngSessionID, lngToken
    AddLog "Session Token: " & CStr(lngToken), True
        If lngToken = 0 Then
            AddLog "Skipping session: " & CStr(lngToken), True
            CloseHandle lngToken
            Exit Function
        End If
    
    CreateEnvironmentBlock lngEnvBlock, lngToken, 0

    strDirectory = PreviousDir(FileToRun)
    CreateProcessAsUser lngToken, 0&, ApplicationName & " " & ApplicationCommand, 0&, 0&, False, CREATE_DEFAULT_ERROR_MODE, 0&, strDirectory, typStartup, typProcess

    AddLog "LastErr: " & Err.LastDllError, True
    DestroyEnvironmentBlock lngEnvBlock
    AddLog "ProcID: " & typProcess.dwProcessID & " hProc: " & typProcess.hProcess, True
    AddLog "AppName: " & ApplicationName & " AppComm: " & ApplicationCommand, True
    If Not Err.LastDllError = 0 Then AddLog "DLLERR: " & CStr(Err.LastDllError) 'SendData BuildGeneric(CDebug, "DLLERR: " & CStr(Err.LastDllError))

    RunAsSessionId = typProcess.dwProcessID
    CloseHandle lngToken

Exit Function

oops:

AddLog "RunAsSessionId Err: " & Err.Number & " " & Err.Description
Resume Next

End Function

Public Function GetCurrentUserSid() As String
    Dim hProcessID      As Long
    Dim hToken          As Long
    Dim lNeeded         As Long
    Dim baBuffer()      As Byte
    Dim sBuffer         As String
    Dim lpSid           As Long
    Dim lpString        As Long

    hProcessID = GetCurrentProcess()
    If hProcessID <> 0 Then
        If OpenProcessToken(hProcessID, &H20008, hToken) = 1 Then 'TOKEN_READ = &H20008
            Call GetTokenInformation(hToken, 1, ByVal 0, 0, lNeeded)
            ReDim baBuffer(0 To lNeeded)
            '--- enum TokenInformationClass { TokenUser = 1, TokenGroups = 2, ... }
            If GetTokenInformation(hToken, 1, baBuffer(0), UBound(baBuffer), lNeeded) = 1 Then
                Call CopyMemory(lpSid, baBuffer(0), 4)
                If ConvertSidToStringSid(lpSid, lpString) Then
                    sBuffer = Space(lstrlen(lpString))
                    Call CopyMemory(ByVal sBuffer, ByVal lpString, Len(sBuffer))
                    Call LocalFree(lpString)
                    GetCurrentUserSid = sBuffer
                End If
            End If
            Call CloseHandle(hToken)
        End If
        Call CloseHandle(hProcessID)
    End If
End Function


