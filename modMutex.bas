Attribute VB_Name = "modMutex"
Option Explicit

Private Declare Function CreateMutex Lib "Kernel32.dll" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "Kernel32.dll" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "Kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Public Function OpenMutex(Name As String, Optional ByRef StatusOnly As Boolean = False) As Long
  Dim handle As Long
  Dim lngError
  ' create mutex object
  Err.Clear
    handle = CreateMutex(0&, 1, Name)
    lngError = Err.LastDllError
    'AddLog "lngError: " & lngError, True
    Err.Clear
    If (lngError = 183&) Then
        OpenMutex = 0
        'AddLog "name: " & Name & " - handle: " & handle & " openmutex: " & 0, True
    Else
        OpenMutex = handle
        'AddLog "name: " & Name & " - handle: " & handle & " openmutex: " & handle, True
    End If
    If StatusOnly Then CloseMutex handle
  
End Function
Public Sub CloseMutex(ByVal handle As Long)
'AddLog "CloseMutex: " & CStr(handle), True
  If handle = 0 Then Exit Sub
  ' release all
  Dim t As Integer
  For t = 1 To 100
    If ReleaseMutex(handle) = 0 Then Exit For
  Next t
  CloseHandle handle
  handle = 0
End Sub

Public Sub CloseAllMutex()
'this only closes the mutex **if the calling process owns it!**
CloseMutex MutexSystemHandle
CloseMutex MutexAdminHandle
CloseMutex MutexUserHandle
End Sub

Public Function IsSysMutexHeld() As Boolean

    If OpenMutex(MutexBaseName + "-SysMutex", True) = 0 Then IsSysMutexHeld = True

End Function

Public Function IsAdminMutexHeld() As Boolean

    If OpenMutex(MutexBaseName + "-AdminMutex", True) = 0 Then IsAdminMutexHeld = True

End Function

Public Function IsUserMutexHeld() As Boolean

    If OpenMutex(MutexBaseName + "-" & getCurrentUser(False), True) = 0 Then IsUserMutexHeld = True

End Function

Public Function GetUserMutex() As Long
    AddLog "attempting to reserve User Mutex..", True
    MutexUserHandle = OpenMutex(MutexBaseName + "-" & getCurrentUser(False), False)
    GetUserMutex = MutexUserHandle
    AddLog "GetUserMutex: " & MutexUserHandle, True
End Function

Public Function GetAdminMutex() As Long
    AddLog "attempting to reserve Admin Mutex..", True
    MutexAdminHandle = OpenMutex(MutexBaseName + "-AdminMutex", False)
    GetAdminMutex = MutexAdminHandle
    AddLog "GetAdminMutex: " & MutexAdminHandle, True
End Function

Public Function GetSysMutex() As Long
    AddLog "attempting to reserve Sys Mutex..", True
    MutexSystemHandle = OpenMutex(MutexBaseName + "-SysMutex", False)
    GetSysMutex = MutexSystemHandle
    AddLog "GetSysMutex: " & MutexSystemHandle, True
End Function

