Option Explicit
Dim strLINK, strEXE, WSHShell
' Be sure to set up strLINK to match your VB6 installation.
strLINK = """C:\Program Files (x86)\Microsoft Visual Studio\VB98\LINK.EXE"""
strEXE = """" & WScript.Arguments(0) & """"
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run strLINK & " /EDIT /SUBSYSTEM:CONSOLE " & strEXE
Set WSHShell = Nothing
WScript.Echo "Complete!"