Attribute VB_Name = "Registry"
Option Explicit
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ = 1
Dim Run As Long

Public Function AutoLoad()
'This function makes the program autoload on windows startup
'first opens the registry key for autostartup list in the windows registry
RegOpenKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", Run
'it then writes semipad to the registry
RegSetValueEx Run, "Semipad Autoload", 0, REG_SZ, ByVal App.Path & "\" & App.EXEName & ".exe", Len(App.Path & "\" & App.EXEName & ".exe")
'it then closes the key
RegCloseKey Run
End Function

Public Function AutoUnload()
'This function stops the program from automatically starting up
'first opens the registry key for autostartup list in the windows registry
RegOpenKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run\", Run
'It then removes the software
RegSetValueEx Run, "Semipad Autoload", 0, REG_SZ, ByVal 0, 0
'then it closes the key
RegCloseKey Run
End Function
