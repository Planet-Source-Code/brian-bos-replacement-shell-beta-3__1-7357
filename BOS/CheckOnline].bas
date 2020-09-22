Attribute VB_Name = "CheckOnline"
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public ReturnCode As Long

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal _
hKey As Long) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias _
"RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As _
String, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, _
lpData As Any, lpcbData As Long) As Long

Public Function ActiveConnection() As Boolean
Dim hKey As Long
Dim lpSubKey As String
Dim phkResult As Long
Dim lpValueName As String
Dim lpReserved As Long
Dim lpType As Long
Dim lpData As Long
Dim lpcbData As Long
ActiveConnection = False
lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, _
phkResult)

If ReturnCode = ERROR_SUCCESS Then
hKey = phkResult
lpValueName = "Remote Connection"
lpReserved = APINULL
lpType = APINULL
lpData = APINULL
lpcbData = APINULL
ReturnCode = RegQueryValueEx(hKey, lpValueName, _
lpReserved, lpType, ByVal lpData, lpcbData)
lpcbData = Len(lpData)
ReturnCode = RegQueryValueEx(hKey, lpValueName, _
lpReserved, lpType, lpData, lpcbData)

If ReturnCode = ERROR_SUCCESS Then
If lpData = 0 Then
ActiveConnection = False
Else
ActiveConnection = True
End If
End If

RegCloseKey (hKey)
End If

End Function

Public Sub Connect(strConnectName As String, blnPressConnect As Boolean)
Shell "rundll32.exe rnaui.dll,RnaDial " & strConnectName, vbNormalFocus
If blnPressConnect Then
DoEvents
SendKeys "{ENTER}", True
End If
End Sub

Public Sub Disconnect(strConnectName As String, blnPressDisconnect As Boolean)
Shell "rundll32.exe rnaui.dll,RnaDial " & strConnectName, 1
If blnPressDisconnect Then
    DoEvents
    SendKeys "{TAB} ", True
End If
End Sub

