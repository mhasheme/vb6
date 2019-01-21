Attribute VB_Name = "GBINI"

Option Explicit
Option Compare Text

Public Function DeleteRegistryValue(ByVal hKey As Long, ByVal KeyName As String, ByVal ValueName As String) As Boolean
    ' Delete a registry value
    ' Return True if successful, False if the value hasn't been found

    Dim handle As Long
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then Exit Function
    
    ' Delete the value (returns 0 if success)
    DeleteRegistryValue = (RegDeleteValue(handle, ValueName) = 0)
    
    ' Close the handle
    RegCloseKey handle
    
End Function

Public Function bGetRegistrySetting(ByVal lRoot As Long, ByVal sPath As String, ByVal sSetting As String, sRegistryEntry As String, Optional ByVal sHost As String) As Boolean
'====================================================================================================
'Created Date: July 23 1998
'By: Brent Horan
Dim bSuccess As Boolean
Dim iLoop As Integer
Dim hKey As Long, lType As Long, lReturnLen As Long, lReturn As Long
Dim sReturnValue As String
'====================================================================================================
On Error GoTo ProcErrorHandler:

sReturnValue = ""
bSuccess = False

'Open the root key...
If sHost <> "" Then
    'Attempt to connect to the remote host's hive...
    lReturn = RegConnectRegistry(sHost, lRoot, hKey)
    
    If lReturn <> ERROR_SUCCESS Then
         MsgBox "Unable to connect to the Registry on " & sHost & ".", vbExclamation + vbOKOnly
    End If
End If
    
'in case this is a host value or not we need two different approches here...
If sHost <> "" Then
    lReturn = RegOpenKeyEx(hKey, sPath, 0&, KEY_QUERY_VALUE, hKey)  'for remote we need the key already opened...
Else
    lReturn = RegOpenKeyEx(lRoot, sPath, 0&, KEY_QUERY_VALUE, hKey)    'for local we need the root...
End If

'check for a couple of specific errors, these are the ones I think we will bump into and we can add as we go...
Select Case lReturn
    Case ERROR_SUCCESS '0
        'init the string...
        'Ticket #24352 - PIPEDA
        If glbHRSSSecure Then
            sReturnValue = String$(344, "0")
        Else
            sReturnValue = String$(255, "0")
        End If
            
        lReturnLen = Len(sReturnValue) + 1
        
        If RegQueryValueEx(hKey, sSetting, 0&, lType, ByVal sReturnValue, lReturnLen) = ERROR_SUCCESS Then
            sReturnValue = Left$(sReturnValue, lReturnLen - 1) 'clear the null returned, the data length return includes the null...
            bSuccess = True
        End If
    Case ERROR_FILE_NOT_FOUND '2
        MsgBox "File Not Found for " & sPath & sSetting, vbExclamation + vbOKOnly
    Case ERROR_PATH_NOT_FOUND '3
        MsgBox "Path Not Found for " & sPath & sSetting, vbExclamation + vbOKOnly
    Case ERROR_ACCESS_DENIED ' 5
        MsgBox "Access to Key " & sPath & sSetting & " Denied.", vbExclamation + vbOKOnly
    Case Else
        MsgBox "Unexpected error reading the Registry Setting for " & sPath & "\" & sSetting, vbExclamation + vbOKOnly
End Select
        
ProcExit:

    'Disconnect from the remote registry if necessary...
    If hKey <> 0 Then lReturn = RegCloseKey(hKey)
    If sHost <> "" Then lReturn = WNetCancelConnection2(sHost & "\ipc$", 0, 0) 'will ignore if not connected...
        
    bGetRegistrySetting = bSuccess
    sRegistryEntry = sReturnValue
    
    Exit Function
    
ProcErrorHandler:
    
    MsgBox "Unexpected error reading the Registry Setting for " & sSetting, vbExclamation + vbOKOnly
    Exit Function
    
End Function


Public Function WriteRegistrySetting(ByVal lRoot As Long, ByVal sPath As String, ByVal sSetting As String, ByVal sValue As String, Optional ByVal sHost As String) As Boolean
'====================================================================================================
'Created Date: July 23 1998
'By: Brent Horan
'====================================================================================================
'Local Vars:
Dim bSuccess As Boolean
Dim hKey As Long, lType As Long, lLen As Long, lReturn As Long
Dim sReturnValue As String
Dim vValue As String
'====================================================================================================
    On Error GoTo ProcErrorHandler1:
    bSuccess = False
    'Open the root key...
    If sHost <> "" Then
        'Attempt to connect to the remote host's hive...
        lReturn = RegConnectRegistry(sHost, lRoot, hKey)
        If lReturn <> ERROR_SUCCESS Then
             MsgBox "Unable to connect to the Registry on " & sHost & ".", vbExclamation + vbOKOnly
        End If
    End If
    'in case this is a host value or not we need two different approches here...
    If sHost <> "" Then
        lReturn = RegOpenKeyEx(hKey, sPath, 0&, KEY_WRITE, hKey)  'for remote we need the key already opened...
    Else
        lReturn = RegOpenKeyEx(lRoot, sPath, 0&, KEY_WRITE, hKey)    'for local we need the root...
    End If
    'check for a couple of specific errors, these are the ones I think we will bump into and we can add as we go...
    Select Case lReturn
        Case ERROR_SUCCESS '0
            'init the string...
            vValue = sValue + Chr$(0)
            lLen = Len(vValue)
            If RegSetValueEx(hKey, sSetting, 0&, 1&, vValue, lLen) = ERROR_SUCCESS Then
                bSuccess = True
            End If
        Case ERROR_FILE_NOT_FOUND '2
            MsgBox "File Not Found for " & sPath & sSetting, vbExclamation + vbOKOnly
        Case ERROR_PATH_NOT_FOUND '3
            MsgBox "Path Not Found for " & sPath & sSetting, vbExclamation + vbOKOnly
        Case ERROR_ACCESS_DENIED ' 5
            MsgBox "Access to Key " & sPath & sSetting & " Denied.", vbExclamation + vbOKOnly
        Case Else
            MsgBox "Unexpected error reading the Registry Setting for " & sPath & "\" & sSetting, vbExclamation + vbOKOnly
    End Select
        
ProcExit:
    'Disconnect from the remote registry if necessary...
    If hKey <> 0 Then lReturn = RegCloseKey(hKey)
    If sHost <> "" Then lReturn = WNetCancelConnection2(sHost & "\ipc$", 0, 0) 'will ignore if not connected...
        
    WriteRegistrySetting = bSuccess
    
    Exit Function
    
ProcErrorHandler1:
    
    MsgBox "Unexpected error writing the Registry Setting for " & sSetting, vbExclamation + vbOKOnly
    Exit Function
    
End Function

Public Function DoesKeyExist(ByVal lRoot As Long, ByVal sPath As String) As Boolean
    Dim hKey As Long

    On Error GoTo ProcErrorHandler
    
    'check for a couple of specific errors, these are the ones I think we will bump into and we can add as we go...
    Select Case RegOpenKeyEx(lRoot, sPath, 0&, KEY_QUERY_VALUE, hKey)
        Case ERROR_SUCCESS '0
            DoesKeyExist = True
            RegCloseKey hKey
        Case ERROR_FILE_NOT_FOUND '2
            DoesKeyExist = False
        Case ERROR_PATH_NOT_FOUND '3
            DoesKeyExist = False
        Case ERROR_ACCESS_DENIED ' 5
            DoesKeyExist = False
        Case Else
            DoesKeyExist = False
    End Select
    
    Exit Function
    
ProcErrorHandler:
    
    MsgBox "Unexpected error reading the Registry Setting" + vbOKOnly
    Exit Function
    
End Function
                       
Public Function INIWrite(sSection As String, sKeyName As String, sNewString As String, sINIFileName As String) As Boolean
    Call WritePrivateProfileString(sSection, sKeyName, sNewString, sINIFileName)
    INIWrite = (Err.Number = 0)
End Function

Public Function INIRead(sSection As String, sKeyName As String, sINIFileName As String) As String
    Dim sRet As String

    sRet = String(255, Chr(0))
    INIRead = Left(sRet, GetPrivateProfileString(sSection, ByVal sKeyName, "", sRet, Len(sRet), sINIFileName))
End Function

