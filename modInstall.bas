Attribute VB_Name = "modInstall"
Option Explicit
' To DO:
' 1 - Groups
' 2 - Uninstall

Sub Main()
    
    On Error GoTo ErrorHandler
           
    '/ Start Setup process
    Load frmSetup1
    
    Exit Sub

ErrorHandler:
    MsgBox Err.Description, vbCritical, "Authenticator Installation 1.0"
End Sub

Function AddGroup(ByVal szProgramGroup As String)
    On Error GoTo ErrorHandler
    Dim bRetOnErr As Boolean
    
    Load frmGroup
    
    Call fCreateOSProgramGroup(frmGroup, szProgramGroup, bRetOnErr, True)
    
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Authenticator Installation 1.0"

End Function
Function AddUninstallFile()
    On Error GoTo ErrorHandler
    
    

    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Authenticator Installation 1.0"

End Function

Function AddRegistryKeysforFiles() As Boolean
    
    Dim strKey As String
    Dim strValueName As String
    Dim strValueData As String
    Dim strLine As String
    Const strSAFE_FOR_SCRIPT_FILES$ = "SafeForScript"
    Const lPART_IN_STRING As Long = 3
    Dim lIndex As Long
    Dim fErr As Boolean
    Dim intOffset  As Integer
    Dim intPlace As Integer
    Const CompareBinary = 0
    
    On Error GoTo ErrorHandler
    
    gstrWinDir = GetWindowsDir()

    '/ Find Setup.lst file
    gstrSetupInfoFile = gstrWinDir & gstrFILE_SETUP
    For lIndex = 1 To 1000 Step 1
        strLine = ReadIniFile(gstrSetupInfoFile, strSAFE_FOR_SCRIPT_FILES$, "File" & Format$(lIndex))
        
        If strLine = gstrNULL Then
            Exit For
        End If
        intOffset = 1
        intOffset = InStr(intOffset, strLine, ",") + 1
        strKey = strExtractFilenameItem(strLine, intOffset, fErr)
        intPlace = InStr(intOffset, strLine, ",")
        
        If intPlace <> 0 Then
            intPlace = intPlace + 1
            strValueName = strExtractFilenameItem(strLine, intPlace, fErr)
            intPlace = InStr(intPlace, strLine, ",")
            If intPlace <> 0 Then
                intPlace = intPlace + 1
                strValueData = strExtractFilenameItem(strLine, intPlace, fErr)
            End If
        End If
        
        If strValueName <> "" And strValueData <> "" Then
            Call AddSecurityKeys(strKey, strValueName, strValueData)
        ElseIf strValueName <> "" And strValueData = "" Then
            Call AddSecurityKeys(strKey, strValueName)
        Else
            Call AddSecurityKeys(strKey)
        End If
    Next lIndex
   
    AddRegistryKeysforFiles = True
       
    Exit Function
ErrorHandler:
    MsgBox Err.Description, vbCritical, "Authenticator Installation 1.0"
    
End Function
Function AddSecurityKeys(ByVal szKey As String, Optional szValueName As Variant, _
    Optional szValueData As Variant) As Boolean
    
    Dim hKey As Long
    On Error GoTo ErrorHandler
    AddSecurityKeys = False
    
    If Not RegCreateKey(HKEY_CLASSES_ROOT, _
                    ByVal szKey, _
                    ByVal "", _
                    hKey) Then
       AddSecurityKeys = True
       Exit Function
    End If
    
    If Not IsMissing(szValueName) Then
        If RegSetStringValue(hKey, szValueName, szValueData, False) Then
        AddSecurityKeys = True
        End If
    End If
    
    RegCloseKey hKey

Exit Function

ErrorHandler:
    MsgBox Err.Description, vbCritical, "Authenticator Installation 1.0"
    
End Function


