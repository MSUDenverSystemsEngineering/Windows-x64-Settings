'*=======================================================================================
'*
'* Disclaimer
'*
'* This script is not supported under any Microsoft standard support program or service. This 
'* script is provided AS IS without warranty of any kind. Microsoft further disclaims all 
'* implied warranties including, without limitation, any implied warranties of merchantability 
'* or of fitness for a particular purpose. The entire risk arising out of the use or performance 
'* of this script remains with you. In no event shall Microsoft, its authors, or anyone else 
'* involved in the creation, production, or delivery of this script be liable for any damages 
'* whatsoever (including, without limitation, damages for loss of business profits, business 
'* interruption, loss of business information, or other pecuniary loss) arising out of the use 
'* of or inability to use this script, even if Microsoft has been advised of the possibility 
'* of such damages.
'*
'*=======================================================================================
' //***************************************************************************
' // ***** Script Header *****
' //
' // Solution:  Solution Accelerator for Business Desktop Deployment
' // File:      ZTI-SpecialFolderLib.vbs
' //
' // Purpose:   Additional functions that can be used by all other scripts
' //
' // Usage:     <script language="VBScript" src="ZTI-SpecialFolderLib.vbs"/>
' //
' // Customer Build Version:      1.0.3
' // Customer Script Version:     1.0.3
' //
' // Customer History:
' // 1.0.0   MDM  07/26/2008  Created.
' // 1.0.1   MDM  07/26/2008  Added GetAllProfileFolders function.
' // 1.0.2   MDM  07/26/2008  Added GetPublicProfile function.  Updated
' //                          GetAllUsersProfile and GetDefaultUsersProfile to
' //                          return the full path on all OS versions.  Updated
' //                          GetAllProfileFolders to include Public.
' // 1.0.3   MDM  04/28/2009  Added missing GetSystemProfile function.
' //
' // ***** End Header *****
' //***************************************************************************

'On Error Resume Next

Function UserExit(sType, sWhen, sDetail, bSkip)

    oLogging.CreateEntry "USEREXIT:ZTI-SpecialFolderLib.vbs started: " & sType & " " & sWhen & " " & sDetail, LogTypeInfo

    UserExit = Success

End Function


Function GetProfileList()

    On Error Resume Next

    Dim arrProfileEntries()
    Const HKEY_LOCAL_MACHINE  = &H80000002

    strProfileList = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Set objRegistry = GetObject("winmgmts:root\default:StdRegProv")

    lKeyRC = objRegistry.EnumKey(HKEY_LOCAL_MACHINE, strProfileList, sKeys)

    If (lKeyRC = 0) And (Err.Number = 0) Then
        For intKey = LBound(sKeys) To UBound(skeys)
            ReDim Preserve arrProfileEntries(intKey)
            arrProfileEntries(intKey) = sKeys(intKey)
            'WScript.Echo intKey & " " & sKeys(intKey)
        Next
    Else
        oLogging.CreateEntry "GetProfileList Function: Error enumerating Registry keys in " & strProfileList, LogTypeError
    End If

    GetProfileList = arrProfileEntries

End Function


Function GetProfileImagePath(strProfileSid)

    On Error Resume Next

    GetProfileImagePath = null
    Const HKEY_LOCAL_MACHINE  = &H80000002

    strProfileList = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Set objRegistry = GetObject("winmgmts:root\default:StdRegProv")

    lRC = objRegistry.GetExpandedStringValue(HKEY_LOCAL_MACHINE, strProfileList & "\" & strProfileSid, "ProfileImagePath", strProfileImagePath)
    If (lRC = 0) And (Err.Number = 0) Then
        GetProfileImagePath = strProfileImagePath
    Else
        oLogging.CreateEntry "Function-GetProfileImagePath: Error retrieving ProfileImagePath Registry entry for " & strProfileSid, LogTypeError
    End If

End Function


Function GetAllProfileFolders()

    On Error Resume Next

    Dim arrProfileImagePathEntries()
    Const HKEY_LOCAL_MACHINE  = &H80000002
    intProfileFolderCount = 0


    ' Profile folders in ProfileList
    strProfileList = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Set objRegistry = GetObject("winmgmts:root\default:StdRegProv")

    lKeyRC = objRegistry.EnumKey(HKEY_LOCAL_MACHINE, strProfileList, sKeys)

    If (lKeyRC = 0) And (Err.Number = 0) Then
        For intKey = LBound(sKeys) To UBound(skeys)
            ReDim Preserve arrProfileImagePathEntries(intKey)

            strProfileImagePath = GetProfileImagePath(sKeys(intKey))
            'WScript.Echo "strProfileImagePath " & intKey & ": " & strProfileImagePath
            arrProfileImagePathEntries(intKey) = strProfileImagePath
            intProfileFolderCount = intProfileFolderCount + 1
        Next
    Else
        oLogging.CreateEntry "GetAllProfileFolders Function: Error enumerating Registry keys in " & strProfileList, LogTypeError
    End If

    ' All Users Profile folder
    ReDim Preserve arrProfileImagePathEntries(intProfileFolderCount)
    arrProfileImagePathEntries(intProfileFolderCount) = GetAllUsersProfile
    intProfileFolderCount = intProfileFolderCount + 1

    ' Default User Profile folder
    ReDim Preserve arrProfileImagePathEntries(intProfileFolderCount)
    arrProfileImagePathEntries(intProfileFolderCount) = GetDefaultUserProfile
    intProfileFolderCount = intProfileFolderCount + 1

    ' Public folder on Vista and higher
    If GetOSMajorMinorVersion >= 6.0 Then
        ReDim Preserve arrProfileImagePathEntries(intProfileFolderCount)
        arrProfileImagePathEntries(intProfileFolderCount) = GetPublicProfile
        intProfileFolderCount = intProfileFolderCount + 1
    End If

    GetAllProfileFolders = arrProfileImagePathEntries

End Function


Function GetProfilesDirectory()

    On Error Resume Next

    Dim strProfileList
    Dim objRegistry
    Dim lRC

    GetProfilesDirectory = null
    Const HKEY_LOCAL_MACHINE  = &H80000002

    strProfileList = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Set objRegistry = GetObject("winmgmts:root\default:StdRegProv")

    lRC = objRegistry.GetExpandedStringValue(HKEY_LOCAL_MACHINE, strProfileList, "ProfilesDirectory", strProfilesDirectory)
    If (lRC = 0) And (Err.Number = 0) Then
        GetProfilesDirectory = strProfilesDirectory
    Else
        oLogging.CreateEntry "Function-GetProfilesDirectory: Error retrieving ProfilesDirectory Registry entry from " & strProfileList, LogTypeError
    End If

End Function


Function GetAllUsersProfile()

    On Error Resume Next

    GetAllUsersProfile = null
    Const HKEY_LOCAL_MACHINE  = &H80000002

    strProfileList = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Set objRegistry = GetObject("winmgmts:root\default:StdRegProv")

    If GetOSMajorMinorVersion < 6.0 Then
        strvalueName = "AllUsersProfile"
    Else
        strvalueName = "ProgramData"
    End If

    lRC = objRegistry.GetExpandedStringValue(HKEY_LOCAL_MACHINE, strProfileList, strvalueName, strAllUsersProfile)
    If (lRC = 0) And (Err.Number = 0) Then
        If GetOSMajorMinorVersion < 6.0 Then
            GetAllUsersProfile = GetProfilesDirectory & "\" & strAllUsersProfile
        Else
            GetAllUsersProfile = strAllUsersProfile
        End If
    Else
        oLogging.CreateEntry "Function-GetAllUsersProfile: Error retrieving " & strvalueName & " Registry entry from " & strProfileList, LogTypeError
    End If

End Function


Function GetDefaultUserProfile()

    On Error Resume Next

    GetDefaultUserProfile = null
    Const HKEY_LOCAL_MACHINE  = &H80000002

    strProfileList = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Set objRegistry = GetObject("winmgmts:root\default:StdRegProv")

    If GetOSMajorMinorVersion < 6.0 Then
        strvalueName = "DefaultUserProfile"
    Else
        strvalueName = "Default"
    End If

    lRC = objRegistry.GetExpandedStringValue(HKEY_LOCAL_MACHINE, strProfileList, strvalueName, strDefaultUserProfile)
    If (lRC = 0) And (Err.Number = 0) Then
        If GetOSMajorMinorVersion < 6.0 Then
            GetDefaultUserProfile = GetProfilesDirectory & "\" & strDefaultUserProfile
        Else
            GetDefaultUserProfile = strDefaultUserProfile
        End If
    Else
        oLogging.CreateEntry "Function-GetDefaultUserProfile: Error retrieving " & strvalueName & " Registry entry from " & strProfileList, LogTypeError
    End If

End Function


Function GetPublicProfile()

    On Error Resume Next

    GetDefaultUserProfile = null
    Const HKEY_LOCAL_MACHINE  = &H80000002

    If GetOSMajorMinorVersion < 6.0 Then
        GetPublicProfile = ""
        Exit Function
    End If

    strProfileList = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    Set objRegistry = GetObject("winmgmts:root\default:StdRegProv")

    strvalueName = "Public"
    lRC = objRegistry.GetExpandedStringValue(HKEY_LOCAL_MACHINE, strProfileList, strvalueName, strPublicProfile)
    If (lRC = 0) And (Err.Number = 0) Then
        GetPublicProfile = strPublicProfile
    Else
        oLogging.CreateEntry "Function-GetPublicProfile: Error retrieving " & strvalueName & " Registry entry from " & strProfileList, LogTypeError
    End If

End Function


Function GetSystemProfile()

    On Error Resume Next

    GetSystemProfile = null

    GetSystemProfile = GetProfileImagePath("S-1-5-18")

End Function


'---------------------------------------------------------------------
' Get operating system Major.Minor version
'---------------------------------------------------------------------
Function GetOSMajorMinorVersion()

    On Error Resume Next

    GetOSMajorMinorVersion = ""
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

    Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
    For Each objOperatingSystem in colOperatingSystems
        GetOSMajorMinorVersion = Trim(Left(objOperatingSystem.Version, ((Len(objOperatingSystem.Version) - (InStrRev(objOperatingSystem.Version, "."))) - 1)))
    Next

End Function


'---------------------------------------------------------------------
' Get operating system caption
'---------------------------------------------------------------------
Function GetOSCaption()

On Error Resume Next

    GetOSCaption = ""
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

    Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
    For Each objOperatingSystem in colOperatingSystems
        GetOSCaption = objOperatingSystem.Caption
    Next

End Function


Function GetSpecialFolder(strCsidl)

    On Error Resume Next

    Const CSIDL_ADMINTOOLS = &H30
    Const CSIDL_ALTSTARTUP = &H1D
    Const CSIDL_APPDATA = &H1A
    Const CSIDL_BITBUCKET = &HA
    Const CSIDL_CDBURN_AREA = &H003b
    Const CSIDL_COMMON_ADMINTOOLS = &H2F
    Const CSIDL_COMMON_ALTSTARTUP = &H1E
    Const CSIDL_COMMON_APPDATA = &H23
    Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    Const CSIDL_COMMON_DOCUMENTS = &H2E
    Const CSIDL_COMMON_FAVORITES = &H1F
    Const CSIDL_COMMON_MUSIC = &H0035
    Const CSIDL_COMMON_OEM_LINKS = &H003a
    Const CSIDL_COMMON_PICTURES = &H0036
    Const CSIDL_COMMON_PROGRAMS = &H17
    Const CSIDL_COMMON_STARTMENU = &H16
    Const CSIDL_COMMON_STARTUP = &H18
    Const CSIDL_COMMON_TEMPLATES = &H2D
    Const CSIDL_COMMON_VIDEO = &H0037
    Const CSIDL_COMPUTERSNEARME = &H003d
    Const CSIDL_CONNECTIONS = &H31
    Const CSIDL_CONTROLS = &H3
    Const CSIDL_COOKIES = &H21
    Const CSIDL_DESKTOP = &H0
    Const CSIDL_DESKTOPDIRECTORY = &H10
    Const CSIDL_DRIVES = &H11
    Const CSIDL_FAVORITES = &H6
    Const CSIDL_FLAG_CREATE = &H8000
    Const CSIDL_FLAG_DONT_UNEXPAND = &H2000
    Const CSIDL_FLAG_DONT_VERIFY = &H4000
    Const CSIDL_FLAG_MASK = &HFF00
    Const CSIDL_FLAG_NO_ALIAS = &H1000
    Const CSIDL_FLAG_PER_USER_INIT = &H0800
    Const CSIDL_FLAG_PFTI_TRACKTARGET = &H4000  'CSIDL_FLAG_DONT_VERIFY
    Const CSIDL_FOLDER_MASK = &H00ff
    Const CSIDL_FONTS = &H14
    Const CSIDL_HISTORY = &H22
    Const CSIDL_INTERNET = &H1
    Const CSIDL_INTERNET_CACHE = &H20
    Const CSIDL_LOCAL_APPDATA = &H1C
    Const CSIDL_MY_DOCUMENTS = &H5
    Const CSIDL_MYMUSIC = &H000d
    Const CSIDL_MYPICTURES = &H27
    Const CSIDL_MYVIDEO = &H000e
    Const CSIDL_NETHOOD = &H13
    Const CSIDL_NETWORK = &H12
    Const CSIDL_PERSONAL = &H5   ' My Documents
    Const CSIDL_PRINTERS = &H4
    Const CSIDL_PRINTHOOD = &H1B
    Const CSIDL_PROFILE = &H28
    Const CSIDL_PROFILES = &H3e
    Const CSIDL_PROGRAM_FILES = &H26
    Const CSIDL_PROGRAM_FILES_COMMON = &H2B
    Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    Const CSIDL_PROGRAM_FILESX86 = &H2A
    Const CSIDL_PROGRAMS = &H2
    Const CSIDL_RECENT = &H8
    Const CSIDL_RESOURCES = &H38
    Const CSIDL_RESOURCES_LOCALIZED = &H39
    Const CSIDL_SENDTO = &H9
    Const CSIDL_STARTMENU = &HB
    Const CSIDL_STARTUP = &H7
    Const CSIDL_SYSTEM = &H25
    Const CSIDL_SYSTEMX86 = &H29
    Const CSIDL_TEMPLATES = &H15
    Const CSIDL_WINDOWS = &H24


    Set objWshShell	= CreateObject("WScript.Shell")

    strCsidl = UCase(strCsidl)
    strCsidlExpanded = objWshShell.ExpandEnvironmentStrings("%" & strCsidl & "%")

    'WScript.Echo "strCsidl: " & strCsidl & " - " & "strCsidlExpanded: " & strCsidlExpanded

    If (strCsidlExpanded) <> ("%" & strCsidl & "%") Then

        GetSpecialFolder = strCsidlExpanded

    Else

        If strCsidl = "ALLUSERSAPPDATA" Then strCsidl = "CSIDL_COMMON_APPDATA"
     
        If strCsidl = "ALLUSERSPROFILE" Then 
            GetSpecialFolder = GetAllUsersProfile
            Exit Function
        End If
     
        If strCsidl = "COMMONPROGRAMFILES" Then strCsidl = "CSIDL_PROGRAM_FILES_COMMON"
     
        If strCsidl = "COMMONPROGRAMFILES(X86)" Then strCsidl = "CSIDL_PROGRAM_FILES_COMMONX86"

        If strCsidl = "DEFAULTUSERPROFILE" Then 
            GetSpecialFolder = GetDefaultUserProfile
            Exit Function
        End If
     
        If strCsidl = "PROFILESFOLDER" Then 
            GetSpecialFolder = GetProfilesDirectory
            Exit Function
        End If
     
        If strCsidl = "PROGRAMDATA" Then 
            GetSpecialFolder = objWshShell.ExpandEnvironmentStrings("%PROGRAMDATA%")
            Exit Function
        End If

        If strCsidl = "PROGRAMFILES" Then strCsidl = "CSIDL_PROGRAM_FILES"
     
        If strCsidl = "PROGRAMFILES(X86)" Then strCsidl = "CSIDL_PROGRAM_FILESX86"
     
        If strCsidl = "SYSTEM" Then strCsidl = "CSIDL_SYSTEM"
     
        If strCsidl = "SYSTEM16" Then 
            GetSpecialFolder = objWshShell.ExpandEnvironmentStrings("%windir%\system")
            Exit Function
        End If
     
        If strCsidl = "SYSTEM32" Then strCsidl = "CSIDL_SYSTEM"

        If strCsidl = "SYSWOW64" Then strCsidl = "CSIDL_SYSTEMX86"

        If strCsidl = "SYSTEMPROFILE" Then 
            GetSpecialFolder = GetSystemProfile
            Exit Function
        End If


        If strCsidl = "SYSTEMROOT" Then  strCsidl = "CSIDL_WINDOWS"

        If strCsidl = "SYSTEMDRIVE" Then 
            GetSpecialFolder = objWshShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
            Exit Function
        End If

        If strCsidl = "WINDIR" Then 
            GetSpecialFolder = objWshShell.ExpandEnvironmentStrings("%WINDIR%")
            Exit Function
        End If


        strStatementToExecute = "strCsidlValue = " & strCsidl
        Execute strStatementToExecute

        Set objShell = CreateObject("Shell.Application")
        Set objFolder = objShell.NameSpace(strCsidlValue)

        Err.Clear
        GetSpecialFolder = objFolder.Self.Path
        If Err.Number <> 0 Then
            GetSpecialFolder = Err.Number
            oLogging.CreateEntry "Function-GetSpecialFolder: Cannot determine special folder for " & strCsidl, LogTypeInfo
        End If
    
    End If

End Function


Sub SetSpecialFolderEnvVars()

    On Error Resume Next
    
    Set objWshShell	= CreateObject("WScript.Shell")
    set objProcessEnv = objWshShell.Environment("PROCESS")
    
    objProcessEnv("CSIDL_ALTSTARTUP") = GetSpecialFolder("CSIDL_ALTSTARTUP")
    objProcessEnv("CSIDL_APPDATA") = GetSpecialFolder("CSIDL_APPDATA")
    objProcessEnv("CSIDL_BITBUCKET") = GetSpecialFolder("CSIDL_BITBUCKET")
    objProcessEnv("CSIDL_CDBURN_AREA") = GetSpecialFolder("CSIDL_CDBURN_AREA")
    objProcessEnv("CSIDL_COMMON_ADMINTOOLS") = GetSpecialFolder("CSIDL_COMMON_ADMINTOOLS")
    objProcessEnv("CSIDL_COMMON_ALTSTARTUP") = GetSpecialFolder("CSIDL_COMMON_ALTSTARTUP")
    objProcessEnv("CSIDL_COMMON_APPDATA") = GetSpecialFolder("CSIDL_COMMON_APPDATA")
    objProcessEnv("CSIDL_COMMON_DESKTOPDIRECTORY") = GetSpecialFolder("CSIDL_COMMON_DESKTOPDIRECTORY")
    objProcessEnv("CSIDL_COMMON_DOCUMENTS") = GetSpecialFolder("CSIDL_COMMON_DOCUMENTS")
    objProcessEnv("CSIDL_COMMON_FAVORITES") = GetSpecialFolder("CSIDL_COMMON_FAVORITES")
    objProcessEnv("CSIDL_COMMON_MUSIC") = GetSpecialFolder("CSIDL_COMMON_MUSIC")
    objProcessEnv("CSIDL_COMMON_OEM_LINKS") = GetSpecialFolder("CSIDL_COMMON_OEM_LINKS")
    objProcessEnv("CSIDL_COMMON_PICTURES") = GetSpecialFolder("CSIDL_COMMON_PICTURES")
    objProcessEnv("CSIDL_COMMON_PROGRAMS") = GetSpecialFolder("CSIDL_COMMON_PROGRAMS")
    objProcessEnv("CSIDL_COMMON_STARTMENU") = GetSpecialFolder("CSIDL_COMMON_STARTMENU")
    objProcessEnv("CSIDL_COMMON_STARTUP") = GetSpecialFolder("CSIDL_COMMON_STARTUP")
    objProcessEnv("CSIDL_COMMON_TEMPLATES") = GetSpecialFolder("CSIDL_COMMON_TEMPLATES")
    objProcessEnv("CSIDL_COMMON_VIDEO") = GetSpecialFolder("CSIDL_COMMON_VIDEO")
    objProcessEnv("CSIDL_COMPUTERSNEARME") = GetSpecialFolder("CSIDL_COMPUTERSNEARME")
    objProcessEnv("CSIDL_CONNECTIONS") = GetSpecialFolder("CSIDL_CONNECTIONS")
    objProcessEnv("CSIDL_CONTROLS") = GetSpecialFolder("CSIDL_CONTROLS")
    objProcessEnv("CSIDL_COOKIES") = GetSpecialFolder("CSIDL_COOKIES")
    objProcessEnv("CSIDL_DESKTOP") = GetSpecialFolder("CSIDL_DESKTOP")
    objProcessEnv("CSIDL_DESKTOPDIRECTORY") = GetSpecialFolder("CSIDL_DESKTOPDIRECTORY")
    objProcessEnv("CSIDL_DRIVES") = GetSpecialFolder("CSIDL_DRIVES")
    objProcessEnv("CSIDL_FAVORITES") = GetSpecialFolder("CSIDL_FAVORITES")
    objProcessEnv("CSIDL_FONTS") = GetSpecialFolder("CSIDL_FONTS")
    objProcessEnv("CSIDL_HISTORY") = GetSpecialFolder("CSIDL_HISTORY")
    objProcessEnv("CSIDL_INTERNET") = GetSpecialFolder("CSIDL_INTERNET")
    objProcessEnv("CSIDL_INTERNET_CACHE") = GetSpecialFolder("CSIDL_INTERNET_CACHE")
    objProcessEnv("CSIDL_LOCAL_APPDATA") = GetSpecialFolder("CSIDL_LOCAL_APPDATA")
    objProcessEnv("CSIDL_MY_DOCUMENTS") = GetSpecialFolder("CSIDL_MY_DOCUMENTS")
    objProcessEnv("CSIDL_MYMUSIC") = GetSpecialFolder("CSIDL_MYMUSIC")
    objProcessEnv("CSIDL_MYPICTURES") = GetSpecialFolder("CSIDL_MYPICTURES")
    objProcessEnv("CSIDL_MYVIDEO") = GetSpecialFolder("CSIDL_MYVIDEO")
    objProcessEnv("CSIDL_NETHOOD") = GetSpecialFolder("CSIDL_NETHOOD")
    objProcessEnv("CSIDL_NETWORK") = GetSpecialFolder("CSIDL_NETWORK")
    objProcessEnv("CSIDL_PERSONAL") = GetSpecialFolder("CSIDL_PERSONAL")
    objProcessEnv("CSIDL_PRINTERS") = GetSpecialFolder("CSIDL_PRINTERS")
    objProcessEnv("CSIDL_PRINTHOOD") = GetSpecialFolder("CSIDL_PRINTHOOD")
    objProcessEnv("CSIDL_PROFILE") = GetSpecialFolder("CSIDL_PROFILE")
    objProcessEnv("CSIDL_PROFILES") = GetSpecialFolder("CSIDL_PROFILES")
    objProcessEnv("CSIDL_PROGRAM_FILES") = GetSpecialFolder("CSIDL_PROGRAM_FILES")
    objProcessEnv("CSIDL_PROGRAM_FILES_COMMON") = GetSpecialFolder("CSIDL_PROGRAM_FILES_COMMON")
    objProcessEnv("CSIDL_PROGRAM_FILES_COMMONX86") = GetSpecialFolder("CSIDL_PROGRAM_FILES_COMMONX86")
    objProcessEnv("CSIDL_PROGRAM_FILESX86") = GetSpecialFolder("CSIDL_PROGRAM_FILESX86")
    objProcessEnv("CSIDL_PROGRAMS") = GetSpecialFolder("CSIDL_PROGRAMS")
    objProcessEnv("CSIDL_RECENT") = GetSpecialFolder("CSIDL_RECENT")
    objProcessEnv("CSIDL_RESOURCES") = GetSpecialFolder("CSIDL_RESOURCES")
    objProcessEnv("CSIDL_RESOURCES_LOCALIZED") = GetSpecialFolder("CSIDL_RESOURCES_LOCALIZED")
    objProcessEnv("CSIDL_SENDTO") = GetSpecialFolder("CSIDL_SENDTO")
    objProcessEnv("CSIDL_STARTMENU") = GetSpecialFolder("CSIDL_STARTMENU")
    objProcessEnv("CSIDL_STARTUP") = GetSpecialFolder("CSIDL_STARTUP")
    objProcessEnv("CSIDL_SYSTEM") = GetSpecialFolder("CSIDL_SYSTEM")
    objProcessEnv("CSIDL_SYSTEMX86") = GetSpecialFolder("CSIDL_SYSTEMX86")
    objProcessEnv("CSIDL_TEMPLATES") = GetSpecialFolder("CSIDL_TEMPLATES")
    objProcessEnv("CSIDL_WINDOWS") = GetSpecialFolder("CSIDL_WINDOWS")

    objProcessEnv("ALLUSERSAPPDATA") = GetSpecialFolder("ALLUSERSAPPDATA")
    objProcessEnv("ALLUSERSPROFILE") = GetSpecialFolder("ALLUSERSPROFILE")
    objProcessEnv("COMMONPROGRAMFILES") = GetSpecialFolder("COMMONPROGRAMFILES")
    objProcessEnv("COMMONPROGRAMFILES(X86)") = GetSpecialFolder("COMMONPROGRAMFILES(X86)")
    objProcessEnv("DEFAULTUSERPROFILE") = GetSpecialFolder("DEFAULTUSERPROFILE")
    objProcessEnv("PROFILESFOLDER") = GetSpecialFolder("PROFILESFOLDER")
    objProcessEnv("PROGRAMDATA") = GetSpecialFolder("PROGRAMDATA")
    objProcessEnv("PROGRAMFILES") = GetSpecialFolder("PROGRAMFILES")
    objProcessEnv("PROGRAMFILES(X86)") = GetSpecialFolder("PROGRAMFILES(X86)")
    objProcessEnv("SYSTEM") = GetSpecialFolder("SYSTEM")
    objProcessEnv("SYSTEM16") = GetSpecialFolder("SYSTEM16")
    objProcessEnv("SYSTEM32") = GetSpecialFolder("SYSTEM32")
    objProcessEnv("SYSWOW64") = GetSpecialFolder("SYSWOW64")
    objProcessEnv("SYSTEMPROFILE") = GetSpecialFolder("SYSTEMPROFILE")
    objProcessEnv("SYSTEMROOT") = GetSpecialFolder("SYSTEMROOT")
    objProcessEnv("SYSTEMDRIVE") = GetSpecialFolder("SYSTEMDRIVE")
    objProcessEnv("WINDIR") = GetSpecialFolder("WINDIR")

    'objWshShell.Run "cmd /k", 1, False

End Sub
