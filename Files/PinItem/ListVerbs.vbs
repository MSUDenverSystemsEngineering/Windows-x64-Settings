' Windows Script Host Sample Script
'
' ------------------------------------------------------------------------
'               Copyright (C) 2009 Microsoft Corporation
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Microsoft and the author
' have no warranty, obligations or liability for any Sample Application Files.
' ------------------------------------------------------------------------
'********************************************************************
'*
'* File:           ListVerbs.vbs
'* Date:           04/08/2009
'* Version:        1.0.0
'*
'* Main Function:  List the Shell verbs for shell objects
'*
'* Usage:  cscript ListVerbs.vbs "<path to exe>" "<path to exe>" ...
'*
'* Copyright (C) 2009 Microsoft Corporation
'*
'* Revisions:
'*
'* 1.0.0 - 04/08/2009 - Created.
'*
'********************************************************************

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")


Set objArgs = WScript.Arguments
For i = 0 to objArgs.Count - 1

    WScript.Echo "Verbs for item: " & objArgs(i)
    WScript.Echo "================" & String(Len(objArgs(i)), "=")

    strFolderPath = objFSO.GetParentFolderName(objArgs(i))
    strFileName = objFSO.GetFileName(objArgs(i))

    Set objFolder = objShell.Namespace(strFolderPath)
    Set objFolderItem = objFolder.ParseName(strFileName)
     
    Set colVerbs = objFolderItem.Verbs
    For Each objVerb in colVerbs
        Wscript.Echo objVerb
    Next

    WScript.Echo ""
    WScript.Echo ""

Next

