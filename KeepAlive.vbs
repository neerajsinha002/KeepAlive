On Error Resume Next
vCounter = 0
LogFolderPath="C:\Users\Public\Keep Alive Trails"
MaxFileAge=2
vbScriptToCheck = "KeepAlive.vbs"
Set objSWbemServices = GetObject("WinMgmts:Root\Cimv2")
Set colProcess = objSWbemServices.ExecQuery _
("Select * From Win32_Process where name = 'wscript.exe'")
For Each objProcess In colProcess
    If objProcess.CommandLine<> "" Then
        m=split(objProcess.CommandLine,"""")
        n=m(3)
        o=InStr(1,n,vbscriptToCheck)
        if o>0 Then
            vCounter=vCounter+1
        End If
    End If
Next

If vCounter>1 Then
    Wscript.Quit
End If

Set objSWbemServices = Nothing
Set colProcess  = Nothing

libLogFilePath=LogFolderPath &"\"&"KeepAlive Audit Trails" & "_" & Month(now) & "_" & Day(now) & "_" & Year(now) & ".log"
libDescription="Script Started."
Call logToFile(LogFolderPath,libDescription,libLogFilePath)
Set ws = CreateObject("WScript.Shell")

Do
    ws.SendKeys "{F15}"
    Wscript.Sleep 2000
    ws.SendKeys "{F14}"

    libLogFilePath=LogFolderPath &"\"&"KeepAlive Audit Trails" & "_" & Month(now) & "_" & Day(now) & "_" & Year(now) & ".log"
    Wscript.Sleep 50000

libDescription="Keys Pressed Successfully."
Call logToFile(LogFolderPath,libDescription,libLogFilePath)
Call DeleteFiles(LogFolderPath,MaxFileAge)

Loop


Function logToFile(LogFolderPath,libDescription,libLogFilePath)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If NOT (FSO.FolderExists(LogFolderPath)) Then
        FSO.CreateFolder(LogFolderPath)
    end If

    If FSO.FileExists(libLogFilePath) Then
        Set objFile31=FSO.OpenTextFile(libLogFilePath,8)
    Else
        Set objFile31=FSO.CreateTextFile(libLogFilePath)
    End If

    objFile31.WriteLine "(" & Now & ")" & "|" & "|" & libDescription

    objFile31.Close
    Set FSO = Nothing
    Set objFile31 = Nothing
End Function

Function DeleteFiles(ByVal sFolder, MaxFileAge)
    today = Date
    Set oFileSys = WScript.CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFileSys.GetFolder(sFolder)
    Set aFiles = oFolder.Files

    For Each file In aFiles
        dFileCreated = FormatDateTime(file.DateLastModified,"2")
        If DateDiff("d",dFileCreated,today)>MaxFileAge Then
            file.Delete(True)
        End If
    Next ' file

    Set oFileSys = Nothing
    Set oFolder = Nothing
    Set aFiles = Nothing
End Function ' DeleteFiles
