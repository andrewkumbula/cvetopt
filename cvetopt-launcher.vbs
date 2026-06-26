' cvetopt — как обычное Windows-приложение.
' 1) Стартует сервер в фоне (без чёрного окна).
' 2) Открывает окно программы (Edge/Chrome --app).
' 3) Пока окно открыто — сервер работает.
' 4) Закрыли окно — сервер останавливается.
'
' Ярлык: wscript.exe "C:\путь\cvetopt\cvetopt-launcher.vbs"
Option Explicit

Const APP_URL = "http://127.0.0.1:8000/"
Const HEALTH_URL = "http://127.0.0.1:8000/api/state"
Const START_TIMEOUT_SEC = 90

Dim shell, fso, here, batPath, stopPath, waited

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

here = fso.GetParentFolderName(WScript.ScriptFullName)
batPath = here & "\cvetopt.bat"
stopPath = here & "\cvetopt-stop.bat"

If Not fso.FileExists(batPath) Then
  MsgBox "Не найден cvetopt.bat рядом с лаунчером:" & vbCrLf & batPath, vbCritical, "cvetopt"
  WScript.Quit 1
End If

shell.CurrentDirectory = here

If Not IsServerUp() Then
  shell.Environment("Process")("CVETOPT_HIDDEN") = "1"
  shell.Environment("Process")("CVETOPT_NO_BROWSER") = "1"
  shell.Run """" & batPath & """", 0, False

  waited = 0
  Do While waited < START_TIMEOUT_SEC
    WScript.Sleep 1000
    waited = waited + 1
    If IsServerUp() Then Exit Do
  Loop

  If Not IsServerUp() Then
    MsgBox _
      "cvetopt: сервер не ответил за " & START_TIMEOUT_SEC & " с." & vbCrLf & vbCrLf & _
      "Проверьте вручную:" & vbCrLf & batPath, _
      vbExclamation, "cvetopt"
    WScript.Quit 1
  End If
End If

' Ждём, пока пользователь закроет окно приложения.
OpenAsAppAndWait APP_URL

StopServer

Function IsServerUp()
  On Error Resume Next
  Dim http, ok
  Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
  http.Open "GET", HEALTH_URL, False
  http.setTimeouts 2000, 2000, 2000, 2000
  http.Send
  ok = (Err.Number = 0 And http.Status = 200)
  On Error GoTo 0
  IsServerUp = ok
End Function

Sub OpenAsAppAndWait(url)
  Dim exe, arg, candidates, i, pf, pf86
  pf = shell.ExpandEnvironmentStrings("%ProgramFiles%")
  pf86 = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%")

  candidates = Array( _
    pf & "\Microsoft\Edge\Application\msedge.exe", _
    pf86 & "\Microsoft\Edge\Application\msedge.exe", _
    pf & "\Google\Chrome\Application\chrome.exe", _
    pf86 & "\Google\Chrome\Application\chrome.exe" _
  )

  arg = "--app=" & url

  For i = 0 To UBound(candidates)
    exe = candidates(i)
    If fso.FileExists(exe) Then
      shell.Run """" & exe & """ " & arg, 1, True
      Exit Sub
    End If
  Next

  shell.Run url, 1, True
End Sub

Sub StopServer()
  If fso.FileExists(stopPath) Then
    shell.Environment("Process")("CVETOPT_QUIET") = "1"
    shell.Run """" & stopPath & """", 0, True
  End If
End Sub
