' Только сервер в фоне, без окна программы (сервер не останавливается сам).
' Для обычной работы используйте cvetopt-launcher.vbs — окно закрыли, сервер выключился.
' Этот файл — если нужен постоянный сервер (Планировщик заданий при входе в Windows).
' Остановить сервер: cvetopt-stop.bat
Option Explicit

Dim shell, fso, here, batPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

here = fso.GetParentFolderName(WScript.ScriptFullName)
batPath = here & "\cvetopt.bat"

shell.CurrentDirectory = here
shell.Environment("Process")("CVETOPT_HIDDEN") = "1"
shell.Environment("Process")("CVETOPT_NO_BROWSER") = "1"

' 0 = окно скрыто; False = не ждать завершения (сервер работает дальше).
shell.Run """" & batPath & """", 0, False
