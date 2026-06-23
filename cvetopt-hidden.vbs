' Запуск cvetopt без видимого окна консоли.
' Двойной клик по этому файлу: стартует cvetopt.bat в скрытом окне,
' сервер работает в фоне, браузер откроется сам на http://127.0.0.1:8000.
' Остановить сервер: cvetopt-stop.bat (или Диспетчер задач → процессы python/uvicorn).
Option Explicit

Dim shell, fso, here, batPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

here = fso.GetParentFolderName(WScript.ScriptFullName)
batPath = here & "\cvetopt.bat"

shell.CurrentDirectory = here
' Флаг для батника: в скрытом режиме не делать pause (иначе зависнет невидимо).
shell.Environment("Process")("CVETOPT_HIDDEN") = "1"

' 0 = окно скрыто; False = не ждать завершения (сервер работает дальше).
shell.Run """" & batPath & """", 0, False
