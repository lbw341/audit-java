Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\Administrator\Desktop\新建文件夹 (5)\audit-java"
WshShell.Run "cmd /c java -jar target\module-audit-1.0.0.jar", 0, False
