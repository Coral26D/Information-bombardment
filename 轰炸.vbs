Set WshShell=WScript.CreateObject("WScript.Shell")
WshShell.AppActivate"凉茶"
for i=1 to 3
WScript.Sleep 500
WshShell.SendKeys"^v"
WshShell.SendKeys i
WshShell.SendKeys"%s"
Next