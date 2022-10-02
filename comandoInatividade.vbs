set object = CreateObject("WScript.Shell")
Do
 WScript.Sleep(3*60*1000)
 object.SendKeys("{F13}")
Loop