
sdobj = new ActiveXObject ("Scripting.{8C6E39B0-BC52-43C0-B3E2-A5F83FAEB897}")
sdobj.Invoke (1, 2, 3)
b = sdobj.Invoke
sdobj.Invoke = 123
sdobj.Invoke = b
sdobj.Invoke ("str") = new String ("abc")
WScript.Echo ("Before delete sdobj.")
delete sdobj
WScript.Echo ("After delete sdobj.")
