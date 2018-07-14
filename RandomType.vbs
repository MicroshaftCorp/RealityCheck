set shellobj = CreateObject("WScript.Shell")
shellobj.run "cmd"

do

shellobj.sendkeys "D"
wscript.sleep 200
Shellobj.sendkeys "o "
Shellobj.sendkeys "g "
Shellobj.sendkeys "e "
Shellobj.sendkeys "{ENTER}"
wscript.sleep 200

loop
