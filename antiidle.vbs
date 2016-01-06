dim wshell
set wshell=createobject("wscript.shell")
while 1
    wscript.sleep 60 * 10 * 1000
    wshell.sendkeys "{SCROLLLOCK}"
    wscript.sleep 1 * 1000
    wshell.sendkeys "{SCROLLLOCK}"
wend
