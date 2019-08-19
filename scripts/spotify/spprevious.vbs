set wshshell = wscript.CreateObject("wscript.shell")
wshshell.AppActivate "Spotify", NormalFocus
wshshell.sendkeys "%"
wshshell.sendkeys "%{P}{r}"
wshshell.sendkeys "{ENTER}"
wscript.sleep "200"
wshshell.sendkeys "%{P}{r}"
wshshell.sendkeys "{ENTER}"