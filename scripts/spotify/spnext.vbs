set wshshell = wscript.CreateObject("wscript.shell")
wshshell.AppActivate "Spotify", NormalFocus
wshshell.sendkeys "%"
wshshell.sendkeys "%{P}{N}"