Dim objShell
Set objShell=CreateObject("WScript.Shell")
strExpression="(New-Object Media.SoundPlayer D:\Programs\HampdenKeepalive\20Hz_44100Hz_16bit_10sec_soft.wav).PlaySync();"
strCMD="powershell -sta -noProfile -NonInteractive  -NoLogo -Command " & Chr(34) &_
"&{" & strExpression &"}" & Chr(34) 
objShell.Run strCMD,0