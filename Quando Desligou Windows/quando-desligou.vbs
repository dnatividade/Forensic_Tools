'Determine the last shutdown time and date in Windows 10 and earlier
'v.:1.0

strValueName1 = "HKLM\SYSTEM\CurrentControlSet\Control\Windows\" & "ShutdownTime"
strValueName2 = "HKLM\System\CurrentControlSet\Control\TimeZoneInformation\TimeZoneKeyName"

Set oShell = CreateObject("WScript.Shell")

Ar = oShell.RegRead(strValueName1)
strTimeZone = oShell.RegRead(strValueName2)

Term = Ar(7)*(2^56) + Ar(6)*(2^48) + Ar(5)*(2^40) + Ar(4)*(2^32) + Ar(3)*(2^24) + Ar(2)*(2^16) + Ar(1)*(2^8) + Ar(0)
Days = Term/(1E7*86400)

WScript.Echo "Data/Hora do desligamento = " & CDate(DateSerial(1601, 1, 1) + Days) & " UTC" _ 
& vbCrLf & "Time Zone: " & strTimeZone _
& vbCrLf & vbCrLf & _ 
"Lembre-se de ajustar as configuracoes de fuso horario."

'FONTE:
'adaptado de :https://www.winhelponline.com/blog/how-to-determine-the-last-shutdown-date-and-time-in-windows/
