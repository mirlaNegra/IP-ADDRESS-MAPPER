On Error Resume Next
Dim pingText, dns, ip, maxTTL
Dim initiPos, finPos, i, j
Dim fso, ws, output, var
Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")
dns = InputBox("Write the web direction", "IAM")
Set output = ws.exec("cmd /c ping " & dns & " -n 1")
pingText = output.StdOut.ReadAll
initPos = Instr(pingText, "TTL=") + 4
maxTTL = CInt(Mid(pingText, initPos, 3))
Set var = fso.CreateTextFile (Replace(wscript.scriptfullname, "IAM.vbs", "Map.txt"))
var.writeLine dns
var.writeLine "TTL: " & maxTTL & Chr(13)
For i=1 To maxTTL
Set output = ws.exec("cmd /c ping " & dns & " -n 1 -i " & i)
pingText = output.StdOut.ReadAll
initPos = Instr(pingText, "Respuesta desde ") + 16
finPos = Instr(pingText, ":")
ip = Mid(pingText, initPos, initPos-finPos-5)
var.writeLine i & "- " & ip
Next