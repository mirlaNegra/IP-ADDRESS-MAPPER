On Error Resume Next
Dim pingText, dns, ip, maxTTL
Dim initiPos, finPos, i
Dim fso, ws, output, sdtIn, stdOut
Set stdIn = WScript.StdIn
Set stdOut = WScript.StdOut
Set ws = CreateObject("WScript.Shell")
stdOut.Write "Write the web direction: " 
dns = stdIn.ReadLine
Set output = ws.exec("cmd /c ping " & dns & " -n 1")
pingText = output.StdOut.ReadAll
initPos = Instr(pingText, "TTL=") + 4
maxTTL = CInt(Mid(pingText, initPos, 3))
stdOut.WriteLine string(3, Chr(13))
stdOut.WriteLine string(20, "*")
stdOut.WriteLine "TTL: " & maxTTL
For i=1 To maxTTL
Set output = ws.exec("cmd /c ping " & dns & " -n 1 -i " & i)
pingText = output.StdOut.ReadAll
initPos = Instr(pingText, "Respuesta desde ") + 16
finPos = Instr(pingText, ": ")
If initPos >= 70 Then
ip = Mid(pingText, initPos, finPos-initPos)
Else 
ip = "Timeout for this request"
End If
stdOut.WriteLine i & "- " & ip
Next
stdOut.WriteLine string(20, "*")
WScript.Sleep 600000