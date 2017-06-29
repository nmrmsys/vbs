@if(0)==(0)echo off
  @echo off
  cscript //nologo "%~f0?.wsf" %*
  goto :EOF
@end
<job><script language="VBScript">
  With WScript
    For i = 0 To .Arguments.Count - 1
      If sCmd <> "" Then
        sCmd = sCmd & " "
      End If
      sCmd = sCmd & .Arguments(i)
    Next
    sCmd = Replace(sCmd, "'", """")
    sCmd = Replace(sCmd, ";", vbCrLf)
    'MsgBox sCmd
    Do Until .StdIn.AtEndOfStream
      S0 = .StdIn.ReadLine & vbCrLf
      Execute sCmd
      If S0 <> "" And InStr(S0,vbCrLf) = 0 Then
        S0 = S0 & vbCrLf
      End If
      sOut = sOut & S0
    Loop
    'MsgBox sOut
    .StdOut.Write sOut
  End With
</script></job>
