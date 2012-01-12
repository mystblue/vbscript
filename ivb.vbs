' File : ivb.vbs
' Usage : CScript ivb.vbs
Do While True
    WScript.StdOut.Write(">>> ")
    ln = Wscript.StdIn.ReadLine
    If LCase(Trim(ln)) = "exit" Then Exit Do
    On Error Resume Next
    Err.Clear
    Execute ln
    If Err.Number <> 0 Then WScript.Echo(Err.Description)
    On Error Goto 0
Loop