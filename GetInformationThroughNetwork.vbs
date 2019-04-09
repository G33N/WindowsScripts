Function CmdOut( pCmd )
  ' Run a command prompt command and get its output.
  ' pCmd <string> - A command prompt command.
  ' CmdOut <string> - The output of the command.
  ' Celiz Matias, https://github.com/G33N
  ' Function CmdOut(p):Dim w,e,r,o:Set w=CreateObject("WScript.Shell"):Set e=w.Exec("Cmd.exe"):e.StdIn.WriteLine p&" 2>&1":e.StdIn.Close:While(InStr(e.StdOut.ReadLine,">"&p)=0)::Wend:Do:o=e.StdOut.ReadLine:If(e.StdOut.AtEndOfStream)Then:Exit Do:Else:r=r&o&vbLf:End If:Loop:CmdOut=r:End Function
  Dim oWss, oExe, Return, Output
  Set oWss = CreateObject( "WScript.Shell" )
  Set oExe = oWss.Exec( "Cmd.exe" )
  Call oExe.StdIn.WriteLine( pCmd & " 2>&1" )
  Call oExe.StdIn.Close()
  While( InStr( oExe.StdOut.ReadLine, ">" & pCmd ) = 0 ) :: Wend
  Do : Output = oExe.StdOut.ReadLine()
    If( oExe.StdOut.AtEndOfStream )Then
      Exit Do
    Else 
      Return = Return & Output & vbLf
    End If
  Loop
  CmdOut = Return
End Function

' Put the full or mini class/sub/function in your script to use.
Function CmdOut(p):Dim w,e,r,o:Set w=CreateObject("WScript.Shell"):Set e=w.Exec("Cmd.exe"):e.StdIn.WriteLine p&" 2>&1":e.StdIn.Close:While(InStr(e.StdOut.ReadLine,">"&p)=0)::Wend:Do:o=e.StdOut.ReadLine:If(e.StdOut.AtEndOfStream)Then:Exit Do:Else:r=r&o&vbLf:End If:Loop:CmdOut=r:End Function

' if you run with cscript instead of wscript you can see the full output when using wscript.echo because msgbox has a character limit
Dim CmdRunCommand
'an answer of no will prompt user to input name of computer to scan and create PC file
strHost = InputBox("Enter the Asset Tag you wish to get data","Hostname"," ")
CmdRunCommand = CmdOut( "wmic /node:" & strHost & " MemoryChip get Manufacturer,Capacity,PartNumber,Speed,MemoryType,DeviceLocator,FormFactor" )
WScript.Echo CmdRunCommand