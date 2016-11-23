'Script to read file and fix issues with too many CRLF and output new file
' Written by : Chris Boylan  11/2016

Option explicit
Dim fso,logfile,dir,logline, kblist,kbfile,kbline
dim temp,display, intcount,outputfile, oShell
dim wshNetwork,strComputer

temp=0
intcount=0
Const ForReading = 1
set oshell=WScript.CreateObject ("WScript.Shell")
set fso=createobject("scripting.filesystemobject")
'select the correct file for your test

kbfile="C:\Users\boylan.cj\Documents\E&D New Stuff\temp delete me\Commentslist.txt"

' read in the computer name to make the results file unique
Set wshNetwork = WScript.CreateObject("WScript.Network")


outputfile="C:\Users\boylan.cj\Documents\E&D New Stuff\temp delete me\CommentslistOUt.txt"
if fso.FileExists(kbfile)Then
	'msgbox ("YEAH File, " & dir & " , found")
	Else
	msgbox ("File, " & kbfile & " , not found")
    wscript.quit
end if
if fso.FileExists(outputfile)Then
	    call fso.deletefile( outputfile, true )
end if
'open file

 Set kblist=fso.OpenTextFile(kbfile,ForReading)
 do until kblist.atEndOfstream = true  ' use of stream to read entire file
	        kbline=kblist.ReadLine   ' reads next line in the file
            'msgbox("Line - " & kbline)
            'temp= check(kbline)
            'msgbox(temp)
            if check(kbline) = 0 then

                display = display &  kbline & vbcrlf
            Else
                display = display & kbline & " "
            end if
'           
        loop


'msgbox display
writetofile(display)
'oShell.run "notepad.exe "+outputfile
'set oShell = Nothing

'subs
'Function to determine if CRLF is needed or not
function check(byval logline)
    	dim intPLace, intplace2,intplace3
        intplace=0
'    msgbox (logline)	
    intplace =instr(logline,vbTab)+1
    if intplace >2 then
        intplace2 = instr(intplace,logline,vbTab)+1
    end if
    if intplace >3 then
        intplace3 = instr(intplace2,logline,vbTab)+1
    end if
    'msgbox intplace & " = " & intplace2 & " = " & intplace3 & " = " &  logline
	if intplace3 >5 then
            'msgbox (instr(intplace3,logline,")"))
		if instr(intplace3,logline,"(amPortfolio)") >1 then
            check =0
            exit function
        Elseif instr(intplace3,logline,"()") >1 then
            check =0
            exit function
        Else
            'check=replace(logline,vbCRLF," ")
            check=1
        end if
    Else
'        'check=replace(logline,vbCRLF," ")
'        if instr(logline,")") >1 then
'            check=0
'        Else
'            check=1
'        End if
        if instr(intplace,logline,"(amPortfolio)") >1 then
            check =0
            exit function
        Elseif instr(intplace,logline,"()") >1 then
            check =0
            exit function
        Else
            'check=replace(logline,vbCRLF," ")
            check=1
        end if
        
	End if
    
    
end function

function writetofile (byval display)
dim fsoMain,fsoFile
Const ForAppending=8
set fsoMain =createobject("scripting.filesystemobject")
set fsoFile=fso.OpenTextFile(outputfile,ForAppending,true)
fsoFile.write display
fsoFile.close
end function
