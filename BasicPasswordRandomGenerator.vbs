Option Explicit
 'Author :Lemmor
 'Date 26.3.2017
Dim Gpw 

Const ForReading = 1, ForWriting = 2
Const OverwriteExisting = True
Dim cn,dchar,Ndchar,mchar,Nmchar,Lstr,Rstr,Dos,Tres,Filename,TextExportPath,Npw,TextExportFile,FileExtension,fso,f,NewFileTemp,myfile


'*****************Change Variable value ************************
'TextExportPath = "C:\Windows\SYSVOL\sysvol\artar.com.sa\Policies\{20F6F4D4-E103-4981-8378-61B03E295E6B}\Machine\Scripts\Startup\"'change this static location
TextExportPath = "H:\"  'Location
Filename = "password"	'Filename 
FileExtension=".bat" 	'example .bat 
Gpw= "ADM!N@253"		'default password,change it
'****************************************************************



dchar= Day(Date)'get the day todays date
mchar = Month(Date)'get the month todays date

'Npw = Lstr & mArray(Cint(Ndchar)) & mArray(Cint(NMchar)) & Rstr 'str,index,length * New password
Npw =Gpw & LPad(dchar, 2, "0") & LPad(mchar, 2, "0") 

	Set fso = CreateObject("Scripting.FileSystemObject")	 
	 NewFileTemp = filename & FileExtension ' this will assign filename to export csv	 
	 'if file already exist delete and create new file.
	 if fso.FileExists(TextExportPath&NewFileTemp) Then		
		fso.DeleteFile(TextExportPath & NewFileTemp)		
	 end if
	 
	 'Open and write text,no need to create when writing text
	 'Concat the network command
	 Set f = fso.OpenTextFile(TextExportPath&NewFileTemp, 8,true)
		f.WriteLine "net user administrator " & Npw & "/logonpasswordchg:yes"


'just padding for string
Function LPad(s, l, c)
  Dim n : n = 0
  If l > Len(s) Then n = l - Len(s)
  LPad = String(n, c) & s
End Function
