Option Explicit

 'Author :Lemmor Adenip
 'Date 26.3.2014

Const ForReading = 1, ForWriting = 2
Const OverwriteExisting = True
Dim cn,dchar,Ndchar,mchar,Nmchar,Lstr,Rstr,Dos,Tres,Filename,TextExportPath,Npw,TextExportFile,FileExtension,fso,f,NewFileTemp,myfile
TextExportPath = "H:\"'change this static location
Filename = "ABC"	'Filename 
FileExtension=".bat" 'example .bat 

Dim mArray (9)'array with single dimension 9 elements
mArray(0)="a"
mArray(1)="b"
mArray(2)="c"
mArray(3)="d"
mArray(4)="e"
mArray(5)="f"
mArray(6)="g"
mArray(7)="h"
mArray(8)="i"
mArray(9)="j"

Dim Gpw 

Gpw= "Admin@123"'default password,change it
'get last number of char

dchar= Day(Date)'get the day todays date
mchar = Month(Date)'get the month todays date
Ndchar = right(dchar,1)'get the last number 
Nmchar =right(mchar,1) 'get the last number

Lstr = Mid(Gpw,1,1)'first char
Dos = Mid(Gpw,2,1) 'second char
Tres = Mid(Gpw,3,1) 'third char
Rstr = Mid(Gpw,4,Len(Gpw))'4th index onwards

Npw = Lstr & mArray(Cint(Ndchar)) & mArray(Cint(NMchar)) & Rstr 'str,index,length * New password
	Set fso = CreateObject("Scripting.FileSystemObject")	 
	 NewFileTemp = filename & FileExtension ' this will assign filename to export csv	 
	 'if file already exist delete and create new file.
	 if fso.FileExists(TextExportPath&NewFileTemp) Then		
		fso.DeleteFile(TextExportPath & NewFileTemp)		
	 end if
	 
	 'Open and write text,no need to create when writing text
	 Set f = fso.OpenTextFile(TextExportPath&NewFileTemp, 8,true)
		f.WriteLine Npw




