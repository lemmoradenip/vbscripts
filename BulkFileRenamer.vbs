'#####################
'Creator: Lemmor Adenip
'Date: 20.12.2018
'Function: RENAME FILES INSIDE THE  PATH  e.i Image copy.png = Image.png
' JUST REPLACE THE TEXT FOR VARIABLE (FindText,ReplaceWith and sDir)
'#####################

 Dim oFS  : Set oFS  = CreateObject("Scripting.FileSystemObject")
 Set fso = CreateObject("Scripting.FileSystemObject")
  Dim sDir : sDir     = "D:\GEXPO\SHOP\TLM PHOTOS\All\"' Source Path & Destination Path sample  D:\Folder\
  Const FindText = " copy" ' Find the text to be replace
  Const ReplaceWith = "" 'text you want to add or remove
  Dim oFile
  For Each oFile In oFS.GetFolder(sDir).Files   	   	  
		oFile.Name = Replace(oFile.Name,FindText,ReplaceWith) ' Replace " copy" with the file you want to remove from the files
	Next