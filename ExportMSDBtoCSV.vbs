Option Explicit
 'Modify by:Rommel C. Pineda
 'Date 2.10.2016
Dim TextExportPath  'Where to put the CSV
Dim TextExportFile  'This will be generated
Dim db              'Database to export from
Dim cn              'ADODB Connection
Dim strCon          'Connection strings
Dim table_name		'Table 
Dim site_code		'Site Code see defined site code from Main DB table name 'artar_web.dbo.site_tbl'
Dim un
Dim pw
Dim ip
Dim port 
'Set local variable value
'=========================================================
db = "C:\YOUR_PATH\YOUR_MSDB.mdb"
TextExportPath = "D:\"
table_name ="AttendanceLOG"
site_code = "SMC"
'=========================================================


TextExportFile = NewFileName(TextExportPath)
 
Set cn = CreateObject("ADODB.Connection")
 
cn.Open _
    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source =" & db
 
 'iterate and insert the date
cn.Execute "SELECT * INTO [text;HDR=Yes;Database=" & TextExportPath & _
   "]." & TextExportFile & " FROM " & table_name
  
  'Function to create the file
Function NewFileName(TextExportPath)

	Dim fso
	Dim NewFileTemp
	Dim a, i
	Const OverwriteExisting = True
	Set fso = CreateObject("Scripting.FileSystemObject")
	 
	 NewFileTemp = site_code & ".csv" ' this will assign filename to export csv
	'a = fs.FileExists(TextExportPath & NewFileTemp)
	 
	 'if file already exist delete and create new file.
	 if fso.FileExists(TextExportPath&NewFileTemp) Then
		'delete
		fso.DeleteFile(TextExportPath & NewFileTemp)	
	 end if
	  
	NewFileName = NewFileTemp
End Function