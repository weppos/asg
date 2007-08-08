<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'



'-----------------------------------------------------------------------------------------
' Rinomina File
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	27.12.2003 |
' Comment:	
'-----------------------------------------------------------------------------------------
Function RinominaFile(fileFrom, fileTo)

	Dim objFso, objFile
	
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.GetFile(fileFrom)
	objFile.Copy fileTo, True
	objFile.Delete True
		
	Set objFso = Nothing
	Set objFile = Nothing

End Function


'-----------------------------------------------------------------------------------------
' Ripristina File
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	27.12.2003 | 27.12.2003
' Comment:	
'-----------------------------------------------------------------------------------------
Function RipristinaFile(fileFrom, fileTo)

	Dim objFso
	
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")
	objFso.CopyFile fileTo, fileTo & ".bak", true 
	objFso.DeleteFile fileTo
	objFso.MoveFile fileFrom, fileTo
	objFso.DeleteFile fileTo & ".bak"

	Set objFso = Nothing

End Function


'-----------------------------------------------------------------------------------------
' Compact Access database
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	| 18.07.2004
' Comment:	
'-----------------------------------------------------------------------------------------
Function CompactAccessDatabase()
	
	Dim strAsgDb, strAsgDbTo
	Dim objAsgJro
	
	strAsgDb = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strAsgMapPath
	strAsgDbTo = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strAsgMapPathTo
	
	set objAsgJro = CreateObject("jro.JetEngine") 
	objAsgJro.CompactDatabase strAsgDb, strAsgDbTo
	Set objAsgJro = Nothing 
	
	'Return function
	CompactAccessDatabase = true
	
End Function


'-----------------------------------------------------------------------------------------
' Optimize MySQL database
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	| 18.07.2004
' Comment:	
'-----------------------------------------------------------------------------------------
Function databaseMySqlOptimize(table, conn)
	
	conn.Execute("OPTIMIZE TABLE " & ASG_TABLE_PREFIX & table)
	
	'Return function
	databaseMySqlOptimize = true
		
End Function


'-----------------------------------------------------------------------------------------
' Get the size of a selected file in the site directory.
'
' @param 	path - the file path	
' @return	the file size converted in Kb.
'
' @since 2.0
'-----------------------------------------------------------------------------------------
public function getFileSize(path)

	Dim objFso
	Dim objFile
	Dim tmpSize
	
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")
	
	' If the file exists get the file size
	if objFso.FileExists(path) then

		Set objFile = objFso.GetFile(path)
		tmpSize = objFile.Size
		' From byte to Kb
		tmpSize = (tmpSize / 1024) & "&nbsp;" & TXT_Kb
		Set objFile = Nothing

	end if
	
	Set objFso = Nothing
	
	' Returns the file size
	getFileSize = tmpSize

end function


'/**
' * Checks whether a file exists.
' * 
' * @param		(string) filename - the file path and name
' * @return 	(boolean) true if the file specified by filename exists,
' *				false otherwise.
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function file_exists(filename)
	
	Dim fso
	Dim return
	
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(Server.MapPath(filename)) then
		return = true
	else
		return = false
	end If
	Set fso = Nothing
	file_exists = return

end function

%>