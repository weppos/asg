<%@LANGUAGE="VBSCRIPT"%>
<%

'Controllo errori
On Error Resume Next

%>
<!--#include file="config.asp" -->
<!--#include file="includes/functions_count.asp" -->
<%

'Buffer FALSE per mostrare l'update
'Response.Buffer = False

Dim strAsgUpdateVersion
strAsgUpdateVersion = Request.QueryString("version")

'							========================================
'---------------------------   	Funzioni di gestione del file		-------------------------------------
'							========================================


'---------------------------------------------------
'	FUNZIONE DISEGNO LAYOUT
'---------------------------------------------------
'//	In funzione per evitare lunghe file di codice
Public function CreateLayout(ByVal layout)
	
	If layout = "OpenRow" Then
		
		'LAYOUT TABELLA
		'//	Layout
		Response.Write(vbCrLf & "  <tr class=""normaltext"">")
		Response.Write(vbCrLf & "    <td width=""100%"" align=""center""><br />")
	
	ElseIf layout = "CloseRow" Then
		
		'LAYOUT TABELLA
		'//	Layout
		Response.Write(vbCrLf & "      <br /><br /><br />")
		Response.Write(vbCrLf & "      <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">")
		Response.Write(vbCrLf & "        <tr><td height=""2"" bgcolor=""#FDD353""></td></tr>")
		Response.Write(vbCrLf & "      </table>")

		Response.Write(vbCrLf & "    </td>")
		Response.Write(vbCrLf & "  </tr>")
	
	End If
	
end function '//	In funzione per evitare lunghe file di codice
Public function CreateLayoutRowNote(ByVal outputtext)
		
	'LAYOUT TABELLA
	'//	Layout
	Response.Write(vbCrLf & "  <tr class=""smalltext"">")
	Response.Write(vbCrLf & "    <td valign=""middle"" align=""center"">NOTE:</td>")
	Response.Write(vbCrLf & "    <td><div align=""justify"">" & outputtext & "</div></td>")
	Response.Write(vbCrLf & "  </tr>")

end function '//	In funzione per evitare lunghe file di codice
Public function CreateLayoutRowSpacer()
		
	'LAYOUT TABELLA
	'//	Layout
	Response.Write(vbCrLf & "  <tr class=""normaltext"">")
	Response.Write(vbCrLf & "    <td valign=""middle"" align=""center"" width=""100%"" colspan=""2"" height=""10""></td>")
	Response.Write(vbCrLf & "  </tr>")

end function '							========================================
'---------------------------   		Esecuzione Update a 2.1			-------------------------------------
'							========================================


'---------------------------------------------------
'	[Tabella]	Creazione tabella 'Traceroute'
'---------------------------------------------------
'// Crea tabella
'

	'// Funzione di Aggiornamento
	Sub v_2_1_edit_column_optckeckupdate()
		
		'
		strAsgSQL = "ALTER TABLE "&strAsgTablePrefix&"Config DROP Column Opt_Check_Update"
		objAsgConn.Execute(strAsgSQL)
		
		'
		strAsgSQL = "ALTER TABLE "&strAsgTablePrefix&"Config ADD Column Opt_Check_Update INTEGER"
		objAsgConn.Execute(strAsgSQL)
		
		'
		strAsgSQL = "UPDATE "&strAsgTablePrefix&"Config SET Opt_Check_Update = " & Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & ""
		objAsgConn.Execute(strAsgSQL)
			
		'Notifica Aggiornamento
		Response.Write("<span class=""menutitle"">COMPLIMENTI: The field 'Opt_Check_Update' has beed modified!</span>")
		
	End Sub


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Update ASP Stats Generator | Version <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

<style type="text/css">
<!--
body, div, p, tr, td {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10pt;
	font-weight: normal;
	color: #000000;
}

.menutitle {
	color: #0066CC;
	line-height: 140%;
}

a {
	color: #0066CC;
	text-decoration: none;
}
a:hover {
	color: #0066CC;
	text-decoration: underline;
}
a:visited  {
	color: #0066CC;
	text-decoration: none;
}
a:visited:hover {
	color: #0066CC;
	text-decoration: underline;
}

.normaltitle, h2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	color: #0066CC;
	font-weight: bold;
	line-height: 140%;
	text-align: center;
	font-variant: small-caps;
}

h3 {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 14px;
	color: #0066CC;
	font-weight: bold;
	line-height: 120%;
	font-variant: small-caps;
}

#content {
	padding-left: 10px;
	padding-right: 10px;
}
-->
</style>
</head>

<body>
<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr><td height="50" valign="middle"><img src="images/asg_setupfile.gif" width="276" height="28" border="0" alt="ASP Stats Generator" /></td></tr>
  <tr><td height="2" bgcolor="#FDD353"></td></tr>
  <tr><td height="15" valign="middle" align="right" class="x-smalltext">Update ASP Stats Generator - v. <%= strAsgVersion %></td></tr>
</table>

<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
<%

'							========================================
'---------------------------   	Gestione Update	  ---> 	to 2.1		-------------------------------------
'							========================================

If strAsgUpdateVersion = "2.1" Then
 
'---------------------------------------------------
'	[Query]		Creazione campo 'SERP_Page'
'---------------------------------------------------
'// Esegui aggiornamento
If Request.QueryString("btnUpdate") = "edit_column_optckeckupdate" Then

	
	'LAYOUT TABELLA
	Call CreateLayout("OpenRow")

	'Correzione
	Call v_2_1_edit_column_optckeckupdate()

	'LAYOUT TABELLA
	Call CreateLayout("CloseRow")


End If	


End If

'							========================================
'---------------------------   	Controllo gestione degli Errori		-------------------------------------
'							========================================


If Err.Number <> 0 Then

	
	'LAYOUT TABELLA
	Call CreateLayout("OpenRow")

	'Errore
	Response.Write("<span class=""menutitle"">ATTENZION - The following error has occured: </span><br />")
    Response.Write("<div align=""left""><strong>Number:</strong> " & Err.Number & "<br />")
    Response.Write("<strong>Description:</strong> " & Err.Description & "<br />")
    Response.Write("<strong>Source:</strong> " & Err.Source & "<br />")
    Err.Clear

	'LAYOUT TABELLA
	Call CreateLayout("CloseRow")


End If

%>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="4" align="center">
  <tr>
    <td width="100%" colspan="2"><div align="justify"><br />
	The file will execute the changes to upgrade your database to the last version.
	Please, follow and press the execute buttons in the order they are displayed:</div>
	</td>
  </tr>
</table>
<%

'							========================================
'---------------------------   	Gestione Update	  ---> 	to 2.0		-------------------------------------
'							========================================

%>
<form action="update.asp" method="get" name="frmUpdate21">
<input type="hidden" name="version" value="2.1" />
<table width="95%" border="0" cellspacing="1" cellpadding="4" align="center">
  <tr class="menutitle">
    <td width="100%" colspan="2"><div align="justify"><strong>:: Update to version 2.1.1</strong></div></td>
  </tr>
<%

	'Spaziatore
	Call CreateLayoutRowSpacer()

%>
  <tr>
    <td width="20%" valign="middle" align="center"><input type="Submit" name="btnUpdate" value="edit_column_optckeckupdate" class="smalltext" /></td>
    <td width="80%" align="left"><span class="menutitle">Edit the '<strong>Opt_Check_Update</strong>' field of the <strong>'Config'</strong> table</span>
	</td>
  </tr>
<%

	'Note
	'Call CreateLayoutRowNote("")
	'Spaziatore
	Call CreateLayoutRowSpacer()

%>
</table>
</form>


<%

'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing
	
'Se si utilizzano le variabili Application aggiornale
If blnApplicationConfig Then
						
	'Forza il ricalcolo delle Application
	Application("blnConfig") = False
	
End If
	

%>

</body>
</html>
