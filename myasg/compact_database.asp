<%@LANGUAGE="VBSCRIPT"%>
<% Option Explicit %>
<!--#include file="config.asp" -->
<!--#include file="includes/inc_array_table.asp" -->
<%

'/**
' * ASP Stats Generator - Powerful and reliable ASP website counter
' *
' * This file is part of the ASP Stats Generator package.
' * (c) 2003-2008 Simone Carletti <weppos@weppos.net>, All Rights Reserved
' *
' * 
' * COPYRIGHT AND LICENSE NOTICE
' *
' * The License allows you to download, install and use one or more free copies of this program 
' * for private, public or commercial use.
' * 
' * You may not sell, repackage, redistribute or modify any part of the code or application, 
' * or represent it as being your own work without written permission from the author.
' * You can however modify source code (at your own risk) to adapt it to your specific needs 
' * or to integrate it into your site. 
' *
' * All links and information about the copyright MUST remain unchanged; 
' * you can modify or remove them only if expressly permitted.
' * In particular the license allows you to change the application logo with a personal one, 
' * but it's absolutly denied to remove copyright information,
' * including, but not limited to, footer credits, inline credits metadata and HTML credits comments.
' *
' * For the full copyright and license information, please view the LICENSE.htm
' * file that was distributed with this source code.
' *
' * Removal or modification of this copyright notice will violate the license contract.
' *
' *
' * @category        ASP Stats Generator
' * @package         ASP Stats Generator
' * @author          Simone Carletti <weppos@weppos.net>
' * @copyright       2003-2008 Simone Carletti
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id$
' */
 
'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


'// ATTENZIONE! Protezione statistiche.
'	Modificare solo se necessario e se sicuri.
'	Impostazioni errate possono compromettere la privacy.
Call AllowEntry("False", "False", "False", intAsgSecurity)


'Inizializza variabili
Dim strAsgMsg		'Messaggio Reset


'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing


'-----------------------------------------------------------------------------------------
' Rinomina File
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	27.12.2003 |
' Commenti:	
'-----------------------------------------------------------------------------------------
function RinominaFile(fileFrom, fileTo)

	Dim objFso, objFile
	
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.GetFile(fileFrom)
	objFile.Copy fileTo, True
	objFile.Delete True
		
	Set objFso = Nothing
	Set objFile = Nothing

end function '-----------------------------------------------------------------------------------------
' Compatta il Database
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	|
' Commenti:	
'-----------------------------------------------------------------------------------------
function CompactAccessDatabase()
	
	Dim strAsgDb, strAsgDbTo
	Dim objAsgJro
	
	strAsgDb = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strAsgMapPath
	strAsgDbTo = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strAsgMapPathTo
	
	set objAsgJro = CreateObject("jro.JetEngine") 
	objAsgJro.CompactDatabase strAsgDb, strAsgDbTo
	Set objAsgJro = Nothing 
	
end function 'Compatta il database dopo i reset	
Call CompactAccessDatabase()

'Dopo aver compattato il database
'ripristina la versione precedente
'con quella compattata
Call RinominaFile(strAsgMapPathTo, strAsgMapPath)


'Nel caso si siano verificati errori valorizza una variabile
'e mostrali poi a video continuando l'esecuzione
If err <> 0 then 
    strAsgMsg = strAsgMsg & strAsgTxtError & ": <br />" & err.description & "<br />" 
Else 
    strAsgMsg = strAsgMsg & "<br />" & strAsgTxtDatabaseSuccessfullyCompactedOn & "<br /><span class=""notetext"">" & strAsgMapPathTo & "</span><br />"
    strAsgMsg = strAsgMsg & strAsgTxtDatabaseSuccessfullyRenamedTo & "<br /><span class=""notetext"">" & strAsgMapPath & "</span><br />"
End if 


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | powered by ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="copyright" content="Copyright (C) 2003-2008 Carletti Simone, All Rights Reserved" />
<meta name="generator" content="ASP Stats Generator <%= strAsgVersion %>" /> <!-- leave this for stats -->

<!--#include file="includes/html-head.asp" -->

<!--
  ASP Stats Generator (release <%= strAsgVersion %>) is a free software package
  completely written in ASP programming language, for real time visitor tracking.
  Get your own copy for free at http://www.asp-stats-com/ !
-->

</head>
<!--#include file="includes/header.asp" -->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgCompactDatabase) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="70%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" align="center" height="16"><%= UCase(strAsgTxtExecutionReport) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="100%" align="center"><br />
			<%= strAsgMsg %><br /><br />
			</td>
		  </tr>
		<%
				
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(2)

		%>
		</table><br />
<%

' Footer
Response.Write(vbCrLf & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
'// Row - Footer Border Line
Call BuildFooterBorderLine()

' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write("<tr align=""center"" valign=""middle"">")
Response.Write("<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer""><a href=""http://www.asp-stats.com/"" class=""linkfooter"" title=""ASP Stats Generator Homepage"">ASP Stats Generator</a> [" & strAsgVersion & "] - &copy; 2003-2008 <a href=""http://www.weppos.com/"" class=""linkfooter"" title=""Weppos.com Homepage"">weppos</a>")
If blnAsgElabTime Then Response.Write(" - " & strAsgTxtThisPageWasGeneratedIn & "&nbsp;" & FormatNumber(Timer() - startAsgElab, 4) & "&nbsp;" & strAsgTxtSeconds)
Response.Write("</td>")
Response.Write("</tr>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write("</table>")
Response.Write("</td></tr>")
Response.Write("</table>")
Response.Write("</td></tr>")
Response.Write("</table>")

%>
<!-- footer -->
<!--#include file="includes/footer.asp" -->
</body></html>