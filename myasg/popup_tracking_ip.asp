<%@LANGUAGE="VBSCRIPT"%>
<% Option Explicit %>
<!--#include file="config.asp" -->
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
Call AllowEntry("True", "False", "False", intAsgSecurity)

Dim dtmAsgElabDate		'Data più recente in elaborazione
Dim strAsgSelectedIP	'IP Passato in QueryString
Dim asgOutputPage

'Richiama informazioni
strAsgSelectedIP = Trim(Request.QueryString("IP"))

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgTxtIPTracking & "&nbsp;" & strAsgTxtFor & "&nbsp;" & strAsgSelectedIP %> | powered by ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2008 Carletti Simone, All Rights Reserved" />
<meta name="generator" content="ASP Stats Generator <%= strAsgVersion %>" /> <!-- leave this for stats -->

<!--#include file="includes/html-head.asp" -->

<!--
  ASP Stats Generator (release <%= strAsgVersion %>) is a free software package
  completely written in ASP programming language, for real time visitor tracking.
  Get your own copy for free at http://www.asp-stats-com/ !
-->

</head>

<%

'HEADER
'---------------------------------------------------|
Response.Write(vbCrLf & "<body bgcolor=""" & strAsgSknPageBgColour & """ background=""" & strAsgSknPageBgImage & """>")

	'CONTENITORE
	'---------------------------------------------------|
Response.Write(vbCrLf & "<table width=""95%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""0"" bgcolor=""" & strAsgSknTableLayoutBorderColour & """>")
Response.Write(vbCrLf & "  <tr><td>")
Response.Write(vbCrLf & "	<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
Response.Write(vbCrLf & "	  <tr><td bgcolor=""" & strAsgSknTableLayoutBgColour & """ background=""" & strAsgSknPathImage & strAsgSknTableLayoutBgImage & """>")

'TITOLO BARRA
'---------------------------------------------------|
Response.Write(vbCrLf & "		<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
Response.Write(vbCrLf & "		  <tr align=""center"" valign=""middle"">")
Response.Write(vbCrLf & "			<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""bartitle"">" & UCase(strAsgTxtIPTracking) & " : " & strAsgSelectedIP & "</td>")
Response.Write(vbCrLf & "		  </tr>")
Response.Write(vbCrLf & "		  <tr bgcolor=""" & strAsgSknTableLayoutBorderColour & """>")
Response.Write(vbCrLf & "			<td align=""center"" height=""1""></td>")
Response.Write(vbCrLf & "		  </tr>")
Response.Write(vbCrLf & "		</table>")

'CONTENUTO
'---------------------------------------------------|

	'TITOLO
	'---------------------------------------------------|
Response.Write(vbCrLf & "		<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
Response.Write(vbCrLf & "		  <tr valign=""middle"" bgcolor=""" & strAsgSknTableContBgColour & """ class=""normaltitle"">")
Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"" widht=""25%"" height=""16"">" & UCase(strAsgTxtTime) & "</td>")
Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"" widht=""75%"">" & UCase(strAsgTxtPage) & "</td>")
Response.Write(vbCrLf & "		  </tr>")
	
	'CONTENUTO
	'---------------------------------------------------|

'Manca l'IP in QueryString
If NOT Len(strAsgSelectedIP) > 0 Then 
	
	Response.Write(vbCrLf & "		  <tr valign=""middle"" bgcolor=""" & strAsgSknTableContBgColour & """ class=""smalltext"">")
	Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"" colspan=""2"" height=""16""><br />" & strAsgTxtMissedDataToElab & "<br /><br /></td>")
	Response.Write(vbCrLf & "		  </tr>")

'IP passato correttamente	
Else
	
	'Richiama le informazioni sull'IP
	strAsgSQL = "SELECT Details_ID, Data, Page FROM "&strAsgTablePrefix&"Detail WHERE IP = '" & strAsgSelectedIP & "' ORDER BY Data DESC "
	objAsgRs.Open strAsgSQL, objAsgConn
	
	'Informazioni Presenti
	If NOT objAsgRs.EOF Then
		
		dtmAsgElabDate = ""
		
		'Stampa tutto
		Do While Not objAsgRs.EOF
		
			'Se la data non corrisponde a quella in memoria stampa la nuova data
			If FormatOutTimeZone(objAsgRs("Data"), "Date") <> dtmAsgElabDate Then
				dtmAsgElabDate = FormatOutTimeZone(objAsgRs("Data"), "Date")
				Response.Write(vbCrLf & "		  <tr valign=""middle"" bgcolor=""" & strAsgSknTableContBgColour & """ class=""smalltext"">")
				Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""left"" colspan=""2"" height=""16""><img src=""images/arrow_small_dx.gif"" align=""absmiddle"" border=""0"" />&nbsp;" & strAsgTxtDetails & "&nbsp;" & dtmAsgElabDate & "</td>")
				Response.Write(vbCrLf & "		  </tr>")
			End If
		
			asgOutputPage = objAsgRs("Page")
			'Taglia tutto il prefisso sito + http:// se non è una pagina sconosciuta
			If Mid(asgOutputPage, 1, Len(strAsgSiteURLremote)) = strAsgSiteURLremote Then asgOutputPage = Mid(asgOutputPage, Len(strAsgSiteURLremote) + 1) 
		
			Response.Write(vbCrLf & "		  <tr valign=""middle"" bgcolor=""" & strAsgSknTableContBgColour & """ class=""smalltext"">")
			Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"">" & FormatOutTimeZone(objAsgRs("Data"), "Time") & "</td>")
			Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""left"">")

			'Verifica la pagina e mostra o meno
			'una icona standard di corrispondenza dominio.
			Response.Write(ChooseDomainIcon(objAsgRs("Page"), "classic" ))
						
			'TAGLIA STRINGHE LUNGHE
			'Se la stringa supera i 55 caratteri inserisci ... in mezzo e accorcia
			'Caso "Nessun Raggruppamento" - Max 55 Caratteri
			Response.Write(StripValueTooLong(asgOutputPage, 55, 25, 25))
			
			Response.Write("</a></td>")
			Response.Write(vbCrLf & "		  </tr>")
	
		objAsgRs.MoveNext
		Loop
	
	'Informazioni NON Presenti
	Else
				
		'// Row - No current record			
		Call BuildTableContNoRecord(2, "standard")
				
	End If
	
	objAsgRs.Close

End If

'CONTENUTO (Chiusura)
'---------------------------------------------------|
Response.Write(vbCrLf & "		</table>")


'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing


'FOOTER
'---------------------------------------------------|
Response.Write(vbCrLf & "		<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")

	'SPACER
	'---------------------------------------------------|
Response.Write(vbCrLf & "		  <tr bgcolor=""" & strAsgSknTableLayoutBorderColour & """>")
Response.Write(vbCrLf & "			<td align=""center"" height=""1""></td>")
Response.Write(vbCrLf & "		  </tr>")

'***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
'***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
	Response.Write(vbCrLf & "		  <tr align=""center"" valign=""middle"">")
	Response.Write(vbCrLf & "			<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer"">ASP Stats Generator [" & strAsgVersion & "] - &copy; 2003-2006 <a href=""http://www.weppos.com/"" class=""linkfooter"">weppos</a></td>")
	Response.Write(vbCrLf & "		  </tr>")
'***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
'***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

	'CONTENITORE (Chiusura)
	'---------------------------------------------------|
Response.Write(vbCrLf & "		</table>")
Response.Write(vbCrLf & "	  </td></tr>")
Response.Write(vbCrLf & "	</table>")
Response.Write(vbCrLf & "  </td></tr>")
Response.Write(vbCrLf & "</table><br />")

Response.Write(vbCrLf & "<div class=""smalltext"" align=""center""><a href=""JavaScript:onClick=window.close();"" class=""linksmalltext"" title=""" & strAsgTxtCloseWindow & """>" & strAsgTxtCloseWindow & "</a></div>")

Response.Write(vbCrLf & "</body></html>")

%>