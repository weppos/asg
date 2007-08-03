<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Option Explicit %>
<!--include virtual="/myasg/config.asp" -->
<!--#include file="config.asp" -->
<%

'/**
' * ASP Stats Generator - Powerful and reliable ASP website counter
' *
' * This file is part of the ASP Stats Generator package.
' * (c) 2003-2007 Simone Carletti <weppos@weppos.net>, All Rights Reserved
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
' * @copyright       2003-2007 Simone Carletti, All Rights Reserved
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
' -->	Sostituito da Avvertimento On Screen! 
'	Call AllowEntry("False", "False", "False", intAsgSecurity)

Dim strAsgSelectedIP	'IP Passato in QueryString
Dim asgOutputPage
Dim strNewFilteredIP	'IP da Filtrare
Dim strCommand			'Comando da Eseguire sull'IP
Dim blnUpdateCompleted	'TRUE se completato

'Richiama informazioni
strAsgSelectedIP = Trim(Request.QueryString("IP"))
strNewFilteredIP = Trim(Request.Form("FilterIP"))
blnUpdateCompleted = False

'Verifica per Inserimento IP nel Filtro
If Request.Form("Submit") = strAsgTxtUpdate AND Len(strNewFilteredIP) > 0 AND Session("AsgLogin") = "Logged" Then

	strCommand = Request.Form("Command")
	
	'Resetta ed Aggiungi
	If strCommand = "Reset" Then
		
		'Aggiornamento
		strAsgSQL = "UPDATE "&strAsgTablePrefix&"Config SET Filter_IP = '" & strNewFilteredIP & "'"
		objAsgConn.Execute(strAsgSQL)
		
		'Imposta a TRUE l'aggiornamento
		blnUpdateCompleted = True
	
	'Aggiungi alla lista
	ElseIf strCommand = "Add" Then

		'Richiama le informazioni sull'IP anche se in memoria
		'ma ci sarebbero troppi controlli da fare!
		strAsgSQL = "SELECT TOP 1 Filter_IP FROM "&strAsgTablePrefix&"Config "
		objAsgRs.Open strAsgSQL, objAsgConn
		
		'Se è vuoto aggiorna unicamente
		If objAsgRs.EOF Then
			
			'
		
		Else
			
			'Rivalorizza Variabile
			strAsgFilterIP = Trim(objAsgRs("Filter_IP"))
			'Pulisci spazi
			strAsgFilterIP = Replace(strAsgFilterIP, " ", "")
			
			'Controlla presenza " , " finale
			If Right(strAsgFilterIP, 1) = "," Then
				strNewFilteredIP = strAsgFilterIP & strNewFilteredIP
			'In mancanza aggiungi
			Else
				strNewFilteredIP = strAsgFilterIP & "," & strNewFilteredIP
			End If
			
		End If
		
		objAsgRs.Close
			
		'Aggiornamento
		strAsgSQL = "UPDATE "&strAsgTablePrefix&"Config SET Filter_IP = '" & strNewFilteredIP & "'"
		objAsgConn.Execute(strAsgSQL)
		
		'Imposta a TRUE l'aggiornamento
		blnUpdateCompleted = True
	
	End If

	
	'Se si utilizzano le variabili Application aggiornale
	If blnApplicationConfig Then
						
		'Aggiorna Variabili Application
		Application("strAsgFilterIP") = strNewFilteredIP
		'Forza il ricalcolo delle Application
		Application("blnConfig") = False
	
	End If

	
End If


'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgTxtFilterIPaddr & "&nbsp;" & strAsgTxtFor & "&nbsp;" & strAsgSelectedIP %> | ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2007 Carletti Simone, All Rights Reserved" />
<link href="stile.css" rel="stylesheet" type="text/css" />
<!--include virtual="/myasg/includes/inc_meta.asp" -->
<!--#include file="includes/inc_meta.asp" -->

<!-- 	ASP Stats Generator <%= strAsgVersion %> è una applicazione gratuita 
		per il monitoraggio degli accessi e dei visitatori ai siti web 
		creata e sviluppata da Simone Carletti.
		
		Puoi scaricarne una copia gratuita sul sito ufficiale http://www.weppos.com/ -->

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
Response.Write(vbCrLf & "			<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""bartitle"">" & UCase(strAsgTxtFilterIPaddr) & " : " & strAsgSelectedIP & "</td>")
Response.Write(vbCrLf & "		  </tr>")
Response.Write(vbCrLf & "		  <tr bgcolor=""" & strAsgSknTableLayoutBorderColour & """>")
Response.Write(vbCrLf & "			<td align=""center"" height=""1""></td>")
Response.Write(vbCrLf & "		  </tr>")
Response.Write(vbCrLf & "		</table>")

'CONTENUTO
'---------------------------------------------------|
Response.Write(vbCrLf & "		<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")

'Mostra solo se Loggato
If Session("AsgLogin") = "Logged" Then
		
		'CONTENUTO AGGIORNAMENTO
		'---------------------------------------------------|
	
	'Aggiornato
	If blnUpdateCompleted Then
		
		Response.Write(vbCrLf & "		  <tr valign=""middle"" bgcolor=""" & strAsgSknTableContBgColour & """ class=""smalltext"">")
		Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"" colspan=""2"" height=""16""><br />" & strAsgTxtUpdateSuccessfullyCompleted & "<br /><br /></td>")
		Response.Write(vbCrLf & "		  </tr>")
	
	Else	
		
		'CONTENUTO
		'---------------------------------------------------|
	
	'Manca l'IP in QueryString
	If NOT Len(strAsgSelectedIP) > 0 Then 
		
		Response.Write(vbCrLf & "		  <tr valign=""middle"" bgcolor=""" & strAsgSknTableContBgColour & """ class=""smalltext"">")
		Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"" colspan=""2"" height=""16""><br />" & strAsgTxtMissedDataToElab & "<br /><br /></td>")
		Response.Write(vbCrLf & "		  </tr>")
	
	'IP passato correttamente	
	Else
		
		Response.Write(vbCrLf & "		<form name=""frmFilterIp"" action=""popup_filter_ip.asp?IP=" & strAsgSelectedIP & """ method=""post"">")
		
		'Form IP
		Response.Write(vbCrLf & "		  <tr valign=""middle"" bgcolor=""" & strAsgSknTableContBgColour & """ class=""smalltext"">")
		Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""right"" width=""25%"" height=""16"">" & strAsgTxtIPAddress & "&nbsp;:&nbsp;&nbsp;</td>")
		Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""left""  width=""75%"">&nbsp;<input type=""text"" size=""15"" maxlenght=""20"" name=""FilterIP"" value=""" & strAsgSelectedIP & """ class=""normalform"" /></td>")
		Response.Write(vbCrLf & "		  </tr>")
		
		'Info RANGE
		Response.Write(vbCrLf & "		  <tr valign=""middle"" bgcolor=""" & strAsgSknTableContBgColour & """ class=""smalltext"">")
		Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""left""  width=""100%"" colspan=""2"">" & strAsgTxtInformationsToExitByIpRange & "</td>")
		Response.Write(vbCrLf & "		  </tr>")
		
		'Azione
		Response.Write(vbCrLf & "		  <tr valign=""middle"" bgcolor=""" & strAsgSknTableContBgColour & """ class=""smalltext"">")
		Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""right"" height=""16"">" & strAsgTxtAction & "&nbsp;:&nbsp;&nbsp;</td>")
		Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""left"">&nbsp;<select name=""Command"" class=""normalform"">")
		Response.Write(vbCrLf & "			  <option value=""Add"">" & strAsgTxtAddToList &"</option>")
		Response.Write(vbCrLf & "			  <option value=""Reset"">" & strAsgTxtResetAndAddToList &"</option>")
		Response.Write(vbCrLf & "			  </select>&nbsp;&nbsp;&nbsp;")
		Response.Write(vbCrLf & "		  <input type=""submit"" name=""Submit"" value=""" & strAsgTxtUpdate & """ class=""normalform"" />")
		Response.Write(vbCrLf & "		  </tr>")
		
		Response.Write(vbCrLf & "		</form>")
	
	End If

	'Fine condizione Aggiornato
	End If	


'Mostra se non Loggato
Else
	
		'AVVERTIMENTO
		'---------------------------------------------------|
	Response.Write(vbCrLf & "		  <tr valign=""middle"" bgcolor=""" & strAsgSknTableContBgColour & """ class=""smalltext"">")
	Response.Write(vbCrLf & "			<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"" widht=""25%"" height=""16""><br />" & strAsgTxtInsufficientPermission & "<br /><br /></td>")
	Response.Write(vbCrLf & "		  </tr>")
	
End If

'CONTENUTO (Chiusura)
'---------------------------------------------------|
Response.Write(vbCrLf & "		</table>")


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

Response.Write(vbCrLf & "<div class=""smalltext"" align=""center""><a href=""JavaScript:onClick=window.opener.location.href = window.opener.location.href; window.close();"" class=""linksmalltext"" title=""" & strAsgTxtCloseWindow & """>" & strAsgTxtCloseWindow & "</a></div>")

Response.Write(vbCrLf & "</body></html>")

%>