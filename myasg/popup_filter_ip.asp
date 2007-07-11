<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Option Explicit %>
<!--include virtual="/myasg/config.asp" -->
<!--#include file="config.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
' Copyright 2003-2006 - Carletti Simone										'
'-------------------------------------------------------------------------------'
'																				'
'	Autore:																		'
'	--------------------------													'
'	Simone Carletti (weppos)													'
'																				'
'	Collaboratori 																'
'	[che ringrazio vivamente per l'impegno ed il tempo dedicato]				'
'	--------------------------													'
'	@ imente 			- www.imente.it | www.imente.org						'
'	@ ToroSeduto		- www.velaforfun.com									'
'																				'
'	Hanno contribuito															'
'	[anche a loro un grazie speciale per le idee apportate]						'
'	--------------------------													'
'	@ Gli utenti del forum con consigli e segnalazioni							'
'	@ subxus (suggerimento generazione grafica dei report)						'
'																				'
'	Verifica le proposte degli utenti, implementate o da implementare al link	'
'	http://www.weppos.com/forum/forum_posts.asp?TID=140&PN=1					'
'																				'
'-------------------------------------------------------------------------------'
'																				'
'	Informazioni sulla Licenza													'
'	--------------------------													'
'	Questo è un programma gratuito; potete modificare ed adattare il codice		'
'	(a vostro rischio) in qualsiasi sua parte nei termini delle condizioni		'
'	della licenza che lo accompagna.											'
'																				'
'	Non è consentito utilizzare l'applicazione per conseguire ricavi 			'
'	personali, distribuirla, venderla o diffonderla come una propria 			'
'	creazione anche se modificata nel codice, senza un esplicito e scritto 		'
'	consenso dell'autore.														'
'																				'
'	Potete modificare il codice sorgente (a vostro rischio) per adattarlo 		'
'	alle vostre esigenze o integrarlo nel sito; nel caso le funzioni possano	'
'	essere di utilità pubblica vi invitiamo a comunicarlo per poterle 			'
'	implementare in una futura versione e per contribuire allo sviluppo 		'
'	del programma.																'
'																				'
'	In nessun caso l'autore sarà responsabile di danni causati da una 			'
'	modifica, da un uso non corretto o da un uso qualsiasi 						'
'	dell'applicazione.															'
'																				'
'	Nell'utilizzo devono rimanere intatte tutte le informazioni sul 			'
'	copyright; è possibile modificare o rimuovere unicamente le indicazioni 	'
'	espressamente specificate.													'
'																				'
'	Numerose ore sono state impiegate nello sviluppo del progetto e, anche 		'
'	se non vincolante ai fini dell'uso, sarebbe gratificante l'inserimento		'
'	di un link all'applicazione sul vostro sito.								'
'																				'
'	NESSUNA GARANZIA															'
'	------------------------- 													'
'	Questo programma è distribuito nella speranza che possa essere utile ma 	'
'	senza GARANZIA DI ALCUN GENERE.												'
'	L'utente si assume tutte le responsabilità nell'uso.						'
'																				'
'-------------------------------------------------------------------------------'

'********************************************************************************'
'*																				*'	
'*	VIOLAZIONE DELLA LICENZA													*'
'*	 																			*'
'*	L'utilizzo dell'applicazione violando le condizioni di licenza comporta la 	*'
'*	perdita immediata della possibilità d'uso ed è PERSEGUIBILE LEGALMENTE!		*'
'*																				*'
'********************************************************************************'


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
<meta name="copyright" content="Copyright (C) 2003-2004 Carletti Simone" />
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