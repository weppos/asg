<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
Call AllowEntry("False", "False", "False", intAsgSecurity)


'Inserimento record
If Request.Form("Impostazioni") = strAsgTxtUpdate AND Request.QueryString("Exc") = "Upd" Then

	Dim strAsgPswNuova
	Dim strAsgPswConferma
	Dim blnAsgErrore
	Dim blnAsgPswIns
	
	blnErrore = False
	blnAsgPswIns = False
	
	strAsgPswNuova = Trim(Request.Form("PswNuova"))
	strAsgPswConferma = Trim(Request.Form("PswConferma"))
	
	If IsNumeric(Request.Form("Protezione")) Then intAsgSecurity = CInt(Request.Form("Protezione"))
	
	strAsgPswNuova = CleanInput(strAsgPswNuova)
	strAsgPswConferma = CleanInput(strAsgPswConferma)
	
	If "[]" & strAsgPswNuova <> "[]" Then
	
		If strAsgPswNuova = strAsgPswConferma Then
			blnAsgPswIns = true
		Else
			blnErrore = True
		End If
	
	End If
	
	'Nessun errore rilevato. Procedi con inserimento.
	If blnErrore = False Then
	
		If blnAsgPswIns = True Then
			strAsgSQL = "UPDATE "&strAsgTablePrefix&"Config SET Sito_PSW = '" & strAsgPswNuova & "', Stats_Protezione = " & intAsgSecurity & ""
		Else
			strAsgSQL = "UPDATE "&strAsgTablePrefix&"Config SET Stats_Protezione = " & intAsgSecurity & ""
		End If
		
		objAsgConn.Execute(strAsgSQL)
	
		'Se si utilizzano le variabili Application aggiornale
		If blnApplicationConfig Then
					
			'Aggiorna Variabili Application
			If blnAsgPswIns = True Then Application("strAsgSitePsw") = strAsgPswNuova
			Application("intAsgSecurity") = CInt(intAsgSecurity)
			'Forza il ricalcolo delle Application
			Application("blnConfig") = False
		
		End If
		
		'Reset Server Objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing
		
		'Reindirizza per rivalorizzare dati
		Response.Redirect("settings_security.asp?Msg=Upd")
	
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
<title><%= strAsgSiteName %> | ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2004 Carletti Simone" />
<link href="stile.css" rel="stylesheet" type="text/css" />

<!-- 	ASP Stats Generator <%= strAsgVersion %> è una applicazione gratuita 
		per il monitoraggio degli accessi e dei visitatori ai siti web 
		creata e sviluppata da Simone Carletti.
		
		Puoi scaricarne una copia gratuita sul sito ufficiale http://www.weppos.com/ -->

</head>
<!--#include file="includes/header.asp" -->
		<form action="settings_security.asp?Exc=Upd" name="frmSicurezza" method="post">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= strAsgTxtSecuritySettings %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <% If Request.QueryString("Msg") = "Upd" Then %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" align="center" height="15"><br /><strong><%= strAsgTxtUpdateSuccessfullyCompleted %></strong><br /><br /></td>		  
		  </tr>
		  <% End If %>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtEntryPassword) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" align="center" height="15">
			<% 
			If blnErrore = True Then 
				Response.Write("<strong>" & strAsgTxtAttentionPasswordNotMatching & "</strong>") 
			Else 
			    Response.Write(strAsgTypeOnlyToChangePassword)
			End If %>
			</td>		  
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right" width="50%"><%= strAsgTxtNewPassword %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left" width="50%">&nbsp;<input type="password" name="PswNuova" value="" size="20" maxlength="20" /></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtConfirmPassword %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<input type="password" name="PswConferma" value="" size="20" maxlength="20" /></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtStatsProtection) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtStatsProtectionLevel %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">
				&nbsp;<input type="radio" name="Protezione" value="0" <% If intAsgSecurity = 0 Then Response.Write "checked" %> /><%= strAsgTxtNone %>
				&nbsp;<input type="radio" name="Protezione" value="1" <% If intAsgSecurity = 1 Then Response.Write "checked" %> /><%= strAsgTxtLimited %>
				&nbsp;<input type="radio" name="Protezione" value="2" <% If intAsgSecurity = 2 Then Response.Write "checked" %> /><%= strAsgTxtFull %>
			</td>
		  </tr>
		  <%
					
			'// Row - End table spacer			
			Call BuildTableContEndSpacer(2)
	
		  %>
		  <tr class="normaltitle">
			<td colspan="2" align="center"><br /><input type="submit" name="Impostazioni" value="<%= strAsgTxtUpdate %>" /></td>
		  </tr>
		</table>
		</form>
<%

'Footer
Response.Write(vbCrLf & "		<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
'// Row - Footer Border Line
Call BuildFooterBorderLine()

'***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
'***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write(vbCrLf & "		  <tr align=""center"" valign=""middle"">")
Response.Write(vbCrLf & "			<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer"">ASP Stats Generator [" & strAsgVersion & "] - &copy; 2003-2006 <a href=""http://www.weppos.com/"" class=""linkfooter"">weppos</a>")
If blnAsgElabTime Then Response.Write(" - " & strAsgTxtThisPageWasGeneratedIn & "&nbsp;" & FormatNumber(Timer() - startAsgElab, 4) & "&nbsp;" & strAsgTxtSeconds)
Response.Write(						"</td>")
Response.Write(vbCrLf & "		  </tr>")
'***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
'***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write(vbCrLf & "		</table>")
Response.Write(vbCrLf & "	  </td></tr>")
Response.Write(vbCrLf & "	</table>")
Response.Write(vbCrLf & "  </td></tr>")
Response.Write(vbCrLf & "</table>")

%>
<!--#include file="includes/footer.asp" -->
</body></html>