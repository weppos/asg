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


'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

'Verifica Password
If Request.Form("Login") = strAsgTxtLogin Then

	Dim strAsgPassword
	Dim blnAsgErrore
	
	blnAsgErrore = False
	
	strAsgPassword = Trim(Request.Form("Password"))
	strAsgPassword = CleanInput(strAsgPassword)

	'Verifica
	If LCase(strAsgPassword) = LCase(strAsgSitePsw) Then
	
		'1° Versione --> Uso variabili di sessione
		'prossima implementazione cookie
		
		Session("AsgLogin") = "Logged"
		
	Else

		blnAsgErrore = True
		Session.Contents.Remove("AsgLogin")
		
	End If

End If

'Logout
If Request.QueryString("Logout") = "True" Then Session.Contents.Remove("AsgLogin")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2004 Carletti Simone" />
<%	If Session("AsgLogin") = "Logged" AND Len(Request.QueryString("backto")) > 0 Then %>
<meta http-equiv="Refresh" content="3;url=<%= Request.QueryString("backto") %>">
<%	ElseIf Session("AsgLogin") = "Logged" AND NOT Len(Request.QueryString("backto")) > 0 Then %>
<meta http-equiv="Refresh" content="3;url=statistiche.asp">
<%	End If %>
<link href="stile.css" rel="stylesheet" type="text/css" />

<!-- 	ASP Stats Generator <%= strAsgVersion %> è una applicazione gratuita 
		per il monitoraggio degli accessi e dei visitatori ai siti web 
		creata e sviluppata da Simone Carletti.
		
		Puoi scaricarne una copia gratuita sul sito ufficiale http://www.weppos.com/ -->

</head>
<!--include virtual="/myasg/includes/header.asp" -->
<!--#include file="includes/header.asp" -->
		<form action="login.asp?backto=<%= Server.URLEncode(Request.QueryString("backto")) %>" name="frmLogin" method="post">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableBarBgColour %>" valign="middle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" align="center" height="20" class="bartitle"><%= UCase(strAsgTxtLogin) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		<% If Session("AsgLogin") <> "Logged" Then %>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtEntryPassword) %></td>
		  </tr>
			  <% If blnAsgErrore Then %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" align="center" height="16"><br /><strong><%= strAsgTxtWrongPassword %></strong><br /><br /></td>		  
		  </tr>
			  <% End If %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="50%" align="right"><%= strAsgTxtTypePassword %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="50%" align="left">&nbsp;<input type="password" name="Password" value="" size="20" maxlength="20" /></td>
		  </tr><%
				
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(2)

		  %><tr class="normaltitle">
			<td colspan="2" align="center"><script>document.frmLogin.Password.focus()</script><br />
				<input type="hidden" name="Login" value="<%= strAsgTxtLogin %>" />
				<input type="submit" name="submit" value="<%= strAsgTxtLogin %>" />
			</td>
		  </tr>
		<% Else %>
		  <tr class="normaltitle" bgcolor="<%= strAsgSknTableTitleBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtEntryAllowed) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center" colspan="2"><br />
				<%= strAsgTxtLoginCompleted & "<br />" & strAsgTxtEntryAllowed %><br /><br />
				<%= strAsgTxtGoingToBeRedirected  %><br />
				<a href="<% If Len(Request.QueryString("backto")) > 0 Then Response.Write(Request.QueryString("backto")) Else Response.Write("statistiche.asp") %>" title="<%= strAsgTxtGoToPage %>" class="linksmalltext"><%= strAsgTxtClickToRedirect %></a><br /><br />
				<a href="login.asp?Logout=True" title="<%= strAsgTxtLogout %>" class="linksmalltext"><%= strAsgTxtClickToLogout %></a><br /><br />
			</td>
		  </tr><%
				
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(2)

		   End If %>
		</table>
		</form>
<%

' Footer
Response.Write(vbCrLf & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
'// Row - Footer Border Line
Call BuildFooterBorderLine()

' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write("<tr align=""center"" valign=""middle"">")
Response.Write("<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer""><a href=""http://www.asp-stats.com/"" class=""linkfooter"" title=""ASP Stats Generator Homepage"">ASP Stats Generator</a> [" & strAsgVersion & "] - &copy; 2003-2007 <a href=""http://www.weppos.com/"" class=""linkfooter"" title=""Weppos.com Homepage"">weppos</a>")
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