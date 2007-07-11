<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Option Explicit %>
<!--#include file="config.asp" -->
<!--#include file="includes/inc_array_table.asp" -->
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


Dim aryAsgTableWarining(10,2)		'Holds the array containing warning informations


'Set warning record level, joining table id to the related included array id
aryAsgTableWarining(1,1) = 5000
aryAsgTableWarining(2,1) = 2000
aryAsgTableWarining(3,1) = 750
aryAsgTableWarining(4,1) = 2000
aryAsgTableWarining(5,1) = 500
aryAsgTableWarining(6,1) = 3500
aryAsgTableWarining(7,1) = 3500
aryAsgTableWarining(8,1) = 2000
aryAsgTableWarining(9,1) = 750
aryAsgTableWarining(10,1) = 3500


'Execute a loop to count the records of each table
For intAsgTableLoop = 1 to Ubound(aryAsgTable)
	
	'Initialise SQL string to count records
	strAsgSQL = "SELECT COUNT(*) FROM "&strAsgTablePrefix& aryAsgTable(intAsgTableLoop,1) & ""
	'Open Rs
	objAsgRs.Open strAsgSQL, objAsgConn
	'Set the number of total hits
	If Not objAsgRs.EOF Then 
		aryAsgTableWarining(intAsgTableLoop,2) = objAsgRs(0)
	Else
		aryAsgTableWarining(intAsgTableLoop,2) = 0
	End If
	'Close Rs
	objAsgRs.Close
	
	'Set warning alert
	If aryAsgTableWarining(intAsgTableLoop,2) > aryAsgTableWarining(intAsgTableLoop,1) Then
		aryAsgTableWarining(intAsgTableLoop,0) = True
	Else
		aryAsgTableWarining(intAsgTableLoop,0) = False
	End If

Next

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
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtResetSettings) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="70%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <% If Request.QueryString("msg") = "error" Then %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" align="center" height="15"><br />
			<strong><%= strAsgTxtErrorOccured %><br />
			<%= strAsgTxtCheckTableMatching %></strong><br /><br />
			</td>		  
		  </tr>
		  <% End If %>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtTableReset) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableContBgColour %>" class="smalltext" align="center">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" height="30">
			<img src="<%= strAsgSknPathImage %>warning_icon.gif" border="0" alt="<%= strAsgTxtAdvice %>" align="absmiddle">&nbsp;<%= strAsgTxtTablesWithWarningIconNeedsReset %></td>
		  </tr>
		<%
		
		For intAsgTableLoop = 0 to Ubound(aryAsgTable)
		
		%>
		  <form action="settings_reset_execute.asp" method="get" name="frmReset">
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="100%" colspan="2" height="20">&nbsp;
				<input type="hidden" name="table" value="<%= intAsgTableLoop %>" />
				<%
					
					'Show an alert icon if the table need a cleaning
					If aryAsgTableWarining(intAsgTableLoop,0) Then
					Response.Write(vbCrLf & "<img src=""" & strAsgSknPathImage & "warning_icon.gif"" border=""0"" align=""absmiddle"">")
					End If
					
					'Write table title and description
					Response.Write(vbCrLf & "<span class=""notetext"">" & aryAsgTable(intAsgTableLoop, 1) & "</span>&nbsp;(" & aryAsgTableWarining(intAsgTableLoop,2) & "&nbsp;" & strAsgTxtRecords & ")&nbsp;-&nbsp;" & aryAsgTable(intAsgTableLoop, 2))
				
				%>
			</td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="60%" height="18" align="right">&nbsp;
				<%= strAsgTxtDataReset %>
				<select name="timerange" class="smallform">
					<option value="full"><%= strAsgTxtFull %></option>
				<% 
				'Nei casi riportati di sotto mostra una select
				'limitata a causa delle impostazioni limitate delle strtture
				If aryAsgTable(intAsgTableLoop, 1) <> "IP" Then %>
					<option value="0"><%= strAsgTxtOlderThan & " " & strAsgTxtCurrent & " " & strAsgTxtMonth %></option>
					<option value="1"><%= strAsgTxtOlderThan & " 1 " & strAsgTxtMonth %></option>
					<% For looptmp = 2 to 12 %>
					<option value="<%= looptmp %>"><%= strAsgTxtOlderThan & " " & looptmp & " " & strAsgTxtMonths %></option>
					<% Next 
				ElseIf aryAsgTable(intAsgTableLoop, 1) = "Detail" Then %>
					<option value="0"><%= strAsgTxtOlderThan & " " & strAsgTxtCurrent & " " & strAsgTxtWeek %></option>
					<option value="1"><%= strAsgTxtOlderThan & " 1 " & strAsgTxtWeek %></option>
					<% For looptmp = 2 to 12 %>
					<option value="<%= looptmp %>"><%= strAsgTxtOlderThan & " " & looptmp & " " & strAsgTxtWeeks %></option>
					<% Next 
				End If
				%>
				</select>&nbsp;&nbsp;
			</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="40%">&nbsp;
				<input type="image" src="images/delete.gif" name="delete" value="deletenormal" /></td>
		  </tr>
		  <% If aryAsgTable(intAsgTableLoop, 1) = "Detail" Then %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="60%" height="18" align="right">&nbsp;
				<%= strAsgTxtDataReset %>
				<select name="weekrange" class="smallform">
					<option value=""></option>
					<option value="0"><%= strAsgTxtOlderThan & " " & strAsgTxtCurrent & " " & strAsgTxtWeek %></option>
					<option value="1"><%= strAsgTxtOlderThan & " 1 " & strAsgTxtWeek %></option>
					<% For looptmp = 2 to 12 %>
					<option value="<%= looptmp %>"><%= strAsgTxtOlderThan & " " & looptmp & " " & strAsgTxtWeeks %></option>
					<% Next %>
				</select>&nbsp;&nbsp;
			</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="40%">&nbsp;
				<input type="image" src="images/delete.gif" name="delete" value="deleteweek" /></td>
		  </tr>
		  <% End If 'Condizione tabella details	%>
		  </form>
		<%
		
		Next  
		
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(2)

		%>
		</table><br />
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