<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Option Explicit %>
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
'	Questo � un programma gratuito; potete modificare ed adattare il codice		'
'	(a vostro rischio) in qualsiasi sua parte nei termini delle condizioni		'
'	della licenza che lo accompagna.											'
'																				'
'	Non � consentito utilizzare l'applicazione per conseguire ricavi 			'
'	personali, distribuirla, venderla o diffonderla come una propria 			'
'	creazione anche se modificata nel codice, senza un esplicito e scritto 		'
'	consenso dell'autore.														'
'																				'
'	Potete modificare il codice sorgente (a vostro rischio) per adattarlo 		'
'	alle vostre esigenze o integrarlo nel sito; nel caso le funzioni possano	'
'	essere di utilit� pubblica vi invitiamo a comunicarlo per poterle 			'
'	implementare in una futura versione e per contribuire allo sviluppo 		'
'	del programma.																'
'																				'
'	In nessun caso l'autore sar� responsabile di danni causati da una 			'
'	modifica, da un uso non corretto o da un uso qualsiasi 						'
'	dell'applicazione.															'
'																				'
'	Nell'utilizzo devono rimanere intatte tutte le informazioni sul 			'
'	copyright; � possibile modificare o rimuovere unicamente le indicazioni 	'
'	espressamente specificate.													'
'																				'
'	Numerose ore sono state impiegate nello sviluppo del progetto e, anche 		'
'	se non vincolante ai fini dell'uso, sarebbe gratificante l'inserimento		'
'	di un link all'applicazione sul vostro sito.								'
'																				'
'	NESSUNA GARANZIA															'
'	------------------------- 													'
'	Questo programma � distribuito nella speranza che possa essere utile ma 	'
'	senza GARANZIA DI ALCUN GENERE.												'
'	L'utente si assume tutte le responsabilit� nell'uso.						'
'																				'
'-------------------------------------------------------------------------------'

'********************************************************************************'
'*																				*'	
'*	VIOLAZIONE DELLA LICENZA													*'
'*	 																			*'
'*	L'utilizzo dell'applicazione violando le condizioni di licenza comporta la 	*'
'*	perdita immediata della possibilit� d'uso ed � PERSEGUIBILE LEGALMENTE!		*'
'*																				*'
'********************************************************************************'


'// ATTENZIONE! Protezione statistiche.
'	Modificare solo se necessario e se sicuri.
'	Impostazioni errate possono compromettere la privacy.
Call AllowEntry("True", "False", "False", intAsgSecurity)


'Dichiara Variabili
Dim anno				'Riferimento per output
Dim giorno				'Riferimento per output
Dim intAsgCiclo

Dim intAsgNumCol		'Numero Colonne
Dim intAsgAltColMax		'Altezza massima in px delle colonne dipendente dalla pag
Dim intAsgLarCol		'Larghezza delle colonne dipendente dalla pag

Dim intAsgMax				'Valore massimo di pagine visitate

Dim intAsgTotMeseHits		'Valore totale per mese di pagine visitate
Dim intAsgTotMeseVisits		'Valore totale per mese di accessi unici

Dim intAsgParte

Dim intAsgValHits(12)		'Valori assunti per l'immagine
Dim intAsgValVisits(12)		'Valori assunti per l'immagine

Dim intAsgTotHits(12)		'Valore totale di pagine visitate 	| Per statistica grafica
Dim intAsgTotVisits(12)		'Valore totale di accessi unici		| Per statistica grafica


'Read setting variables from querystring
anno = Request.QueryString("mese")


'If period variable is empty then set it to the current year
If anno = "" Then anno = FormatOutTimeZone(dtmAsgNow, "Year")
'If a time period has been chosen then build the variable to query the database
If Request.QueryString("showperiod") = strAsgTxtShow Then anno = Request.QueryString("anno")


'Set max total column width
intAsgNumCol = 12 + 1				'Numero colonne | 1 per ogni mese
intAsgAltColMax = 200				'Altezza massima colonne | Rapportata alla dimensione della pagina
intAsgLarCol = (800/intAsgNumCol)	'Larghezza per ogni colonna | Calcolata sul totale delle necessarie


'Richiama totale
strAsgSQL = "SELECT Sum(Hits) As SumHits, Sum(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Daily WHERE Mese Like '%" & anno & "' "
objAsgRs.Open strAsgSQL, objAsgConn
If objAsgRs.EOF Then
	intAsgTotMeseHits = 0
	intAsgTotMeseVisits = 0
Else
	intAsgTotMeseHits = objAsgRs("SumHits")
	intAsgTotMeseVisits = objAsgRs("SumVisits")
End If
objAsgRs.Close
'Accertati che non siano nulli
If intAsgTotMeseHits = 0 OR "[]" & intAsgTotMeseHits = "[]" Then intAsgTotMeseHits = 0
If intAsgTotMeseVisits = 0 OR "[]" & intAsgTotMeseVisits = "[]" Then intAsgTotMeseVisits = 0


'Richiama valore Massimo
strAsgSQL = "SELECT Sum(Hits) As SumHits FROM "&strAsgTablePrefix&"Daily WHERE Mese Like '%" & anno & "' GROUP BY Mese ORDER BY Sum(Hits) DESC "
objAsgRs.Open strAsgSQL, objAsgConn, 2, 3
If objAsgRs.EOF Then
	intAsgMax = 0
Else
	objAsgRs.MoveFirst
	intAsgMax = objAsgRs("SumHits")
End If
objAsgRs.Close

'Calcola unit� singola
If intAsgMax = 0 OR "[]" & intAsgMax = "[]" Then intAsgMax = 1
intAsgParte = intAsgAltColMax/intAsgMax


'Richiama valori statistica
strAsgSQL = "SELECT Sum(Hits) AS SumHits, Sum(Visits) AS SumVisits, Mese FROM "&strAsgTablePrefix&"Daily WHERE Mese Like '%" & anno & "' GROUP BY Mese ORDER BY Mese"

objAsgRs.Open strAsgSQL, objAsgConn
If objAsgRs.EOF Then
'
Else

	Do While NOT objAsgRs.EOF
	
		intAsgTotHits(Left(objAsgRs("Mese"), 2)) = objAsgRs("SumHits")
		intAsgTotVisits(Left(objAsgRs("Mese"), 2)) = objAsgRs("SumVisits")
	
	objAsgRs.MoveNext
	Loop

End If
objAsgRs.Close

'Ripassiamo tutto per filtrare i valori nulli o vuoti
'...contemporaneamente impostiamo i valori
For intAsgCiclo = 1 to (intAsgNumCol - 1)
	
	If "[]" & intAsgTotHits(intAsgCiclo) = "[]" Then intAsgTotHits(intAsgCiclo) = 0
	If "[]" & intAsgTotVisits(intAsgCiclo) = "[]" Then intAsgTotVisits(intAsgCiclo) = 0
	
	intAsgValHits(intAsgCiclo) = FormatNumber(intAsgTotHits(intAsgCiclo)*intAsgParte, 2)
	intAsgValVisits(intAsgCiclo) = FormatNumber(intAsgTotVisits(intAsgCiclo)*intAsgParte, 2)
	
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
<!--#include file="includes/inc_meta.asp" -->

<!-- 	ASP Stats Generator <%= strAsgVersion %> � una applicazione gratuita 
		per il monitoraggio degli accessi e dei visitatori ai siti web 
		creata e sviluppata da Simone Carletti.
		
		Puoi scaricarne una copia gratuita sul sito ufficiale http://www.weppos.com/ -->

</head>
<!--#include file="includes/header.asp" -->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle">ACCESSI per MESE</td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" width="100%" height="16"><%= UCase(strAsgTxtStatsOfTheYear) & "&nbsp;" & anno %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="100%"><br />
			<!-- Grafico -->
			<table width="100%" border="0" cellspacing="1" cellpadding="0" align="center">
			  <tr valign="bottom" align="center">
				<td width="<%= intAsgLarCol %>" background="<%= strAsgSknPathImage %>layout/bg_graph_value.gif" nowrap>
					<table border="0" cellpadding="0" cellspacing="0" width="100%">
					<% For intAsgCiclo = 1 to 5 %>
					<tr height="<%= intAsgAltColMax/5 %>"><td width="100%" class="smalltext" valign="top" align="right"><%= CLng(intAsgMax / intAsgCiclo) %></td></tr>
					<% Next %>
					</table>
				</td>
				<% For intAsgCiclo = 1 to (intAsgNumCol - 1) %>
				<td width="<%= intAsgLarCol %>" background="<%= strAsgSknPathImage %>layout/bg_graph_cell.gif" nowrap>
					<img src="images/bar_graph_image_visits.gif" width="5" height="<%= intAsgValVisits(intAsgCiclo) %>" alt="<%= CalcolaPercentuale(intAsgTotMeseVisits, intAsgTotVisits(intAsgCiclo)) & " --> " & intAsgTotVisits(intAsgCiclo) & "&nbsp;" & strAsgTxtVisits %>" />
					<img src="images/bar_graph_image_hits.gif" width="5" height="<%= intAsgValHits(intAsgCiclo) %>" alt="<%= CalcolaPercentuale(intAsgTotMeseHits, intAsgTotHits(intAsgCiclo)) & " --> " & intAsgTotHits(intAsgCiclo) & "&nbsp;" & strAsgTxtHits %>" />
				</td>
			  	<% Next %>
			  </tr>
			  <tr class="smalltext" align="center">
				<td width="<%= intAsgLarCol %>"></td>
				<% For intAsgCiclo = 1 to (intAsgNumCol - 1) %>
				<td width="<%= intAsgLarCol %>"><a href="stats_daily.asp?mese=<%= Right("0" & intAsgCiclo, 2) & "-" & anno %>" title="<%= strAsgTxtShow & "&nbsp;" & aryAsgMonth(intAsgCiclo,2) %>" class="linksmalltext"><%= Left(aryAsgMonth(intAsgCiclo,2), 3) %></a></td>
			  	<% Next %>
			  </tr>
<!--#include file="includes/inc_graph_legend.asp" -->
			</table>
			<!-- Fine Grafico -->
			<br />
			</td>
		  </tr>
		  <%
					
			'// Row - End table spacer			
			Call BuildTableContEndSpacer(intAsgNumCol)
	
		  %>
		  <tr class="smalltext" align="center" valign="top">
			<td width="100%" height="25"><br />
			<!-- Visualizza in base a anno -->
			<table width="200" border="0" cellspacing="0" cellpadding="0" height="30">
			<form action="stats_monthly.asp" method="get">
			  <tr valign="middle" align="center">
				<td width="25%"><%= strAsgTxtShow %></td>
				<td width="65%">
				<select name="anno" class="smallform">
					<% For looptmp = Year(dtmAsgStartStats) to dtmAsgYear %>
					<option value="<%= looptmp %>" <% If CInt(anno) = CInt(looptmp) Then Response.Write "selected" End If %>><%= looptmp %></option>
					<% Next %>
				</select>
				</td>
				<td width="10%"><input type="Submit" name="showperiod" value="<%= strAsgTxtShow %>" /></td>
			  </tr>
			</form>
			</table>
			<!-- Fine Visualizza in base a anno -->
			</td>
		  </tr>
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