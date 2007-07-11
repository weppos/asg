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
Dim intAsgCiclo
Dim asgOutput

Dim giorno				'Riferimento per output
Dim mese				'Riferimento per output
Dim anno				'Riferimento per output
Dim mesenext			'Mese Successivo in Output
Dim annonext			'Anno Successivo in Output
Dim weekdayon			'Valore Primo giorno del mese
Dim weekdayoff			'Valore Ultimo giorno del mese calcolato sul primo del mese successivo - 1
DIm dayon				'Data Primo giorno del mese
Dim dayoff				'Data Ultimo giorno del mese calcolato sul primo del mese successivo - 1
Dim blnIsSunday			'Imposta a Vero se � Domenica

'Grafico
Dim intAsgNumCol		'Numero Colonne
Dim intAsgAltColMax		'Altezza massima in px delle colonne dipendente dalla pag
Dim intAsgLarCol		'Larghezza delle colonne dipendente dalla pag

Dim intAsgTotHits(31)		'Valore totale di pagine visitate 	| Per statistica grafica
Dim intAsgTotVisits(31)		'Valore totale di accessi unici		| Per statistica grafica


'-----------------------------------------------------------------------------------------
' Richiama Dati
'-----------------------------------------------------------------------------------------
giorno = Trim(Request("giorno"))
If giorno = "" then
	giorno = dtmAsgDay
End If


'Read setting variables from querystring
mese = Trim(Request.QueryString("mese"))

'If a time period has been chosen then build the variable to query the database
If Request.QueryString("showperiod") = strAsgTxtShow Then mese = Request.QueryString("periodmm") & "-" & Request.QueryString("periodyy")


If mese = "" Then 
	anno = dtmAsgYear
	mese = dtmAsgMonth
Else
	anno = CInt(Right(mese, 4))
	mese = CInt(Left(mese, 2))
End If



'-----------------------------------------------------------------------------------------
' Accenrtamento chiusura Anno
'-----------------------------------------------------------------------------------------
If mese = 12 then
	mesenext = 1
	annonext = anno + 1
else 
	mesenext = mese + 1
	annonext = anno
End If
	
dayon = CDate(DateSerial(Cstr(anno), Cstr(mese), "01"))
dayoff = DateAdd("d", -1, CDate(DateSerial(Cstr(anno), Cstr(mesenext), "01")))
'Response.Write(dayon) : Response.Write("<br>") : Response.Write(dayoff) : Response.Write("<br>") 
weekdayon = Weekday(Cdate("1" & "/"+Cstr(mese)+"/" & Cstr(anno)))-1
weekdayoff = Weekday(dayoff)-1
'Response.Write(weekdayon) : Response.Write("<br>") : Response.Write(weekdayoff)
	  
If weekdayoff = 0 then
	weekdayoff = 7
End If
	  
If weekdayon = 0 then
	weekdayon = 7
End If

intAsgCiclo = 1
blnIsSunday = False

'Ricalcola i giorni per mese
Call GiorniPerMese(Left(mese, 2))

'Set max total column width
intAsgNumCol = intStsGiorniPerMese	'Numero colonne | 1 per ogni giorno del mese
intAsgAltColMax = 200				'Altezza massima colonne | Rapportata alla dimensione della pagina
intAsgLarCol = (800/intAsgNumCol)	'Larghezza per ogni colonna | Calcolata sul totale delle necessarie


'Richiama valori statistica
strAsgSQL = "SELECT * FROM "&strAsgTablePrefix&"Daily WHERE Mese = '" & Right("0" & mese, 2) & "-" & anno & "' ORDER BY Data"

objAsgRs.Open strAsgSQL, objAsgConn
If objAsgRs.EOF Then
'
Else

	Do While NOT objAsgRs.EOF
	
		intAsgTotHits(Right("0" & Day(objAsgRs("Data")), 2)) = objAsgRs("Hits")
		intAsgTotVisits(Right("0" & Day(objAsgRs("Data")), 2)) = objAsgRs("Visits")
	
	objAsgRs.MoveNext
	Loop

End If
objAsgRs.Close



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
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle">CALENDARIO ACCESSI</td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" width="100%" height="16" colspan="7"><%= UCase(strAsgTxtCalendar) & "&nbsp;" & UCase(strAsgTxtStatsOfTheYear) & "&nbsp;" & mese & "-" & anno %></td>
		  </tr>
				<%
						 
				Dim ilgiorno
						  
				ilgiorno = (dayon - weekdayon + intAsgCiclo) ': Response.Write(ilgiorno)
				%>

		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<%

			Do While ilgiorno <= (dayoff + 7 - weekdayoff)
				
			%>
			<td width="14%" align="left"
			<% 
		
			If WeekDay(CDate(ilgiorno)) = 1 Then 
				'// Sfondo Rossastro
				'Response.Write("bgcolor=""#FF9966""")
				'// Sfondo Classico
				Response.Write("background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """")
				'Imposta a vero Variabile Domenica
				blnIsSunday = True
			Else 
				Response.Write("background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """")
				blnIsSunday = False
			End If

			%>><br />
			  <font color="<% If ilgiorno >= dayon AND ilgiorno <= dayoff Then
					  Response.Write("#000000")
				  Else
					  Response.Write("#CCCCCC")
				  End If%>">
				<% 
				'Mese in considerazione!
				If ilgiorno >= dayon AND ilgiorno <= dayoff Then 
					
					Response.Write("<div align=""center"">")
					'Controllo se la data � quella di oggi!
					If Day(dayon - weekdayon + intAsgCiclo) = dtmAsgDay AND Month(dayon - weekdayon + intAsgCiclo) = CInt(dtmAsgMonth) Then 
						If blnIsSunday Then Response.Write("<font color=""#FF0000"">")
						Response.Write("<font color=""#0000FF"">" & day(dayon - weekdayon + intAsgCiclo) & "</font>")
						If blnIsSunday Then Response.Write("</font>")
					Else 
						If blnIsSunday Then Response.Write("<font color=""#FF0000"">")
						Response.Write(day(dayon - weekdayon + intAsgCiclo))
						If blnIsSunday Then Response.Write("</font>")
					End If
					
					Response.Write(vbCrLf & "<a href="""" title=""" & strAsgTxtShow & """ class=""linksmalltext"">" & Left(aryAsgMonth(Month(ilgiorno), 2), 3) & "</a>")
					
					Response.Write("</div>")
					
					'Controllo ed Output Totali Hits
					Response.Write(vbCrLf & "<br /><img src=""images/bar_graph_image_hits.gif"" alt=""" & strAsgTxtHits & """ align=""absmiddle"" />&nbsp;&nbsp;" & strAsgTxtSmHits & "&nbsp;:&nbsp;")
					If IsNumeric(intAsgTotHits(Day(ilgiorno))) AND Len(intAsgTotHits(Day(ilgiorno))) > 0 Then
						Response.Write(intAsgTotHits(Day(ilgiorno)))
					Else
						Response.Write("-")
					End If
					
					'Controllo ed Output Visits
					Response.Write(vbCrLf & "<br /><img src=""images/bar_graph_image_visits.gif"" alt=""" & strAsgTxtVisits & """ align=""absmiddle"" />&nbsp;&nbsp;" & strAsgTxtSmVisits & "&nbsp;:&nbsp;")
					If IsNumeric(intAsgTotVisits(Day(ilgiorno))) AND Len(intAsgTotHits(Day(ilgiorno))) > 0 Then
						Response.Write(intAsgTotVisits(Day(ilgiorno)))
					Else
						Response.Write("-")
					End If
					
				Else

					Response.Write("<div align=""center"">")
					Response.Write(day(dayon - weekdayon + intAsgCiclo))
					Response.Write("</div>")

				End If 
				%>
			  </font>
			</td>
				<%
					If weekday(dayon - weekdayon + intAsgCiclo - 1) = 7 Then
							Response.Write(vbCrLf & "			  </tr>")
						If NOT ilgiorno >= (dayoff + 7 - weekdayoff) Then
							Response.Write(vbCrLf & "			  <tr class=""smalltext"" bgcolor=""" & strAsgSknTableContBgColour & """>")
						End If
					End If
					
					ilgiorno = ilgiorno + 1
					intAsgCiclo = intAsgCiclo + 1
					
				Loop
				
		'Ricomponi mese per funzione generale
		mese = CStr(Right("0" & mese, 2) & "-" & anno)


		'// Row - End table spacer			
		Call BuildTableContEndSpacer(7)

		'// Row - Data output panels
		Response.Write(vbCrLf & "<tr class=""smalltext"" align=""center"" valign=""top""><td colspan=""7"" height=""25""><br />")
		Call GoToPeriod("stats_calendar.asp", "") 
		Response.Write(vbCrLf & "</td></tr>")
		
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