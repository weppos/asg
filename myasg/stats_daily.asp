<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="config.asp" -->
<%

' 
' = ASP Stats Generator - Powerful and reliable ASP website counter
' 
' Copyright (c) 2003-2008 Simone Carletti <weppos@weppos.net>
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
' 
' 
' @category        ASP Stats Generator
' @package         ASP Stats Generator
' @author          Simone Carletti <weppos@weppos.net>
' @copyright       2003-2008 Simone Carletti
' @license         http://www.opensource.org/licenses/mit-license.php
' @version         SVN: $Id$
' 


'// ATTENZIONE! Protezione statistiche.
'	Modificare solo se necessario e se sicuri.
'	Impostazioni errate possono compromettere la privacy.
Call AllowEntry("True", "False", "False", intAsgSecurity)


'Dichiara Variabili
Dim mese				'Riferimento per output
Dim giorno				'Riferimento per output
Dim intAsgCiclo


'Grafico
Dim intAsgNumCol		'Numero Colonne
Dim intAsgAltColMax		'Altezza massima in px delle colonne dipendente dalla pag
Dim intAsgLarCol		'Larghezza delle colonne dipendente dalla pag

Dim intAsgMax				'Valore massimo di pagine visitate
Dim intAsgParte

Dim intAsgTotMeseHits		'Valore totale per mese di pagine visitate
Dim intAsgTotMeseVisits		'Valore totale per mese di accessi unici

Dim intAsgValHits(31)		'Valori assunti per l'immagine
Dim intAsgValVisits(31)		'Valori assunti per l'immagine

Dim intAsgTotHits(31)		'Valore totale di pagine visitate 	| Per statistica grafica
Dim intAsgTotVisits(31)		'Valore totale di accessi unici		| Per statistica grafica

'Variabili Calendario
Dim dtmAsgValData(31)		'Valori data progressivi
Dim tmpAsgMonth				'Variabile d'elaborazione
Dim tmpAsgYear				'Variabile d'elaborazione


'Read setting variables from querystring
mese = Request.QueryString("mese")
'If a time period has been chosen then build the variable to query the database
If Request.QueryString("showperiod") = strAsgTxtShow Then mese = Request.QueryString("periodmm") & "-" & Request.QueryString("periodyy")


'If period variable is empty then set it to the current month
If mese = "" Then mese = dtmAsgMonth & "-" & dtmAsgYear


'Ricalcola i giorni per mese
Call GiorniPerMese(Left(mese, 2))

'Set max total column width
intAsgNumCol = intStsGiorniPerMese + 2	'Numero colonne | 1 per ogni giorno del mese
intAsgAltColMax = 200					'Altezza massima colonne | Rapportata alla dimensione della pagina
intAsgLarCol = (800/intAsgNumCol)		'Larghezza per ogni colonna | Calcolata sul totale delle necessarie


'Richiama totale
strAsgSQL = "SELECT Sum(Hits) As SumHits, Sum(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Daily WHERE Mese = '" & mese & "' "
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
strAsgSQL = "SELECT Max(Hits) As MaxHits FROM "&strAsgTablePrefix&"Daily WHERE Mese = '" & mese & "' "
objAsgRs.Open strAsgSQL, objAsgConn
If objAsgRs.EOF Then
	intAsgMax = 0
Else
	intAsgMax = objAsgRs("MaxHits")
End If
objAsgRs.Close


'Calcola unit� singola
If intAsgMax = 0 OR "[]" & intAsgMax = "[]" Then intAsgMax = 1
intAsgParte = intAsgAltColMax/intAsgMax


'Richiama valori statistica
strAsgSQL = "SELECT * FROM "&strAsgTablePrefix&"Daily WHERE Mese = '" & mese & "' ORDER BY Data"

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


'Prima del ciclo definisci alcune variabili temporanee
'da utilizzare per le variabili calendario
tmpAsgMonth = CInt(Left(mese, 2))
tmpAsgYear = Cint(Right(Mese, 4))
'Ripassiamo tutto per filtrare i valori nulli o vuoti
'...contemporaneamente impostiamo i valori
For intAsgCiclo = 1 to  (intAsgNumCol - 2)
	
	If "[]" & intAsgTotHits(intAsgCiclo) = "[]" Then intAsgTotHits(intAsgCiclo) = 0
	If "[]" & intAsgTotVisits(intAsgCiclo) = "[]" Then intAsgTotVisits(intAsgCiclo) = 0
	
	intAsgValHits(intAsgCiclo) = FormatNumber(intAsgTotHits(intAsgCiclo)*intAsgParte, 2)
	intAsgValVisits(intAsgCiclo) = FormatNumber(intAsgTotVisits(intAsgCiclo)*intAsgParte, 2)
	
	dtmAsgValData(intAsgCiclo) = CDate(DateSerial(tmpAsgYear, tmpAsgMonth, intAsgCiclo))
	
Next


'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | powered by ASP Stats Generator <%= ASG_VERSION %></title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="copyright" content="Copyright (C) 2003-2008 Carletti Simone, All Rights Reserved" />
<meta name="generator" content="ASP Stats Generator <%= ASG_VERSION %>" /> <!-- leave this for stats -->

<!--#include file="asg-includes/layout/head.asp" -->

<!--
  ASP Stats Generator (release <%= ASG_VERSION %>) is a free software package
  completely written in ASP programming language, for real time visitor tracking.
  Get your own copy for free at http://www.asp-stats-com/ !
-->

</head>

<!--#include file="asg-includes/layout/header.asp" -->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle">ACCESSI per GIORNO</td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" width="100%" height="16"><%= UCase(strAsgTxtStatsOfTheMonth) & "&nbsp;" & mese %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="100%"><br />
			<!-- Grafico -->
			<table width="100%" border="0" cellspacing="1" cellpadding="0" align="center">
			  <tr valign="bottom" align="center">
				<td width="<%= intAsgLarCol * 2 %>" background="<%= strAsgSknPathImage %>layout/bg_graph_value.gif" colspan="2" nowrap>
					<table border="0" cellpadding="0" cellspacing="0" width="100%">
					<% For intAsgCiclo = 1 to 5 %>
					<tr height="<%= intAsgAltColMax/5 %>" align="right"><td width="100%" class="smalltext" valign="top"><%= CLng(intAsgMax / 5) * (6 - intAsgCiclo) %></td></tr>
					<% Next %>
					</table>
				</td>
				<% For intAsgCiclo = 1 to (intAsgNumCol - 2) %>
				<td width="<%= intAsgLarCol %>" background="<%= strAsgSknPathImage %>layout/bg_graph_cell.gif" nowrap>
					<img src="images/bar_graph_image_visits.gif" width="5" height="<%= intAsgValVisits(intAsgCiclo) %>" alt="<%= CalcolaPercentuale(intAsgTotMeseVisits, intAsgTotVisits(intAsgCiclo)) & " --> " & intAsgTotVisits(intAsgCiclo) & "&nbsp;" & strAsgTxtVisits %>" />
					<img src="images/bar_graph_image_hits.gif" width="5" height="<%= intAsgValHits(intAsgCiclo) %>" alt="<%= CalcolaPercentuale(intAsgTotMeseHits, intAsgTotHits(intAsgCiclo)) & " --> " & intAsgTotHits(intAsgCiclo) & "&nbsp;" & strAsgTxtHits %>" />
				</td>
			  	<% Next %>
			  </tr>
			  <tr class="smalltext" align="center">
				<td width="<%= intAsgLarCol * 2 %>" colspan="2" align="right"><a href="stats_monthly.asp?showperiod=<%= strAsgTxtShow %>&anno=<%= Right(mese , 4) %>" title="<%= strAsgTxtShow & "&nbsp;" & asgMonthName(Left(mese, 2)) %>" class="linksmalltext"><%= Left(asgMonthName(Left(mese, 2)), 3) %></a></td>
				<% 
					For intAsgCiclo = 1 to (intAsgNumCol - 2)
						
						'Evidenzia le Domeniche
						If Weekday(dtmAsgValData(intAsgCiclo)) = 1 Then
							Response.Write(VbCrLf & "<td width=""" & intAsgLarCol & """>")
							Response.Write("<font color=""#FF0000"" title=""" & strAsgTxtSunday & """>" & intAsgCiclo & "</font>")
						Else
							Response.Write(VbCrLf & "<td width=""" & intAsgLarCol & """>")
							Response.Write(intAsgCiclo)
						End If
						
						Response.Write("</td>")
						
			  		Next
				%>
			  </tr>
<!--#include file="templates/_graph_legend.asp" -->
			</table>
			<!-- Fine Grafico -->
			<br />
			</td>
		  </tr>
		<%
					
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(intAsgNumCol)

		'// Row - Data output panels
		Response.Write(vbCrLf & "<tr class=""smalltext"" align=""center"" valign=""top""><td colspan=""" & intAsgNumCol & """ height=""25""><br />")
		Call GoToPeriod("stats_daily.asp", "")
		Response.Write(vbCrLf & "</td></tr>")
			
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
Response.Write("<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer""><a href=""http://www.asp-stats.com/"" class=""linkfooter"" title=""ASP Stats Generator Homepage"">ASP Stats Generator</a> " & ASG_VERSION & " - &copy; 2003-2008 <a href=""http://www.weppos.com/"" class=""linkfooter"" title=""Weppos.com Homepage"">weppos</a>")
if ASG_CONFIG_ELABTIME then Response.Write(" - " & asgElabtime())
Response.Write("</td>")
Response.Write("</tr>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write("</table>")

%>
<!--#include file="asg-includes/layout/footer.asp" -->

</body></html>
