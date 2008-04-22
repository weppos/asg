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
Dim blnIsSunday			'Imposta a Vero se è Domenica

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
<title><%= strAsgSiteName %> | powered by ASP Stats Generator <%= strAsgVersion %></title>
<meta name="copyright" content="Copyright (C) 2003-2008 Carletti Simone, All Rights Reserved" />
<meta name="generator" content="ASP Stats Generator <%= strAsgVersion %>" /> <!-- leave this for stats -->

<!--#include file="asg-includes/layout/head.asp" -->

<!--
  ASP Stats Generator (release <%= strAsgVersion %>) is a free software package
  completely written in ASP programming language, for real time visitor tracking.
  Get your own copy for free at http://www.asp-stats-com/ !
-->

</head>

<!--#include file="asg-includes/layout/header.asp" -->
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
					'Controllo se la data è quella di oggi!
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

' Footer
Response.Write(vbCrLf & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
'// Row - Footer Border Line
Call BuildFooterBorderLine()

' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write("<tr align=""center"" valign=""middle"">")
Response.Write("<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer""><a href=""http://www.asp-stats.com/"" class=""linkfooter"" title=""ASP Stats Generator Homepage"">ASP Stats Generator</a> [" & strAsgVersion & "] - &copy; 2003-2008 <a href=""http://www.weppos.com/"" class=""linkfooter"" title=""Weppos.com Homepage"">weppos</a>")
if ASG_CONFIG_ELABTIME then Response.Write(asgElabtime())
Response.Write("</td>")
Response.Write("</tr>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write("</table>")

%>
<!--#include file="asg-includes/layout/footer.asp" -->

</body></html>
