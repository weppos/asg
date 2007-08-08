<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'// WARNING! Program protection.
'	Changing default values may allow users to access the page.
Call checkPermission("True", "False", "False", appAsgSecurity)


'Dimension variables
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
Dim intAsgColNum		'Numero Colonne
Dim ASG_COL_MAXHEIGHT		'Altezza massima in px delle colonne dipendente dalla pag
Dim intAsgColWidth		'Larghezza delle colonne dipendente dalla pag

Dim intAsgTotHits(31)		'Valore totale di pagine visitate 	| Per statistica grafica
Dim intAsgTotVisits(31)		'Valore totale di accessi unici		| Per statistica grafica


'-----------------------------------------------------------------------------------------
' Collect period information
'-----------------------------------------------------------------------------------------
'Check the year value
anno = Trim(Request.QueryString("periody"))
If IsNumeric(anno) AND Len(anno) > 0 then
	anno = CInt(anno)
Else
	anno = dtmAsgYear
End if
'Check the month value
mese = Trim(Request.QueryString("periodm"))
If IsNumeric(mese) AND Len(mese) > 0 then
	mese = CInt(mese)
Else
	mese = dtmAsgMonth
End if


'-----------------------------------------------------------------------------------------
' Accertamento chiusura Anno
'-----------------------------------------------------------------------------------------
If mese = 12 then
	mesenext = 1
	annonext = anno + 1
else 
	mesenext = mese + 1
	annonext = anno
End If
	
dayon = CDate(DateSerial(anno, mese, "01"))
dayoff = DateAdd("d", -1, CDate(DateSerial(anno, mesenext, "01")))
'Response.Write(dayon) : Response.Write("<br>") : Response.Write(dayoff) : Response.Write("<br>") 
weekdayon = Weekday(Cdate(DateSerial(anno, mese, "01")))-1
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
intAsgColNum = intStsGiorniPerMese	'Numero colonne | 1 per ogni giorno del mese
ASG_COL_MAXHEIGHT = 200				'Altezza massima colonne | Rapportata alla dimensione della pagina
intAsgColWidth = (800/intAsgColNum)	'Larghezza per ogni colonna | Calcolata sul totale delle necessarie


'Richiama valori statistica
strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "daily WHERE daily_period = '" & Right("0" & mese, 2) & "-" & anno & "' ORDER BY daily_date"

objAsgRs.Open strAsgSQL, objAsgConn
If objAsgRs.EOF Then
'
Else

	Do While NOT objAsgRs.EOF
	
		intAsgTotHits(Right("0" & Day(objAsgRs("Data")), 2)) = objAsgRs("daily_hits")
		intAsgTotVisits(Right("0" & Day(objAsgRs("Data")), 2)) = objAsgRs("daily_visits")
	
	objAsgRs.MoveNext
	Loop

End If
objAsgRs.Close



' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing


%>
<%= STR_ASG_PAGE_DOCTYPE %>
<html>
<head>
<title><%= appAsgSiteName %> | powered by ASP Stats Generator v<%= ASG_VERSION %></title>
<%= STR_ASG_PAGE_CHARSET %>
<meta name="copyright" content="Copyright (C) 2003-2005 Carletti Simone" />
<!--#include file="includes/meta.inc.asp" -->

<!-- ASP Stats Generator v. <%= ASG_VERSION %> is created and developed by Simone Carletti.
To download your Free copy visit the official site http://www.weppos.com/asg/ -->

</head>

<!--#include file="includes/header.asp" -->
<%
			  
' TableBar			
Call buildTableBar(MENUSECTION_MonthlyCalendar, MENUGROUP_VisitorProfiles)
	
' 
Response.Write(vbCrLf & "<div class=""table_layout"">")
		  
%>
<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr bgcolor="<%= STR_ASG_SKIN_TABLE_TITLE_BGCOLOUR %>" align="center" class="normaltitle">
	<td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_TITLE_BGIMAGE %>" width="100%" height="16" colspan="7"><%= UCase(TXT_Calendar) & "&nbsp;" & UCase(TXT_StatsOfTheYear) & "&nbsp;" & mese & "-" & anno %></td>
  </tr>
<%
				 
Dim ilgiorno
ilgiorno = (dayon - weekdayon + intAsgCiclo) ': Response.Write(ilgiorno)

%>
  <tr>
<%

	Do While ilgiorno <= (dayoff + 7 - weekdayoff)
		
	%>
	<td width="14%" align="left" <% 
		
	If WeekDay(CDate(ilgiorno)) = 1 Then 
		' Red bgcolor
		'Response.Write("bgcolor=""#FF9966""")
		' Classic bgcolor
		buildTableContRollover("table_cont_row")
		' Sunday boolean value
		blnIsSunday = True
	Else 
		buildTableContRollover("table_cont_row")
		blnIsSunday = False
	End If

	%>><p align="center"><% 
	
		'Mese in considerazione!
		If ilgiorno >= dayon AND ilgiorno <= dayoff Then 
				
			'Controllo se la data è quella di oggi!
			If Day(dayon - weekdayon + intAsgCiclo) = dtmAsgDay AND Month(dayon - weekdayon + intAsgCiclo) = CInt(dtmAsgMonth) Then 
				Response.Write("<font color=""#0000FF"">" & day(dayon - weekdayon + intAsgCiclo) & "</font>")
			Else 
				If blnIsSunday Then Response.Write("<font color=""#FF0000"">")
				Response.Write(day(dayon - weekdayon + intAsgCiclo))
				If blnIsSunday Then Response.Write("</font>")
			End If
			Response.Write(vbCrLf & "<a href="""" title=""" & MENUSECTION_MonthlyReports & "&nbsp;(" & aryAsgMonth(1, Month(ilgiorno)) & ")" & """>" & Left(aryAsgMonth(1, Month(ilgiorno)), 3) & "</a><br />")
					
			'Controllo ed Output Totali Hits
			Response.Write(vbCrLf & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "bar_graph_image_hits.gif"" alt=""" & TXT_pageviews & """ align=""absmiddle"" />&nbsp;")
			If IsNumeric(intAsgTotHits(Day(ilgiorno))) AND Len(intAsgTotHits(Day(ilgiorno))) > 0 Then
				Response.Write(intAsgTotHits(Day(ilgiorno)))
			Else
				Response.Write("-")
			End If
					
			'Controllo ed Output Visits
			Response.Write(vbCrLf & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "bar_graph_image_visits.gif"" alt=""" & TXT_visits & """ align=""absmiddle"" />&nbsp;")
			If IsNumeric(intAsgTotVisits(Day(ilgiorno))) AND Len(intAsgTotHits(Day(ilgiorno))) > 0 Then
				Response.Write(intAsgTotVisits(Day(ilgiorno)))
			Else
				Response.Write("-")
			End If
					
		Else
			Response.Write(day(dayon - weekdayon + intAsgCiclo))
		End If 
		%>
	  </p>
	</td>
	<%

	If weekday(dayon - weekdayon + intAsgCiclo - 1) = 7 Then
		Response.Write(vbCrLf & "			  </tr>")
		If NOT ilgiorno >= (dayoff + 7 - weekdayoff) Then
			Response.Write(vbCrLf & "<tr bgcolor=""" & STR_ASG_SKIN_TABLE_CONT_BGCOLOUR & """>")
		End If
	End If
					
	ilgiorno = ilgiorno + 1
	intAsgCiclo = intAsgCiclo + 1
					
	Loop
				
'Ricomponi mese per funzione generale
mese = CStr(Right("0" & mese, 2) & "-" & anno)

'// Row - End table spacer			
Call buildTableContEndSpacer(7)
				
'// Row - Legend			
Call buildTableContLegend(7)
				
'// Row - End table spacer			
Call buildTableContEndSpacer(7)

'// Row - Data output panels
Response.Write(vbCrLf & "<tr align=""center"" valign=""top""><td colspan=""7"" height=""25""><br />")
Call labelShowPeriod("stats_monthly_calendar.asp", "") 
Response.Write(vbCrLf & "</td></tr>")
		
		%>
</table>
<%

' 
Response.Write(vbCrLf & "</div>")

' Footer
' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write(vbCrLf & "<div class=""table_footerbar"">Powered by <a href=""http://www.weppos.com/asg/"" title=""ASP Stats Generator"">ASP Stats Generator</a> v" & ASG_VERSION & " &copy; 2003-2004 <a href=""http://www.weppos.com/"">weppos</a>")
If ASG_ELABORATION_TIME Then Response.Write(" - " & TXT_ThisPageWasGeneratedIn & "&nbsp;" & FormatNumber(Timer() - startAsgElab, 4) & "&nbsp;" & TXT_seconds)
Response.Write("</div>")
Response.Write(vbCrLf & "<br /><div class=""footer"" align=""center"">Powered by <a href=""http://www.weppos.com/asg/"" title=""ASP Stats Generator"">ASP Stats Generator</a> v" & ASG_VERSION & " <br />Copyright &copy; 2003-2005 <a href=""http://www.weppos.com/"">weppos</a><div>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

%>
<!--#include file="includes/footer.asp" -->
</body></html>