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

' Include commons variable, declarations 
' and data filtering.
%><!--#include file="includes/variables.inc.asp" --><%

' Graph
' Const intAsgBarMaxWidth = 300		' Max total column width
Dim intAsgColNum			' Holds the column number
Dim intAsgColWidth		' Holds the max column width in px

' Other variables
Dim i
' Dim dtmAsgValData(31)		' Valori data progressivi
Dim intAsgValHits(24)		' Valori assunti per l'immagine
Dim intAsgValVisits(24)		' Valori assunti per l'immagine
Dim intAsgTotHits(24)		' Valore totale di pagine visitate 	| Per statistica grafica
Dim intAsgTotVisits(24)		' Valore totale di accessi unici		| Per statistica grafica

'Set max total column width
intAsgColNum = 24 + 2	'Numero colonne | 1 per ogni ora
intAsgColWidth = (600 / intAsgColNum)	'Larghezza per ogni colonna | Calcolata sul totale delle necessarie


' Get the total value to create the graph
if strAsgMode = "month" then 
	strAsgSQL = "SELECT Sum(hourly_hits) As SumHits, Sum(hourly_visits) AS SumVisits " &_
		"FROM " & ASG_TABLE_PREFIX & "hourly WHERE hourly_period = '" & strAsgPeriod & "' "
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT Sum(hourly_hits) As SumHits, Sum(hourly_visits) AS SumVisits " &_
		"FROM " & ASG_TABLE_PREFIX & "hourly "
End If
objAsgRs.Open strAsgSQL, objAsgConn
if objAsgRs.EOF then
	intAsgTotMonthHits = 0
	intAsgTotMonthVisits = 0
else
	intAsgTotMonthHits = objAsgRs("SumHits")
	intAsgTotMonthVisits = objAsgRs("SumVisits")
end if
objAsgRs.Close
' Filter null values
if intAsgTotMonthHits = 0 OR "[]" & intAsgTotMonthHits = "[]" then intAsgTotMonthHits = 0
if intAsgTotMonthVisits = 0 OR "[]" & intAsgTotMonthVisits = "[]" then intAsgTotMonthVisits = 0

' Get the max item value
if strAsgMode = "month" then 
	strAsgSQL = "SELECT Max(hourly_hits) As MaxHits " &_
		"FROM " & ASG_TABLE_PREFIX & "hourly " &_
		"WHERE hourly_period = '" & strAsgPeriod & "' "
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT Sum(hourly_hits) AS MaxHits " &_
		"FROM " & ASG_TABLE_PREFIX & "hourly " &_
		"GROUP BY hourly_hour "
	'Change sorting mode depending on database
	if ASG_USE_MYSQL then
		strAsgSQL = strAsgSQL & " ORDER BY MaxHits DESC"
	else
		strAsgSQL = strAsgSQL & " ORDER BY Sum(hourly_hits) DESC"
	end if
end if
' Open Rs
objAsgRs.Open strAsgSQL, objAsgConn
If objAsgRs.EOF Then
	intAsgMaxRecValue = 0
Else
	intAsgMaxRecValue = objAsgRs("MaxHits")
End If
objAsgRs.Close

' Calculate the minimal part to build the graph
if intAsgMaxRecValue = 0 OR "[]" & intAsgMaxRecValue = "[]" then intAsgMaxRecValue = 1
intAsgBarPart = ASG_COL_MAXHEIGHT / intAsgMaxRecValue

' Get data from database
if strAsgMode = "month" then 
	strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "hourly " &_
		"WHERE hourly_period = '" & strAsgPeriod & "' " &_
		"ORDER BY hourly_hour "
	' Open Rs
	objAsgRs.Open strAsgSQL, objAsgConn
	if not objAsgRs.EOF then
		Do While NOT objAsgRs.EOF
		
			intAsgTotHits(objAsgRs("hourly_hour")) = objAsgRs("hourly_hits")
			intAsgTotVisits(objAsgRs("hourly_hour")) = objAsgRs("hourly_visits")
		
		objAsgRs.MoveNext
		Loop
	end If
	objAsgRs.Close
	
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT Sum(hourly_hits) As SumHits, Sum(hourly_visits) As SumVisits, hourly_hour " &_
		"FROM " & ASG_TABLE_PREFIX & "hourly " &_
		"GROUP BY hourly_hour " &_
		"ORDER BY hourly_hour "
	' Open Rs
	objAsgRs.Open strAsgSQL, objAsgConn
	if not objAsgRs.EOF then
		Do While NOT objAsgRs.EOF
		
			intAsgTotHits(objAsgRs("hourly_hour")) = objAsgRs("SumHits")
			intAsgTotVisits(objAsgRs("hourly_hour")) = objAsgRs("SumVisits")
		
		objAsgRs.MoveNext
		Loop
	end If
	objAsgRs.Close
	
end If

' Filter values
for i = 0 to (intAsgColNum - 2) -1
	
	' Filter null values
	if "[]" & intAsgTotHits(i) = "[]" then intAsgTotHits(i) = 0
	if "[]" & intAsgTotVisits(i) = "[]" then intAsgTotVisits(i) = 0
	' Format values
	intAsgValHits(i) = FormatNumber(intAsgTotHits(i) * intAsgBarPart, 2)
	intAsgValVisits(i) = FormatNumber(intAsgTotVisits(i) * intAsgBarPart, 2)
	
next

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

<body>
<!--#include file="includes/header.asp" -->

<div align="center">
	<div id="layout">
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Reports & " &raquo;</span> " & MENUSECTION_HourlyReports %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_HourlyReports
Response.Write(builTableTlayout("", "open", MENUSECTION_HourlyReports))

%>
<div class="treport_col_grapcont"><%
if strAsgMode = "month" then 
%><div class="treport_title"><%= TXT_StatsOfTheMonth & "&nbsp;" & strAsgPeriod %></div><%
elseif strAsgMode = "all" then 
%><div class="treport_title"><%= TXT_StatsOfTheYear & "&nbsp;" & intAsgPeriodY %></div><%
end if 
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr class="treport_row" style="text-align: center;">
	<td class="treport_col_graphval" width="<%= intAsgColWidth * 2 %>" colspan="2" nowrap="nowrap">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<% For i = 1 to 5 %>
			<tr style="height: <%= ASG_COL_MAXHEIGHT / 5 %>px; text-align: right;"><td width="100%" valign="top"><%= CLng((intAsgMaxRecValue / 5) * (6 - i)) %></td></tr>
			<% Next %>
		</table>
	</td>
	<% 
		' Create js comments
		Dim strAsgJsComment

		Dim strAsgPercVisits
		Dim strAsgPercHits
		
		' Loop all months
		for i = 0 to (intAsgColNum - 2) -1

		' Get percentual values
		strAsgPercVisits = calcPercValue(intAsgTotMonthVisits, intAsgTotVisits(i))
		strAsgPercHits = calcPercValue(intAsgTotMonthHits, intAsgTotHits(i))
		
		' Write js comments
		strAsgJsComment = strAsgJsComment & vbCrLf & "Comment[" & (i * 2) & "]=[""<img src='" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_chart.png' alt='" & TXT_Graph & "' border='0' align='middle' />&nbsp;&nbsp;" & TXT_Graph & "&nbsp;(" & strAsgPercVisits & ")"",""<strong>" & TXT_visits & "</strong>:&nbsp;" & intAsgTotVisits(i)& """]"
		strAsgJsComment = strAsgJsComment & vbCrLf & "Comment[" & (i * 2) + 1 & "]=[""<img src='" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_chart.png' alt='" & TXT_Graph & "' border='0' align='middle' />&nbsp;&nbsp;" & TXT_Graph & "&nbsp;(" & strAsgPercHits & ")"",""<strong>" & TXT_pageviews & "</strong>:&nbsp;" & intAsgTotHits(i)& """]"

		%>
	<td class="treport_col_graphcell" width="<%= intAsgColWidth %>" nowrap="nowrap">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_visits.gif" width="5" height="<%= intAsgValVisits(i) %>" alt="<%= strAsgPercVisits %>" <%= "onmouseover=""stm(Comment[" & (i * 2) & "],Style[4])"" onmouseout=""htm()""" %> />
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_hits.gif" width="5" height="<%= intAsgValHits(i) %>" alt="<%= strAsgPercHits %>" <%= "onmouseover=""stm(Comment[" & (i * 2) + 1 & "],Style[5])"" onmouseout=""htm()""" %> />
	</td>
	<% 
		next
	%>
  </tr>
  <tr class="treport_row" style="text-align: center;">
	<td class="treport_col" width="<%= intAsgColWidth * 2 %>" colspan="2" align="right">
		<a href="stats_daily.asp?periodm=<%= intAsgPeriodM %>&amp;periody=<%= intAsgPeriodY %>&amp;showsubmit=<%= TXT_button_show %>" title="<%= MENUSECTION_DailyReports & "&nbsp;(" & strAsgPeriod & ")" %>"><%= Left(aryAsgMonth(1, intAsgPeriodM), 3) %></a>
		<a href="stats_monthly.asp?periody=<%= intAsgPeriodY %>&amp;showsubmit=<%= TXT_button_show %>" title="<%= MENUSECTION_MonthlyReports & "&nbsp;(" & intAsgPeriodY & ")" %>"><%= intAsgPeriodY %></a>
	</td>
	<% 
		for i = 0 to (intAsgColNum - 2) -1
				
			' Sunday
			Response.Write(VbCrLf & "<td width=""" & intAsgColWidth & """>" & i & "</td>")
					
  		next
	%>
  </tr>
</table>
<%
		
		' Print js comments 
		strAsgJsComment = "<script language=""JavaScript"" type=""text/javascript""><!--" & strAsgJsComment & "//--></script>"
		Response.Write(strAsgJsComment)

%></div><%
		
	' Report Legend
	Response.Write(buildLayerReportLegend())
				
' :: Close tlayout :: MENUSECTION_HourlyReports
Response.Write(builTableTlayout("", "close", ""))


Response.Write(vbCrLf & "<br />")


' :: Open tlayout :: BARLABEL_DataView
Response.Write(builTableTlayout("rowNavy", "open", buildSwapDisplay("rowNavy", BARLABEL_DataView)))
			
	' Open the Navy form
	Response.Write(buildLayerForm("open"))
			
	' Period selection layer
	Response.Write(buildLayerPeriod())
			
	' Period selection layer
	Response.Write(buildLayerMode())

	' Close the Navy form
	Response.Write(buildLayerForm("close"))
				
' :: Close tlayout :: BARLABEL_DataView
Response.Write(builTableTlayout("", "close", ""))


Response.Write(vbCrLf & "<br />")


' :: Open tlayout :: BARLABEL_DataExport
Response.Write(builTableTlayout("x-rowExport", "open", buildSwapDisplay("rowExport", BARLABEL_DataExport)))

	' Row - Layers search
	' Response.Write(buildLayerSearch("", "Browser"))
	Response.Write("&nbsp;")
				
' :: Close tlayout :: BARLABEL_DataExport
Response.Write(builTableTlayout("", "close", ""))

%>

		</div>
	</div>
</div>

<br /></div>
<!-- / body -->
<%

' Footer
Response.Write(vbCrLf & "<div id=""footer"">")
' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write(vbCrLf & "<br /><div style=""text-align: center;"">Powered by <a href=""http://www.weppos.com/asg/"" title=""ASP Stats Generator"">ASP Stats Generator</a> v" & ASG_VERSION & " ") 
if ASG_BUILDINFO then Response.Write("build " & ASG_VERSION_BUILD)
Response.Write(vbCrLf & "<br />Copyright &copy; 2003-2005 <a href=""http://www.weppos.com/"">weppos</a></div>")
if ASG_ELABORATION_TIME then Response.Write("<div class=""elabtime"">" & Replace(TXT_elabtime, "$time$", FormatNumber(Timer() - startAsgElab, 4)) & "</div>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******
Response.Write(vbCrLf & "</div>")

%>
<!--#include file="includes/footer.asp" -->
</body></html>