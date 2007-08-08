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
Const ASG_COL_LEGEND_WITH = 2
Dim intAsgColNum			' Holds the column number
Dim intAsgColWidth		' Holds the max column width in px
Dim intAsgDays

' Other variables
Dim ii
Dim dtmAsgValData(31)		' Date values
Dim intAsgValHits(31)		' Graph values
Dim intAsgValVisits(31)		' Graph values
Dim intAsgTotHits(31)		' Visited pages
Dim intAsgTotVisits(31)		' Unique visitors

' Calculate the number of day per month
intAsgDays = getDayCount(intAsgPeriodM)

' Set max total column width
intAsgColNum = intAsgDays + ASG_COL_LEGEND_WITH	' Column number: 1 per each hour
intAsgColWidth = (600 / intAsgColNum)	' Column width: depending on column number


' Get the total value to create the graph
strAsgSQL = "SELECT Sum(daily_hits) As SumHits, Sum(daily_visits) AS SumVisits " &_
	"FROM " & ASG_TABLE_PREFIX & "daily " &_
	"WHERE daily_period = '" & strAsgPeriod & "' "
' Open Rs
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
strAsgSQL = "SELECT Max(daily_hits) As MaxHits " &_
	"FROM " & ASG_TABLE_PREFIX & "daily " &_
	"WHERE daily_period = '" & strAsgPeriod & "' "
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
strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "daily " &_
	"WHERE daily_period = '" & strAsgPeriod & "' " &_
	"ORDER BY daily_date"
' Open Rs
objAsgRs.Open strAsgSQL, objAsgConn
if not objAsgRs.EOF then

	Do While NOT objAsgRs.EOF
	
		intAsgTotHits(Right("0" & Day(objAsgRs("daily_date")), 2)) = objAsgRs("daily_hits")
		intAsgTotVisits(Right("0" & Day(objAsgRs("daily_date")), 2)) = objAsgRs("daily_visits")
	
	objAsgRs.MoveNext
	Loop

end if
objAsgRs.Close

' Filter values
for ii = 1 to (intAsgColNum - ASG_COL_LEGEND_WITH)
	
	' Filter null values
	if "[]" & intAsgTotHits(ii) = "[]" then intAsgTotHits(ii) = 0
	if "[]" & intAsgTotVisits(ii) = "[]" then intAsgTotVisits(ii) = 0
	' Format values
	intAsgValHits(ii) = FormatNumber(intAsgTotHits(ii) * intAsgBarPart, 2)
	intAsgValVisits(ii) = FormatNumber(intAsgTotVisits(ii) * intAsgBarPart, 2)
	' Get the data
	dtmAsgValData(ii) = CDate(DateSerial(intAsgPeriodY, intAsgPeriodM, ii))
	
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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Reports & " &raquo;</span> " & MENUSECTION_DailyReports %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_DailyReports
Response.Write(builTableTlayout("", "open", MENUSECTION_DailyReports))

%>
<div class="treport_col_grapcont">
<div class="treport_title"><%= TXT_StatsOfTheMonth & "&nbsp;" & strAsgPeriod %></div>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr class="treport_row" style="text-align: center;">
	<td class="treport_col_graphval" width="<%= intAsgColWidth * ASG_COL_LEGEND_WITH %>" colspan="2" nowrap="nowrap">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<% for ii = 1 to 5 %>
			<tr style="height: <%= ASG_COL_MAXHEIGHT / 5 %>px; text-align: right;"><td width="100%" valign="top"><%= CLng((intAsgMaxRecValue / 5) * (6 - ii)) %></td></tr>
			<% next %>
		</table>
	</td>
	<% 
		' Create js comments
		Dim strAsgJsComment
		strAsgJsComment = strAsgJsComment & vbCrLf & "Comment[0]=["",""]"

		Dim strAsgPercVisits
		Dim strAsgPercHits
		
		' Loop all months
		for ii = 1 to (intAsgColNum - ASG_COL_LEGEND_WITH)

		' Get percentual values
		strAsgPercVisits = calcPercValue(intAsgTotMonthVisits, intAsgTotVisits(ii))
		strAsgPercHits = calcPercValue(intAsgTotMonthHits, intAsgTotHits(ii))
		
		' Write js comments
		strAsgJsComment = strAsgJsComment & vbCrLf & "Comment[" & (ii * 2) & "]=[""<img src='" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_chart.png' alt='" & TXT_Graph & "' border='0' align='middle' />&nbsp;&nbsp;" & TXT_Graph & "&nbsp;(" & strAsgPercVisits & ")"",""<strong>" & TXT_visits & "</strong>:&nbsp;" & intAsgTotVisits(ii)& """]"
		strAsgJsComment = strAsgJsComment & vbCrLf & "Comment[" & (ii * 2) + 1 & "]=[""<img src='" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_chart.png' alt='" & TXT_Graph & "' border='0' align='middle' />&nbsp;&nbsp;" & TXT_Graph & "&nbsp;(" & strAsgPercHits & ")"",""<strong>" & TXT_pageviews & "</strong>:&nbsp;" & intAsgTotHits(ii)& """]"

		%>
	<td class="treport_col_graphcell" width="<%= intAsgColWidth %>" nowrap="nowrap">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_visits.gif" width="5" height="<%= intAsgValVisits(ii) %>" alt="<%= strAsgPercVisits %>" <%= "onmouseover=""stm(Comment[" & (ii * 2) & "],Style[4])"" onmouseout=""htm()""" %> />
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_hits.gif" width="5" height="<%= intAsgValHits(ii) %>" alt="<%= strAsgPercHits %>" <%= "onmouseover=""stm(Comment[" & (ii * 2) + 1 & "],Style[5])"" onmouseout=""htm()""" %> />
	</td>
	<% 
		next
	%>
  </tr>
  <tr class="treport_row" style="text-align: center;">
	<td class="treport_col" width="<%= intAsgColWidth * 2 %>" colspan="2" align="right">
		<%= Left(aryAsgMonth(1, intAsgPeriodM), 3) %>
		<a href="stats_monthly.asp?periody=<%= intAsgPeriodY %>&amp;showsubmit=<%= TXT_button_show %>" title="<%= MENUSECTION_MonthlyReports & "&nbsp;(" & intAsgPeriodY & ")" %>"><%= intAsgPeriodY %></a>
	</td>
	<% 
		for ii = 1 to (intAsgColNum - ASG_COL_LEGEND_WITH)
				
			' Sunday
			Response.Write(VbCrLf & "<td width=""" & intAsgColWidth & """>")
			if Weekday(dtmAsgValData(ii)) = 1 then
					Response.Write("<font class=""notetext"" title=""" & TXT_Sunday & """>" & ii & "</font>")
			else
					Response.Write(ii)
			end If
			Response.Write("</td>")
					
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
				
' :: Close tlayout :: MENUSECTION_DailyReports
Response.Write(builTableTlayout("", "close", ""))


Response.Write(vbCrLf & "<br />")


' :: Open tlayout :: BARLABEL_DataView
Response.Write(builTableTlayout("rowNavy", "open", buildSwapDisplay("rowNavy", BARLABEL_DataView)))
			
	' Open the Navy form
	Response.Write(buildLayerForm("open"))
			
	' Period selection layer
	Response.Write(buildLayerPeriod())

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