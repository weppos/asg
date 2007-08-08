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
Const ASG_COL_LEGEND_WITH = 1
Dim intAsgColNum			' Holds the column number
Dim intAsgColWidth		' Holds the max column width in px

' Other variables
Dim ii
Dim dtmAsgValData()		' Date values
Dim intAsgValHits()		' Graph values
Dim intAsgValVisits()	' Graph values
Dim intAsgTotHits()		' Visited pages
Dim intAsgTotVisits()	' Unique visitors

' Count the number of records
strAsgSQL = "SELECT COUNT(counter_id) " &_
	"FROM " & ASG_TABLE_PREFIX & "counter "
objAsgRs.Open strAsgSQL, objAsgConn
	intAsgColNum = Cint(objAsgRs(0))
objAsgRs.Close

' Set max total column width
' intAsgColNum = Year(Date()) - Year(appAsgProgramSetup)	' Column number: from the year of the program setup
intAsgColWidth = (600 / (intAsgColNum + ASG_COL_LEGEND_WITH))	' Column width: depending on column number

Redim dtmAsgValData(intAsgColNum)
Redim intAsgValHits(intAsgColNum)
Redim intAsgValVisits(intAsgColNum)
Redim intAsgTotHits(intAsgColNum)
Redim intAsgTotVisits(intAsgColNum)

' Get the total value to create the graph
strAsgSQL = "SELECT Sum(counter_hits) As SumHits, Sum(counter_visits) AS SumVisits " &_
	"FROM " & ASG_TABLE_PREFIX & "counter "
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

' Get data from database
intAsgMaxRecValue = 1
ii = 0
strAsgSQL = "SELECT counter_hits, counter_visits, counter_periody " &_
	"FROM " & ASG_TABLE_PREFIX & "counter " &_
	"ORDER BY counter_periody"
' Open Rs
objAsgRs.Open strAsgSQL, objAsgConn
if not objAsgRs.EOF then

	Do While NOT objAsgRs.EOF

		ii = ii + 1
		dtmAsgValData(ii) = Cint(objAsgRs("counter_periody"))
		intAsgTotHits(ii) = Clng(objAsgRs("counter_hits"))
		intAsgTotVisits(ii) = Clng(objAsgRs("counter_visits"))
		
		if objAsgRs("counter_hits") > intAsgMaxRecValue then intAsgMaxRecValue = Clng(objAsgRs("counter_hits"))
	
	objAsgRs.MoveNext
	Loop

end If
objAsgRs.Close 

' Calculate the minimal part to build the graph
intAsgBarPart = ASG_COL_MAXHEIGHT / intAsgMaxRecValue 

' Filter values
for ii = 1 to (intAsgColNum)
	
	' Filter null values
	if "[]" & intAsgTotHits(ii) = "[]" then intAsgTotHits(ii) = 0
	if "[]" & intAsgTotVisits(ii) = "[]" then intAsgTotVisits(ii) = 0
	' Format values
	intAsgValHits(ii) = FormatNumber(intAsgTotHits(ii) * intAsgBarPart, 2)
	intAsgValVisits(ii) = FormatNumber(intAsgTotVisits(ii) * intAsgBarPart, 2)
	
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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Reports & " &raquo;</span> " & MENUSECTION_YearlyReports %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_MonthlyReports
Response.Write(builTableTlayout("", "open", MENUSECTION_MonthlyReports))

%>
<div class="treport_col_grapcont">
<div class="treport_title"><%= "ASP STATS GENERATOR" %></div>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr class="treport_row" style="text-align: center;">
	<td class="treport_col_graphval" width="<%= intAsgColWidth * ASG_COL_LEGEND_WITH %>" nowrap="nowrap">
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
		
		' Loop all years
		for ii = 1 to (intAsgColNum) 

		' Get percentual values
		strAsgPercVisits = calcPercValue(intAsgTotMonthVisits, intAsgTotVisits(ii))
		strAsgPercHits = calcPercValue(intAsgTotMonthHits, intAsgTotHits(ii))
		
		' Write js comments
		strAsgJsComment = strAsgJsComment & vbCrLf & "Comment[" & (ii * 2) & "]=[""<img src='" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_chart.png' alt='" & TXT_Graph & "' border='0' align='middle' />&nbsp;&nbsp;" & TXT_Graph & "&nbsp;(" & strAsgPercVisits & ")"",""<strong>" & TXT_visits & "</strong>:&nbsp;" & intAsgTotVisits(ii) & """]"
		strAsgJsComment = strAsgJsComment & vbCrLf & "Comment[" & (ii * 2) + 1 & "]=[""<img src='" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_chart.png' alt='" & TXT_Graph & "' border='0' align='middle' />&nbsp;&nbsp;" & TXT_Graph & "&nbsp;(" & strAsgPercHits & ")"",""<strong>" & TXT_pageviews & "</strong>:&nbsp;" & intAsgTotHits(ii) & """]"

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
	<td class="treport_col" width="<%= intAsgColWidth %>" align="right">&nbsp;</td>
	<% for ii = 1 to (intAsgColNum) %>
	<td width="<%= intAsgColWidth %>"><a href="stats_monthly.asp?periody=<%= dtmAsgValData(ii) %>&amp;showsubmit=<%= TXT_button_show %>" title="<%= MENUSECTION_MonthlyReports & "&nbsp;(" & dtmAsgValData(ii) & ")" %>"><%= dtmAsgValData(ii) %></a></td>
  	<% next %>
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
	Response.Write(buildLayerPeriodY())

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