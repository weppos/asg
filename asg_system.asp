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
Const intAsgBarMaxWidth = 100		' Max total column width
' Other variables
Dim strAsgReportOs
Dim strAsgReportBrowser

' Create a tmp querystring without new values
strAsgAppend = appendToQuerystring("sortby||sortorder")


' Get the total value to create the graph
if strAsgMode = "month" then 
	strAsgSQL = "SELECT Sum(system_hits) As SumHits, Sum(system_visits) AS SumVisits " &_
		"FROM " & ASG_TABLE_PREFIX & "system " &_
		"WHERE system_period = '" & strAsgPeriod & "' "
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT Sum(system_hits) As SumHits, Sum(system_visits) AS SumVisits " &_
		"FROM " & ASG_TABLE_PREFIX & "system "
end if
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

' Get the max item value
if strAsgMode = "month" then 
	strAsgSQL = "SELECT system_hits " &_
		"FROM " & ASG_TABLE_PREFIX & "system " &_
		"WHERE system_period = '" & strAsgPeriod & "' "
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT SUM(system_hits) AS SumHits " &_
		"FROM " & ASG_TABLE_PREFIX & "system "
end if
' Month and Year
if strAsgMode = "month" then 
		strAsgSQL = strAsgSQL & "ORDER BY system_hits DESC"
elseif strAsgMode = "all" then 
	if ASG_USE_MYSQL then
		strAsgSQL = strAsgSQL & " ORDER BY SumHits DESC"
	else
		strAsgSQL = strAsgSQL & " ORDER BY SUM(system_hits) DESC"
	end if
End If
' Set Rs properties
if ASG_USE_MYSQL then
	objAsgRs.CursorLocation = 3
end if
objAsgRs.CursorType = 1
objAsgRs.LockType = 3
' Open Rs
objAsgRs.Open strAsgSQL, objAsgConn
If objAsgRs.EOF Then
	intAsgMaxRecValue = 0
Else
	objAsgRs.MoveFirst
	if strAsgMode = "month" then 
		intAsgMaxRecValue = objAsgRs("system_hits")
	elseif strAsgMode = "all" then 
		intAsgMaxRecValue = objAsgRs("SumHits")
	End if
End If
objAsgRs.Close
' Calculate the minimal part to build the graph
if intAsgMaxRecValue = 0 OR "[]" & intAsgMaxRecValue = "[]" then intAsgMaxRecValue = 1
intAsgBarPart = intAsgBarMaxWidth / intAsgMaxRecValue


' Read sorting order from querystring
' Filter QS values and associate them 
' with their respective database fields
Select Case strAsgSortBy
	Case "hits" 	strAsgSortByFld = formatSortingField("SumHits", "SUM(system_hits)", ASG_USE_MYSQL)
	Case "visits"	strAsgSortByFld = formatSortingField("SumVisits", "SUM(system_visits)", ASG_USE_MYSQL)
	Case "browser"	strAsgSortByFld = formatSortingField("system_browser", "system_browser", ASG_USE_MYSQL)
	Case "os"		strAsgSortByFld = formatSortingField("system_os", "system_os", ASG_USE_MYSQL)
	Case "color"	strAsgSortByFld = formatSortingField("system_color", "system_color", ASG_USE_MYSQL)
	Case "reso"		strAsgSortByFld = formatSortingField("system_reso", "system_reso", ASG_USE_MYSQL)
	Case else		strAsgSortByFld = formatSortingField("SumVisits", "SUM(system_visits)", ASG_USE_MYSQL)
End Select

' Call advanced data sorting configuration and variables
Call dimAdvDataSorting

' Order message
Dim strOrderBy_message
strOrderBy_message = "&nbsp;"

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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Visitors & " &raquo; " & MENUSECTION_VisitorSystems & " &raquo;</span> " & MENUSECTION_Systems %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_Systems
Response.Write(builTableTlayout("", "open", MENUSECTION_Systems))

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
	<td width="4%"  class="treport_title">&nbsp;</td>
	<td width="22%" class="treport_title"><%= TXT_os %>
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_os) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=os&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=os&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="26%" class="treport_title"><%= TXT_browser %>
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_os) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=browser&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=browser&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="10%" class="treport_title"><%= TXT_reso_short %>
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_reso_short) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=reso&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=reso&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="8%" class="treport_title"><%= TXT_color_short %>
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_color_short) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=color&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=color&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="8%" class="treport_title"><%= TXT_traffic %></td>
	<td width="22%" class="treport_title">
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_pageviews) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=hits&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=hits&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a>
		&nbsp;&nbsp;<%= TXT_Graph %>&nbsp;&nbsp;
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_visits) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=visits&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=visits&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a>
	</td>
  </tr>
<%

' Initialise SQL string to select data
strAsgSQL = "SELECT system_browser, system_os, system_color, system_reso, SUM(system_hits) AS SumHits, SUM(system_visits) AS SumVisits " &_
	"FROM " & ASG_TABLE_PREFIX & "system "
' Month and Year		
if strAsgMode = "month" then 
	strAsgSQL = strAsgSQL & "WHERE system_period = '" & strAsgPeriod & "'  "
	strAsgSQL = searchFor(strAsgSQL, false)
elseif strAsgMode = "all" then 
	strAsgSQL = searchFor(strAsgSQL, true)
end if
' Group information by following fields
strAsgSQL = strAsgSQL & "GROUP BY system_browser, system_os, system_color, system_reso "
' Order record by selected field 
strAsgSQL = strAsgSQL & "ORDER BY " & strAsgSortByFld & " " & strAsgSortOrder & ""
		
' Set Rs properties
if ASG_USE_MYSQL then
	objAsgRs.CursorLocation = 3
end if
objAsgRs.CursorType = 1
objAsgRs.LockType = 3
	
' Open Rs
objAsgRs.Open strAsgSQL, objAsgConn
			
	' The recordset is empty
	if objAsgRs.EOF Then
				
		' If it is a search query then show no results advise
		if Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 then

			' No current record for search terms		
			Response.Write(buildTableContNoRecord(7, "search"))
					
		' Else show general no record information
		Else
	
			' No current record			
			Response.Write(buildTableContNoRecord(7, "standard"))
					
		End If
				
	Else
			
		objAsgRs.PageSize = RecordsPerPage
		objAsgRs.AbsolutePage = page
				
		for loopAdvDataSorting = 1 to RecordsPerPage
					
			if not objAsgRs.EOF then			
				strAsgReportBrowser = objAsgRs("system_browser")
				strAsgReportOs = objAsgRs("system_os")
%>		  
  <tr <%= buildTableContRollover("treport_row") %> >
	<td class="treport_col">&nbsp;</td>
	<td class="treport_col" align="left"><img src="<%= STR_ASG_SKIN_PATH_IMAGE %>def/os.asp?icon=<%= strAsgReportOs %>" alt="<%= strAsgReportOs %>" align="middle" /> <%= searchTerms(strAsgReportOs, "system_os", asgSearchfor, asgSearchIn) %></td>
	<td class="treport_col" align="left"><img src="<%= STR_ASG_SKIN_PATH_IMAGE %>def/browser.asp?icon=<%= strAsgReportBrowser %>" alt="<%= strAsgReportBrowser %>" align="middle" /> <%= searchTerms(strAsgReportBrowser, "system_browser", asgSearchfor, asgSearchIn) %></td>
	<td class="treport_col" align="center"><%= searchTerms(objAsgRs("system_reso"), "system_reso", asgSearchfor, asgSearchIn) %></td>
	<td class="treport_col" align="center"><%= searchTerms(objAsgRs("system_color"), "system_color", asgSearchfor, asgSearchIn) & "&nbsp;" & TXT_Bit %></td>
	<td class="treport_col" align="right"><%= objAsgRs("SumHits") & "<br />" & objAsgRs("SumVisits") %></td>
	<td class="treport_col" style="text-align: left;">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_hits_h.gif" width="<%= FormatNumber(objAsgRs("SumHits") * intAsgBarPart, 2) %>" height="9" alt="<%= TXT_pageviews %>" align="middle" /> [<%= calcPercValue(intAsgTotMonthHits, objAsgRs("SumHits")) %>]<br />
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_visits_h.gif" width="<%= FormatNumber(objAsgRs("SumVisits") * intAsgBarPart, 2) %>" height="9" alt="<%= TXT_visits %>" align="middle" /> [<%= calcPercValue(intAsgTotMonthVisits, objAsgRs("SumVisits")) %>]
	</td>
  </tr>
<%
				
		objAsgRs.MoveNext
		end if
	next
	end if

%></table><%

' Advanced data sorting
strLayerAdvDataSorting = buildLayerAdvDataSorting()

objAsgRs.Close

' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing
		
	' Report Legend
	Response.Write(buildLayerReportLegend())
				
' :: Close tlayout :: MENUSECTION_Browsers
Response.Write(builTableTlayout("", "close", ""))


Response.Write(vbCrLf & "<br />")


' :: Open tlayout :: BARLABEL_DataView
Response.Write(builTableTlayout("rowNavy", "open", buildSwapDisplay("rowNavy", BARLABEL_DataView)))

	' Advanced data sorting layer
	Response.Write(strLayerAdvDataSorting)
			
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
'Response.Write(builTableTlayout("x-rowExport", "open", buildSwapDisplay("rowExport", BARLABEL_DataExport)))

	' Row - Layers search
	' Response.Write(buildLayerSearch("", "Browser"))
'	Response.Write("&nbsp;")
				
' :: Close tlayout :: BARLABEL_DataExport
'Response.Write(builTableTlayout("", "close", ""))


'Response.Write(vbCrLf & "<br />")


' :: Open tlayout :: BARLABEL_DataSearch
Response.Write(builTableTlayout("x-rowSearch", "open", buildSwapDisplay("rowSearch", BARLABEL_DataSearch)))

	' Row - Layers search
	Response.Write(buildLayerSearch("", "system_browser|system_os|system_color|system_reso"))
				
' :: Close tlayout :: BARLABEL_DataSearch
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