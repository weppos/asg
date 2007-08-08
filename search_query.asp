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
Dim intAsgBarMaxWidth		' Max total column width

' Other variables
Dim strAsgDetails
strAsgDetails = Request.QueryString("details")
Dim strAsgGroup
strAsgGroup = formatSetting("group", "engine")
Dim intNumColspan
Dim lngAsgReportHits
Dim lngAsgReportVisits

'Set max total column width
If strAsgGroup = "engine" then 
	intAsgBarMaxWidth = 200				'Largezza massima colonne | Rapportata alla dimensione della pagina
elseif strAsgGroup = "query" then 
	intAsgBarMaxWidth = 200				'Largezza massima colonne | Rapportata alla dimensione della pagina
end if

' Create a tmp querystring without new values
strAsgAppend = appendToQuerystring("sortby||sortorder||details")


' Get the total value to create the graph
if strAsgMode = "month" then 
	strAsgSQL = "SELECT Sum(query_hits) As SumHits, Sum(query_visits) AS SumVisits " &_
		"FROM " & ASG_TABLE_PREFIX & "query " &_
		"WHERE query_period = '" & strAsgPeriod & "' "
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT Sum(query_hits) As SumHits, Sum(query_visits) AS SumVisits " &_
		"FROM " & ASG_TABLE_PREFIX & "query "
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
if intAsgTotMonthVisits = 0 OR "[]" & intAsgTotMonthVisits = "[]" then intAsgTotMonthVisits = 0

'Get the max value
If strAsgMode = "month" AND strAsgGroup = "engine" then 
	strAsgSQL = "SELECT MAX(query_hits) AS SumHits FROM " & ASG_TABLE_PREFIX & "query WHERE query_period = '" & strAsgPeriod & "' "
elseif strAsgMode = "all" AND strAsgGroup = "engine" then 
	strAsgSQL = "SELECT MAX(query_hits) AS SumHits FROM " & ASG_TABLE_PREFIX & "query "
elseif strAsgMode = "month" AND strAsgGroup = "query" then 
	strAsgSQL = "SELECT SUM(query_hits) AS SumHits FROM " & ASG_TABLE_PREFIX & "query WHERE query_period = '" & strAsgPeriod & "' "
elseif strAsgMode = "all" AND strAsgGroup = "query" then 
	strAsgSQL = "SELECT SUM(query_hits) AS SumHits FROM " & ASG_TABLE_PREFIX & "query "
end if
'Change grouping mode depending on database
if strAsgGroup = "query" then
	if ASG_USE_MYSQL then
		strAsgSQL = strAsgSQL & " GROUP BY query_keyphrase ORDER BY SumHits DESC"
	else
		strAsgSQL = strAsgSQL & " GROUP BY query_keyphrase ORDER BY SUM(query_hits) DESC"
	end if
elseif strAsgGroup = "engine" then
	if ASG_USE_MYSQL then
		strAsgSQL = strAsgSQL & " ORDER BY SumHits DESC"
	else
		strAsgSQL = strAsgSQL & " ORDER BY SUM(query_hits) DESC"
	end if
end if
' Set Rs properties
if ASG_USE_MYSQL then
	objAsgRs.CursorLocation = 3
end if
objAsgRs.CursorType = 1
objAsgRs.LockType = 3
' Open Rs
objAsgRs.Open strAsgSQL, objAsgConn
if objAsgRs.EOF then
	intAsgMaxRecValue = 0
else
	objAsgRs.MoveFirst
	intAsgMaxRecValue = Clng(objAsgRs("SumHits"))
end if
objAsgRs.Close
' Calculate the minimal part to build the graph
if intAsgMaxRecValue = 0 OR "[]" & intAsgMaxRecValue = "[]" then intAsgMaxRecValue = 1
intAsgBarPart = intAsgBarMaxWidth / intAsgMaxRecValue


' Read sorting order from querystring
' Filter QS values and associate them 
' with their respective database fields
if strAsgGroup = "engine" then

	select case strAsgSortBy
		Case "hits" 	strAsgSortByFld = formatSortingField("query_hits", "query_hits", ASG_USE_MYSQL)
		Case "visits"	strAsgSortByFld = formatSortingField("query_visits", "query_visits", ASG_USE_MYSQL)
		Case "engine"	strAsgSortByFld = formatSortingField("engine_name", "engine_name", ASG_USE_MYSQL)
		Case "query"	strAsgSortByFld = formatSortingField("query_keyphrase", "query_keyphrase", ASG_USE_MYSQL)
		Case else	 	strAsgSortByFld = formatSortingField("query_hits", "query_hits", ASG_USE_MYSQL)
'		Case else		strAsgSortByFld = formatSortingField("query_visits", "query_visits", ASG_USE_MYSQL)
	end select
	intNumColspan = 5
	
elseif strAsgGroup = "query" then
	
	select case strAsgSortBy
		Case "hits" 	strAsgSortByFld = formatSortingField("SumHits", "SUM(query_hits)", ASG_USE_MYSQL)
		Case "visits"	strAsgSortByFld = formatSortingField("SumVisits", "SUM(query_visits)", ASG_USE_MYSQL)
		Case "query"	strAsgSortByFld = formatSortingField("query_keyphrase", "query_keyphrase", ASG_USE_MYSQL)
		Case else		strAsgSortByFld = formatSortingField("SumHits", "SUM(query_hits)", ASG_USE_MYSQL)
'		Case else		strAsgSortByFld = formatSortingField("SumVisits", "SUM(query_visits)", ASG_USE_MYSQL)
	end select
	intNumColspan = 4
	
end if

' Call advanced data sorting configuration and variables
Call dimAdvDataSorting

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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Marketing & " &raquo;</span> " & MENUSECTION_SearchQueries %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_SearchQueries
Response.Write(builTableTlayout("", "open", MENUSECTION_SearchEngines))

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
	<td width="5%"  class="treport_title">&nbsp;</td>
	<% If strAsgGroup = "engine" then %>
	<td width="30%" class="treport_title"><%= TXT_Query %>
		<a href="?<%= strAsgAppend & "&amp;sortby=query&amp;sortorder=DESC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_Query & "&nbsp;" & TXT_orderDesc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_Query & "&nbsp;" & TXT_orderDesc %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=query&amp;sortorder=ASC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_Query & "&nbsp;" & TXT_orderAsc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_Query & "&nbsp;" & TXT_orderAsc %>" /></a></td>
	<td width="20%" class="treport_title"><%= TXT_search_engine %>
		<a href="?<%= strAsgAppend & "&amp;sortby=engine&amp;sortorder=DESC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_search_engine & "&nbsp;" & TXT_orderDesc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_search_engine & "&nbsp;" & TXT_orderDesc %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=engine&amp;sortorder=ASC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_search_engine & "&nbsp;" & TXT_orderAsc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_search_engine & "&nbsp;" & TXT_orderAsc %>" /></a></td>
	<% elseif strAsgGroup = "query" then %>
	<td width="50%" class="treport_title"><%= TXT_Query %>
		<a href="?<%= strAsgAppend & "&amp;sortby=query&amp;sortorder=DESC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_Query & "&nbsp;" & TXT_orderDesc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_Query & "&nbsp;" & TXT_orderDesc %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=query&amp;sortorder=ASC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_Query & "&nbsp;" & TXT_orderAsc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_Query & "&nbsp;" & TXT_orderAsc %>" /></a></td>
	<% end if %>
	<td width="12%" class="treport_title"><%= TXT_traffic %></td>
	<td width="33%" class="treport_title">
		<a href="?<%= strAsgAppend & "&amp;sortby=hits&amp;sortorder=DESC&amp;details=" & strAsgDetails %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=hits&amp;sortorder=ASC&amp;details=" & strAsgDetails %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a>
		&nbsp;&nbsp;<%= TXT_Graph %>&nbsp;&nbsp;
		<a href="?<%= strAsgAppend & "&amp;sortby=visits&amp;sortorder=DESC&amp;details=" & strAsgDetails %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=visits&amp;sortorder=ASC&amp;details=" & strAsgDetails %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
  </tr>
<%

' Initialise SQL string to select data
if strAsgMode = "month" AND strAsgGroup = "engine" then 
	strAsgSQL = "SELECT query_keyphrase, engine_name, query_serp_page, query_hits, query_visits " &_
		"FROM " & ASG_TABLE_PREFIX & "query " &_
		"WHERE query_period = '" & strAsgPeriod & "' "
	strAsgSQL = searchFor(strAsgSQL, false)		' Search
elseif strAsgMode = "all" AND strAsgGroup = "engine" then 
	strAsgSQL = "SELECT query_keyphrase, engine_name, query_serp_page, query_hits, query_visits " &_
		"FROM " & ASG_TABLE_PREFIX & "query "
	strAsgSQL = searchFor(strAsgSQL, true)		' Search
elseif strAsgMode = "month" AND strAsgGroup = "query" then 
	strAsgSQL = "SELECT query_keyphrase, AVG(query_serp_page) AS AvgSERP, SUM(query_hits) AS SumHits, SUM(query_visits) AS SumVisits " &_
		"FROM " & ASG_TABLE_PREFIX & "query " &_
		"WHERE query_period = '" & strAsgPeriod & "' "
	strAsgSQL = searchFor(strAsgSQL, false)		' Search
	strAsgSQL = strAsgSQL & " GROUP BY query_keyphrase "
elseif strAsgMode = "all" AND strAsgGroup = "query" then 
	strAsgSQL = "SELECT query_keyphrase, AVG(query_serp_page) AS AvgSERP, SUM(query_hits) AS SumHits, SUM(query_visits) AS SumVisits " &_
		"FROM " & ASG_TABLE_PREFIX & "query "
	strAsgSQL = searchFor(strAsgSQL, true)		' Search
	strAsgSQL = strAsgSQL & " GROUP BY query_keyphrase "
end if
' Order record by selected field 
strAsgSQL = strAsgSQL & " ORDER BY " & strAsgSortByFld & " " & strAsgSortOrder & ""

' Set Rs properties
if ASG_USE_MYSQL then
	objAsgRs.CursorLocation = 3
end if
objAsgRs.CursorType = 1
objAsgRs.LockType = 3

' Open Rs
objAsgRs.Open strAsgSQL, objAsgConn
			
	' The recordset is empty
	if objAsgRs.EOF then
				
		' If it is a search query then show no results advise
		if Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 then

			' No current record for search terms		
			Response.Write(buildTableContNoRecord(intNumColspan, "search"))
					
		' Else show general no record information
		Else
	
			' No current record			
			Response.Write(buildTableContNoRecord(intNumColspan, "standard"))
					
		end if
				
	Else
			
		objAsgRs.PageSize = RecordsPerPage
		objAsgRs.AbsolutePage = page
				
		for loopAdvDataSorting = 1 to RecordsPerPage
					
			if not objAsgRs.EOF then

				if strAsgGroup = "engine" then
					lngAsgReportHits = Clng(objAsgRs("query_hits"))
					lngAsgReportVisits = Clng(objAsgRs("query_visits"))
				elseif strAsgGroup = "query" then
					lngAsgReportHits = Clng(objAsgRs("SumHits"))
					lngAsgReportVisits = Clng(objAsgRs("SumVisits"))
				end if
					
%>		  
  <tr <%= buildTableContRollover("treport_row") %> >
	<td class="treport_col" style="text-align: center;"></td>
	<% if strAsgGroup = "engine" then %>
	<td class="treport_col" style="text-align: left;"><% if Cint(objAsgRs("query_serp_page")) > 0 then Response.Write("&nbsp;<a href=""serp.asp?serp=" & objAsgRs("query_serp_page") & "&amp;periodm=" & intAsgPeriodM & "&amp;periody=" & intAsgPeriodY & """ title=""" & TXT_Query & "&nbsp;" & objAsgRs("query_serp_page") & "&deg;&nbsp;" & TXT_page & """><span class=""notetext"">[" & objAsgRs("query_serp_page") & "]</span></a>") %>&nbsp;<%= ShareWords(searchTerms(objAsgRs("query_keyphrase"), "query_keyphrase"), 40) %></td>
	<td class="treport_col" style="text-align: left;">&nbsp;<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>def/engine.asp?icon=<%= objAsgRs("engine_name") %>" alt="<%= objAsgRs("engine_name") %>" align="middle" /> <%= searchTerms(objAsgRs("engine_name"), "engine_name") %></td>
	<% elseif strAsgGroup = "query" then %>
	<td class="treport_col" style="text-align: left;"><% if Cint(objAsgRs("AvgSERP")) > 0 then Response.Write("&nbsp;<span class=""notetext"">[" & objAsgRs("AvgSERP") & "]</span>") %>&nbsp;<%= ShareWords(searchTerms(objAsgRs("query_keyphrase"), "query_keyphrase"), 40) %></td>
	<% end if %>
	<td class="treport_col" style="text-align: right;"><%= lngAsgReportHits & "<br />" & lngAsgReportVisits %></td>
	<td class="treport_col" style="text-align: left;">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_hits_h.gif" width="<%= FormatNumber(lngAsgReportHits * intAsgBarPart, 2) %>" height="9" alt="<%= TXT_pageviews %>" align="middle" /> [<%= calcPercValue(intAsgTotMonthHits, lngAsgReportHits) %>]<br />
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_visits_h.gif" width="<%= FormatNumber(lngAsgReportVisits * intAsgBarPart, 2) %>" height="9" alt="<%= TXT_visits %>" align="middle" /> [<%= calcPercValue(intAsgTotMonthVisits,lngAsgReportVisits) %>]
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
				
' :: Close tlayout :: MENUSECTION_SearchQueries
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
			
	' Group condition layer
	Response.Write(buildLayerGroup("query|engine", "query|engine"))

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


Response.Write(vbCrLf & "<br />")


' :: Open tlayout :: BARLABEL_DataSearch
Response.Write(builTableTlayout("x-rowSearch", "open", buildSwapDisplay("rowSearch", BARLABEL_DataSearch)))

	' Row - Layers search
	if strAsgGroup = "engine" then
		Response.Write(buildLayerSearch("", "query_keyphrase|engine_name"))
	elseif strAsgGroup = "query" then
		Response.Write(buildLayerSearch("", "query_keyphrase"))
	end if
				
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