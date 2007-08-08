<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<!--#include file="lib/functions_images.asp" -->
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
Const intAsgBarMaxWidth = 300		' Max total column width

' Other variables
Dim strAsgDetails
strAsgDetails = Request.QueryString("details")
Dim strAsgReportEngine

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

' Get the max item value
if strAsgMode = "month" then 
	strAsgSQL = "SELECT SUM(query_hits) AS SumHits " &_
		"FROM " & ASG_TABLE_PREFIX & "query " &_
		"WHERE query_period = '" & strAsgPeriod & "' "
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT SUM(query_hits) AS SumHits " &_
		"FROM " & ASG_TABLE_PREFIX & "query "
end if
'Change grouping mode depending on database
if ASG_USE_MYSQL then
	strAsgSQL = strAsgSQL & " GROUP BY engine_name ORDER BY SumHits DESC"
else
	strAsgSQL = strAsgSQL & " GROUP BY engine_name ORDER BY SUM(query_hits) DESC"
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
	intAsgMaxRecValue = objAsgRs("SumHits")
end If
objAsgRs.Close
' Calculate the minimal part to build the graph
if intAsgMaxRecValue = 0 OR "[]" & intAsgMaxRecValue = "[]" then intAsgMaxRecValue = 1
intAsgBarPart = intAsgBarMaxWidth / intAsgMaxRecValue


' Read sorting order from querystring
' Filter QS values and associate them 
' with their respective database fields
Select Case strAsgSortBy
	Case "hits" 	strAsgSortByFld = formatSortingField("SumHits", "SUM(query_hits)", ASG_USE_MYSQL)
	Case "visits"	strAsgSortByFld = formatSortingField("SumVisits", "SUM(query_visits)", ASG_USE_MYSQL)
	Case "engine"	strAsgSortByFld = formatSortingField("engine_name", "engine_name", ASG_USE_MYSQL)
	Case else		strAsgSortByFld = formatSortingField("SumVisits", "SUM(query_visits)", ASG_USE_MYSQL)
End Select

' Call advanced data sorting configuration and variables
Call dimAdvDataSorting

' Call advanced details sorting configuration and variables
Call dimAdvDetDataSorting

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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Marketing & " &raquo;</span> " & MENUSECTION_SearchEngines %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_SearchEngines
Response.Write(builTableTlayout("", "open", MENUSECTION_SearchEngines))

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
	<td width="5%"  class="treport_title">&nbsp;</td>
	<td width="28%" class="treport_title"><%= TXT_search_engine %>
		<a href="?<%= strAsgAppend & "&amp;sortby=engine&amp;sortorder=DESC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_search_engine & "&nbsp;" & TXT_orderDesc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_search_engine & "&nbsp;" & TXT_orderDesc %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=engine&amp;sortorder=ASC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_search_engine & "&nbsp;" & TXT_orderAsc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_search_engine & "&nbsp;" & TXT_orderAsc %>" /></a></td>
	<td width="12%" class="treport_title"><%= TXT_traffic %></td>
	<td width="50%" class="treport_title">
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
if strAsgMode = "month" then 
	strAsgSQL = "SELECT engine_name, SUM(query_hits) AS SumHits, SUM(query_visits) As SumVisits " &_
		"FROM " & ASG_TABLE_PREFIX & "query " &_
		"WHERE query_period = '" & strAsgPeriod & "' "
	strAsgSQL = searchFor(strAsgSQL, false)		' search
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT engine_name, SUM(query_hits) AS SumHits, SUM(query_visits) As SumVisits " &_
		"FROM " & ASG_TABLE_PREFIX & "query "
	strAsgSQL = searchFor(strAsgSQL, true)		' search
end if
'Group information by following fields
strAsgSQL = strAsgSQL & " GROUP BY engine_name "
'Order record by selected field 
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
	if objAsgRs.EOF Then
				
		' If it is a search query then show no results advise
		if Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 then

			' No current record for search terms		
			Response.Write(buildTableContNoRecord(4, "search"))
					
		' Else show general no record information
		Else
	
			' No current record			
			Response.Write(buildTableContNoRecord(4, "standard"))
					
		End If
				
	Else
			
		objAsgRs.PageSize = RecordsPerPage
		objAsgRs.AbsolutePage = page
				
		for loopAdvDataSorting = 1 to RecordsPerPage
					
			if not objAsgRs.EOF then

				strAsgReportEngine = objAsgRs("engine_name")
					
%>		  
  <tr <%= buildTableContRollover("treport_row") %> >
	<td class="treport_col" style="text-align: center;"><img src="<%= STR_ASG_SKIN_PATH_IMAGE %>def/engine.asp?icon=<%= strAsgReportEngine %>" alt="<%= strAsgReportEngine %>" /></td>
	<td class="treport_col" style="text-align: left;"><%
				
		' Create a tmp querystring without new values
		strAsgAppend = appendToQuerystring("details")
	
		' Write an anchor
		Response.Write("<div id=""" & strAsgReportEngine & """></div>")
				
		' Write the detail link with the current detail value and the detpage set to 1
		Response.Write(vbCrLf & "<a href=""?" & strAsgAppend & "&amp;detpage=1&amp;details=" & strAsgReportEngine & "#" & strAsgReportEngine & """ title=""" & TXT_Searchengine & """>")
		Response.Write(showIconDetails(strAsgReportEngine, strAsgDetails, TXT_Searchengine) & "</a>&nbsp;")
		
		Response.Write(searchTerms(strAsgReportEngine, "engine_name")) 
		
	%></td>
	<td class="treport_col" style="text-align: right;"><%= objAsgRs("SumHits") & "<br />" & objAsgRs("SumVisits") %></td>
	<td class="treport_col" style="text-align: left;">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_hits_h.gif" width="<%= FormatNumber(objAsgRs("SumHits") * intAsgBarPart, 2) %>" height="9" alt="<%= TXT_pageviews %>" align="middle" /> [<%= calcPercValue(intAsgTotMonthHits, objAsgRs("SumHits")) %>]<br />
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_visits_h.gif" width="<%= FormatNumber(objAsgRs("SumVisits") * intAsgBarPart, 2) %>" height="9" alt="<%= TXT_visits %>" align="middle" /> [<%= calcPercValue(intAsgTotMonthVisits, objAsgRs("SumVisits")) %>]
	</td>
  </tr>
<%
	' Show detail information
	if Len(strAsgDetails) > 0 AND strAsgReportEngine = strAsgDetails then
		
		Dim objAsgRs2
		Set objAsgRs2 = Server.CreateObject("ADODB.Recordset")
				
		if strAsgMode = "month" then 
			strAsgSQL = "SELECT query_keyphrase, query_hits, query_visits, query_serp_page " &_ 
				"FROM " & ASG_TABLE_PREFIX & "query " &_
				"WHERE engine_name = '" & strAsgDetails & "' AND query_period = '" & strAsgPeriod & "' "
		elseif strAsgMode = "all" then 
			strAsgSQL = "SELECT query_keyphrase, query_hits, query_visits, query_serp_page " &_
				"FROM " & ASG_TABLE_PREFIX & "query " &_
				"WHERE engine_name = '" & strAsgDetails & "' "
		end if
		strAsgSQL = strAsgSQL & " ORDER BY query_visits DESC, query_hits DESC"
		
%>
  <tr class="treport_rowdetails">
	<td class="treport_coldetails" style="text-align: center;">&nbsp;</td>
	<td class="treport_coldetails" style="text-align: center;" colspan="4">
		<!-- details -->
		<div id="tdetails">
			<div id="tdetails_bg">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
			  <tr>
				<td width="76%" class="tdetails_title"><%= TXT_search_engine %></td>
				<td width="12%" class="tdetails_title"><%= TXT_Pages %></td>
				<td width="12%" class="tdetails_title"><%= TXT_visits %></td>
			  </tr>
			  <% 
				
				' Set Rs properties
				if ASG_USE_MYSQL then
					objAsgRs2.CursorLocation = 3
				end if
				objAsgRs2.CursorType = 1
				objAsgRs2.LockType = 3
			
				' Open Rs
				objAsgRs2.Open strAsgSQL, objAsgConn
							
					' The recordset is empty
					If objAsgRs2.EOF Then
								
						' If it is a search query then show no results advise
						if Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 then
				
							' No current record for search terms		
							Response.Write(buildTableContNoRecord(3, "search"))
									
						' Else show general no record information
						Else
					
							' No current record			
							Response.Write(buildTableContNoRecord(3, "standard"))
									
						End If
								
					Else

						objAsgRs2.PageSize = detRecordsPerPage
						objAsgRs2.AbsolutePage = detpage
						
						' Create a tmp querystring without new values
						' strAsgAppend = appendToQuerystring("sortby||sortorder||details")
												
						for loopAdvDetDataSorting = 1 to detRecordsPerPage
							if not objAsgRs2.EOF then			

			  %>
			  <tr <%= buildTableContRollover("tdetails_row") %> >
				<td class="treport_col" style="text-align: left;"><% if objAsgRs2("query_serp_page") > 0 then Response.Write("&nbsp;<a href=""serp.asp?serp=" & objAsgRs2("query_serp_page") & "&amp;periodm=" & intAsgPeriodM & "&amp;periody=" & intAsgPeriodY & """ title=""" & TXT_Queries & "&nbsp;" & TXT_On  & "&nbsp;" & objAsgRs2("query_serp_page") & "&deg;&nbsp;" & TXT_page & """><span class=""notetext"">[" & objAsgRs2("query_serp_page") & "]</span></a>") %>&nbsp;<%= objAsgRs2("query_keyphrase") %></td>
				<td class="treport_col" style="text-align: right;"><%= objAsgRs2("query_hits") %></td>
				<td class="treport_col" style="text-align: right;"><%= objAsgRs2("query_visits") %></td>
			  </tr>
			  <%
						
							objAsgRs2.MoveNext
							End If
						Next
					End If
				
				' Advanced details sorting
				strLayerAdvDataSorting = buildLayerAdvDetDataSorting()
									  
				objAsgRs2.Close
				Set objAsgRs2 = Nothing

			%></table></div><%
				
				' Advanced details sorting layer
				Response.Write(strLayerAdvDataSorting)
			
			%></div>
		<!-- / details -->
	</td>
  </tr>
<%
	' Show detail information
	end If

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
				
' :: Close tlayout :: MENUSECTION_SearchEngines
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
	Response.Write(buildLayerSearch("", "engine_name"))
				
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