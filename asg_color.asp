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

' Additional vars
Dim intAsgBarPartReso
Dim intAsgMaxReso
Dim intAsgBarPartColor
Dim intAsgMaxColor

' Graph
Const intAsgBarMaxWidth = 300		' Max total column width

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
if intAsgTotMonthVisits = 0 OR "[]" & intAsgTotMonthVisits = "[]" then intAsgTotMonthVisits = 0

' Get the max item value
if strAsgMode = "month" then 
	strAsgSQL = "SELECT SUM(system_hits) AS SumHits "&_
		"FROM " & ASG_TABLE_PREFIX & "system " &_
		"WHERE system_period = '" & strAsgPeriod & "' "
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT SUM(system_hits) AS SumHits " &_
		"FROM " & ASG_TABLE_PREFIX & "system "
End If
' Change grouping mode depending on database
If ASG_USE_MYSQL then
	strAsgSQL = strAsgSQL & " GROUP BY system_reso " &_
		"ORDER BY SumHits DESC"
else
	strAsgSQL = strAsgSQL & " GROUP BY system_reso " &_
		"ORDER BY SUM(system_hits) DESC"
end if
' Set Rs properties
if ASG_USE_MYSQL then
	objAsgRs.CursorLocation = 3
end if
objAsgRs.CursorType = 1
objAsgRs.LockType = 3
' Open Rs
objAsgRs.Open strAsgSQL, objAsgConn
If objAsgRs.EOF Then
	intAsgMaxReso = 0
Else
	objAsgRs.MoveFirst
	intAsgMaxReso = objAsgRs("SumHits")
End If
objAsgRs.Close
' Calculate the minimal part to build the graph
if intAsgMaxReso = 0 OR "[]" & intAsgMaxReso = "[]" then intAsgMaxReso = 1
intAsgBarPartReso = intAsgBarMaxWidth / intAsgMaxReso

' Get the max item value
if strAsgMode = "month" then 
	strAsgSQL = "SELECT SUM(system_hits) AS SumHits " &_
		"FROM " & ASG_TABLE_PREFIX & "system " &_
		"WHERE system_period = '" & strAsgPeriod & "' "
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT SUM(system_hits) AS SumHits " &_
		"FROM " & ASG_TABLE_PREFIX & "system "
End If
' Change grouping mode depending on database
If ASG_USE_MYSQL then
	strAsgSQL = strAsgSQL & "GROUP BY system_color " &_
		"ORDER BY SumHits DESC"
else
	strAsgSQL = strAsgSQL & "GROUP BY system_color " &_
		"ORDER BY SUM(Hits) DESC"
end if
' Set Rs properties
if ASG_USE_MYSQL then
	objAsgRs.CursorLocation = 3
end if
objAsgRs.CursorType = 1
objAsgRs.LockType = 3
' Open Rs
objAsgRs.Open strAsgSQL, objAsgConn
If objAsgRs.EOF Then
	intAsgMaxColor = 0
Else
	objAsgRs.MoveFirst
	intAsgMaxColor = objAsgRs("SumHits")
End If
objAsgRs.Close
' Calculate the minimal part to build the graph
if intAsgMaxColor = 0 OR "[]" & intAsgMaxColor = "[]" then intAsgMaxColor = 1
intAsgBarPartColor = intAsgBarMaxWidth / intAsgMaxColor


' Read sorting order from querystring
' Filter QS values and associate them 
' with their respective database fields
Select Case strAsgSortBy
	Case "hits" 	strAsgSortByFld = formatSortingField("SumHits", "SUM(system_hits)", ASG_USE_MYSQL)
	Case "visits"	strAsgSortByFld = formatSortingField("SumVisits", "SUM(system_visits)", ASG_USE_MYSQL)
	Case "value"	
		'
	Case else		strAsgSortByFld = formatSortingField("SumVisits", "SUM(system_visits)", ASG_USE_MYSQL)
End Select

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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Visitors & " &raquo; " & MENUSECTION_VisitorSystems & " &raquo;</span> " & MENUSECTION_ResoBit %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_Reso
Response.Write(builTableTlayout("", "open", MENUSECTION_Reso))

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
	<td width="5%"  class="treport_title">&nbsp;</td>
	<td width="33%" class="treport_title"><%= TXT_reso %>
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_reso) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=value&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=value&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="12%" class="treport_title"><%= TXT_traffic %></td>
	<td width="50%" class="treport_title">
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
strAsgSQL = "SELECT system_reso, SUM(system_hits) AS SumHits, SUM(system_visits) As SumVisits " &_
"FROM " & ASG_TABLE_PREFIX & "system "
if strAsgMode = "month" then 
	strAsgSQL = strAsgSQL & "WHERE system_period = '" & strAsgPeriod & "' "
	' Call the function to search into the database if there are enought information to do that
	strAsgSQL = searchFor(strAsgSQL, false)
elseif strAsgMode = "all" then 
	' Call the function to search into the database if there are enought information to do that
	strAsgSQL = searchFor(strAsgSQL, true)
end if
' Group information by following fields
strAsgSQL = strAsgSQL & "GROUP BY system_reso "
' Order record by selected field 
if strAsgSortBy = "value" then
	strAsgSQL = strAsgSQL & "ORDER BY system_reso " & strAsgSortOrder & ""
else
	strAsgSQL = strAsgSQL & "ORDER BY " & strAsgSortByFld & " " & strAsgSortOrder & ""
end if

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
		
		Do While NOT objAsgRs.EOF
%>		  
  <tr <%= buildTableContRollover("treport_row") %> >
	<td class="treport_col" align="center">&nbsp;</td>
	<td class="treport_col" align="left"><%= searchTerms(objAsgRs("system_reso"), "system_reso", asgSearchfor, asgSearchIn) %></td>
	<td class="treport_col" align="right"><%= objAsgRs("SumHits") & "<br />" & objAsgRs("SumVisits") %></td>
	<td class="treport_col" style="text-align: left;">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_hits_h.gif" width="<%= FormatNumber(objAsgRs("SumHits") * intAsgBarPartReso, 2) %>" height="9" alt="<%= TXT_pageviews %>" align="middle" /> [<%= calcPercValue(intAsgTotMonthHits, objAsgRs("SumHits")) %>]<br />
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_visits_h.gif" width="<%= FormatNumber(objAsgRs("SumVisits") * intAsgBarPartReso, 2) %>" height="9" alt="<%= TXT_visits %>" align="middle" /> [<%= calcPercValue(intAsgTotMonthVisits, objAsgRs("SumVisits")) %>]
	</td>
  </tr>
<%
				
		objAsgRs.MoveNext
		Loop
	End If

%></table><%

objAsgRs.Close
		
	' Report Legend
	Response.Write(buildLayerReportLegend())
				
' :: Close tlayout :: MENUSECTION_Reso
Response.Write(builTableTlayout("", "close", ""))


Response.Write(vbCrLf & "<br />")


' :: Open tlayout :: MENUSECTION_Colors
Response.Write(builTableTlayout("", "open", MENUSECTION_Colors))

%>		  
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
	<td width="5%"  class="treport_title">&nbsp;</td>
	<td width="33%" class="treport_title"><%= TXT_color %>
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_color) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=value&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=value&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="12%" class="treport_title"><%= TXT_traffic %></td>
	<td width="50%" class="treport_title">
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
strAsgSQL = "SELECT system_color, SUM(system_hits) AS SumHits, SUM(system_visits) As SumVisits " &_
	"FROM " & ASG_TABLE_PREFIX & "system "
if strAsgMode = "month" then 
	strAsgSQL = strAsgSQL & "WHERE system_period = '" & strAsgPeriod & "'  "
	strAsgSQL = searchFor(strAsgSQL, false)
elseif strAsgMode = "all" then 
	strAsgSQL = searchFor(strAsgSQL, true)
end if
' Group information by following fields
strAsgSQL = strAsgSQL & " GROUP BY system_color "
' Order record by selected field 
if strAsgSortBy = "value" then
	strAsgSQL = strAsgSQL & " ORDER BY system_color " & strAsgSortOrder & ""
else
	strAsgSQL = strAsgSQL & " ORDER BY " & strAsgSortByFld & " " & strAsgSortOrder & ""
end if

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
			Response.Write(buildTableContNoRecord(4, "search"))
					
		' Else show general no record information
		else
	
			' No current record			
			Response.Write(buildTableContNoRecord(4, "standard"))
					
		end if
				
	else
		
		Do While NOT objAsgRs.EOF
%>		  
  <tr <%= buildTableContRollover("treport_row") %> >
	<td class="treport_col" align="center">&nbsp;</td>
	<td class="treport_col" align="left"><%= searchTerms(objAsgRs("system_color"), "system_color", asgSearchfor, asgSearchIn) %> bit</td>
	<td class="treport_col" align="right"><%= objAsgRs("SumHits") & "<br />" & objAsgRs("SumVisits") %></td>
	<td class="treport_col" style="text-align: left;">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_hits_h.gif" width="<%= FormatNumber(objAsgRs("SumHits") * intAsgBarPartColor, 2) %>" height="9" alt="<%= TXT_pageviews %>" align="middle" /> [<%= calcPercValue(intAsgTotMonthHits, objAsgRs("SumHits")) %>]<br />
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>bar_graph_image_visits_h.gif" width="<%= FormatNumber(objAsgRs("SumVisits") * intAsgBarPartColor, 2) %>" height="9" alt="<%= TXT_visits %>" align="middle" /> [<%= calcPercValue(intAsgTotMonthVisits, objAsgRs("SumVisits")) %>]
	</td>
  </tr>
<%
				
		objAsgRs.MoveNext
		Loop
	end If

%></table><%

objAsgRs.Close

' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing
		
	' Report Legend
	Response.Write(buildLayerReportLegend())

' :: Close tlayout :: MENUSECTION_Colors
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
	Response.Write(buildLayerSearch("", "system_reso|system_color"))
				
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