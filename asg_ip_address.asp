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

' Other variables
Dim intAsgCounter

' Create a tmp querystring without new values
strAsgAppend = appendToQuerystring("sortby||sortorder")


' Read sorting order from querystring
' Filter QS values and associate them 
' with their respective database fields
select case strAsgSortBy
	Case "hits" 	strAsgSortByFld = formatSortingField("SumHits", "SUM(ip_hits)", ASG_USE_MYSQL)
	Case "visits"	strAsgSortByFld = formatSortingField("SumVisits", "SUM(ip_visits)", ASG_USE_MYSQL)
	Case "ip"		strAsgSortByFld = formatSortingField("ip_address", "ip_address", ASG_USE_MYSQL)
	Case "date"		strAsgSortByFld = formatSortingField("MaxData", "MAX(ip_last_access)", ASG_USE_MYSQL)
	Case else		strAsgSortByFld = formatSortingField("SumVisits", "SUM(ip_visits)", ASG_USE_MYSQL)
end select

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
<script language="JavaScript" type="text/javascript" src="tip_info.js.asp"></script>

<div align="center">
	<div id="layout">
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Visitors & " &raquo;</span> " & MENUSECTION_IpAddresses %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_IpAddresses
Response.Write(builTableTlayout("", "open", MENUSECTION_IpAddresses))

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
	<td width="5%"  class="treport_title">&nbsp;</td>
	<td width="35%" class="treport_title"><%= TXT_IPAddress %>
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_IPAddress) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=ip&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=ip&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="31%" class="treport_title"><%= TXT_lastAccess %>
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_lastAccess) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=date&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=date&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="12%" class="treport_title"><%= TXT_pageviews_title %>
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_pageviews) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=hits&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=hits&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="12%" class="treport_title"><%= TXT_visits_title %>
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_visits) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=visits&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=visits&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a>
  </tr>
<%


' Initialise SQL string to select data
strAsgSQL = "SELECT ip_address, Max(ip_last_access) AS MaxData, Sum(ip_visits) AS SumVisits, Sum(ip_hits) AS SumHits " &_
	"FROM " & ASG_TABLE_PREFIX & "ip "
' Month and Year
if strAsgMode = "month" then 
	strAsgSQL = strAsgSQL & "WHERE ip_period = '" & strAsgPeriod & "' "
	strAsgSQL = searchFor(strAsgSQL, false)		' Search
elseif strAsgMode = "all" then 
	strAsgSQL = searchFor(strAsgSQL, true)		' Search
end if
' Group information by following fields
strAsgSQL = strAsgSQL & "GROUP BY ip_address "
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
			Response.Write(buildTableContNoRecord(5, "search"))
					
		' Else show general no record information
		else
	
			' No current record			
			Response.Write(buildTableContNoRecord(5, "standard"))
					
		end if
				
	else
			
		objAsgRs.PageSize = RecordsPerPage
		objAsgRs.AbsolutePage = page
		intAsgCounter = (RecordsPerPage * (page-1))
				
		for loopAdvDataSorting = 1 to RecordsPerPage
					
			if not objAsgRs.EOF then			
			intAsgCounter = intAsgCounter + 1
					
%>		  
  <tr <%= buildTableContRollover("treport_row") %> >
	<td class="treport_col" style="text-align: center;"><%= intAsgCounter %></td>
	<td class="treport_col" style="text-align: left;"><%= searchTerms(objAsgRs("ip_address"), "ip_address", asgSearchfor, asgSearchIn) %></td>
	<td class="treport_col" style="text-align: center;"><%= formatDateTimeValue(objAsgRs("MaxData"), "Date") & "&nbsp;" & TXT_time_at & "&nbsp;" & formatDateTimeValue(objAsgRs("MaxData"), "Time") %></td>
	<td class="treport_col" style="text-align: right;"><%= objAsgRs("SumHits") %></td>
	<td class="treport_col" style="text-align: right;"><%= objAsgRs("SumVisits") %></td>
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
	Response.Write(buildLayerSearch("", "ip_address"))
				
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