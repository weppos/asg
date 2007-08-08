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
Dim strAsgPage
Dim strAsgTmpPage
Dim strAsgDetails
strAsgDetails = Request.QueryString("details")
Dim strAsgGroup
strAsgGroup = formatSetting("group", "none")

' Create a tmp querystring without new values
strAsgAppend = appendToQuerystring("sortby||sortorder||details")


' Read sorting order from querystring
' Filter QS values and associate them 
' with their respective database fields
select case strAsgSortBy
	Case "hits" 	strAsgSortByFld = formatSortingField("SumHits", "SUM(page_hits)", ASG_USE_MYSQL)
	Case "visits"	strAsgSortByFld = formatSortingField("SumVisits", "SUM(page_visits)", ASG_USE_MYSQL)
	Case "path"		strAsgSortByFld = formatSortingField("page_path", "page_path", ASG_USE_MYSQL)
	Case "page"
		if strAsgGroup = "path" then
			strAsgSortByFld = formatSortingField("MaxPage", "Max(page_path)", ASG_USE_MYSQL)
		else
			strAsgSortByFld = formatSortingField("page_path, page_qs", "page_path, page_qs", ASG_USE_MYSQL)
		end if
	Case else		strAsgSortByFld = formatSortingField("SumVisits", "SUM(page_visits)", ASG_USE_MYSQL)
end select

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
<script language="JavaScript" type="text/javascript" src="tip_info.js.asp"></script>

<div align="center">
	<div id="layout">
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Navigation & " &raquo;</span> " & MENUSECTION_VisitedPages %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_VisitedPages
Response.Write(builTableTlayout("", "open", MENUSECTION_VisitedPages))

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
	<td width="5%"  class="treport_title">&nbsp;</td>
	<% if strAsgGroup = "none" then %>
	<td width="73%" class="treport_title"><%= TXT_Pages %>
		<a href="?<%= strAsgAppend & "&amp;sortby=page&amp;sortorder=DESC" %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_page & "&nbsp;" & TXT_orderDesc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_page & "&nbsp;" & TXT_orderDesc %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=page&amp;sortorder=ASC" %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_page & "&nbsp;" & TXT_orderAsc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_page & "&nbsp;" & TXT_orderAsc %>" /></a></td>
	<% elseif strAsgGroup = "path" then %>
	<td width="73%" class="treport_title"><%= TXT_Path %>
		<a href="?<%= strAsgAppend & "&amp;sortby=path&amp;sortorder=DESC" %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_Path & "&nbsp;" & TXT_orderDesc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_Path & "&nbsp;" & TXT_orderDesc %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=path&amp;sortorder=ASC" %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_Path & "&nbsp;" & TXT_orderAsc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_Path & "&nbsp;" & TXT_orderAsc %>" /></a></td>
	<% end if %>			
	<td width="12%" class="treport_title"><%= TXT_pageviews_title %>
		<a href="?<%= strAsgAppend & "&amp;sortby=hits&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=hits&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="12%" class="treport_title"><%= TXT_visits_title %>
		<a href="?<%= strAsgAppend & "&amp;sortby=visits&amp;sortorder=DESC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=visits&amp;sortorder=ASC" %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
  </tr>
<%

' Initialise SQL string to select data
if strAsgGroup = "none" then
	strAsgSQL = "SELECT page_path, page_qs, SUM(page_hits) AS SumHits, SUM(page_visits) AS SumVisits "
elseif strAsgGroup = "path" then
	strAsgSQL = "SELECT page_path, SUM(page_hits) AS SumHits, SUM(page_visits) AS SumVisits "
end if
strAsgSQL = strAsgSQL & "FROM " & ASG_TABLE_PREFIX & "page "
' Month and Year
if strAsgMode = "month" then
	strAsgSQL = strAsgSQL & "WHERE page_period = '" & strAsgPeriod & "' "
	strAsgSQL = searchFor(strAsgSQL, false)		' Search
elseif strAsgMode = "all" then
	strAsgSQL = searchFor(strAsgSQL, true)		' Search
end if
' Group information by following fields
if strAsgGroup = "none" then
	strAsgSQL = strAsgSQL & "GROUP BY page_path, page_qs "
elseif strAsgGroup = "path" then
	strAsgSQL = strAsgSQL & "GROUP BY page_path "
end If
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
	if objAsgRs.EOF then
				
		' If it is a search query then show no results advise
		if Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 then

			' No current record for search terms		
			Response.Write(buildTableContNoRecord(4, "search"))
					
		' Else show general no record information
		else
	
			' No current record			
			Response.Write(buildTableContNoRecord(4, "standard"))
					
		end If
				
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
	<td class="treport_col" style="text-align: left;"><%

	if strAsgGroup = "none" then
		
		strAsgPage = objAsgRs("page_path")
		if Len(objAsgRs("page_qs")) > 0 then strAsgPage = strAsgPage & "?" & objAsgRs("page_qs")
				
		' Trim long strings
		strAsgTmpPage = stripValueTooLong(strAsgPage, 65, 30, 30)
		strAsgTmpPage = searchTerms(strAsgTmpPage, "page_path")
		strAsgTmpPage = searchTerms(strAsgTmpPage, "page_qs")
				
		' :: info - page trimmed ::
		if Len(strAsgPage) > 65 then
			Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[0],Style[1])"" onmouseout=""htm()"" /> ")
		end if

		' Link the page
		Response.Write("<a href=""" & Replace(strAsgPage, "[...]", "") & """ target=""_blank"" title=""" & TXT_gotoPage & "&nbsp;" & strAsgPage & """>")				
				
		' Write the page
		Response.Write(strAsgTmpPage & "</a>")				
		' Response.Write(chooseDomainIcon(objAsgRs("page_path"), "classic"))

	elseif strAsgGroup = "path" then
						
		' Create a tmp querystring without new values
		strAsgAppend = appendToQuerystring("details||detpage")
				
		' Trim long strings
		strAsgTmpPage = searchTerms(stripValueTooLong(objAsgRs("page_path"), 65, 30, 30), "page_path")
	
		' Write an anchor
		Response.Write("<div id=""" & objAsgRs("page_path") & """></div>")
				
		' Write the detail link with the current detail value and the detpage set to 1
		Response.Write(vbCrLf & "<a href=""?" & strAsgAppend & "&amp;detpage=1&amp;details=" & objAsgRs("page_path") & "#" & objAsgRs("page_path") & """ title=""" & TXT_Path & """>")
		Response.Write(showIconDetails(objAsgRs("page_path"), strAsgDetails, TXT_Path) & "</a>&nbsp;")
				
		' :: info - page trimmed ::
		if Len(objAsgRs("page_path")) > 65 then
			Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[0],Style[1])"" onmouseout=""htm()"" /> ")
		end if
				
		' Write the page
		Response.Write(strAsgTmpPage)				
		' Response.Write(chooseDomainIcon(objAsgRs("page_path"), "classic"))

		end If
				
	%></td>
	<td class="treport_col" style="text-align: right;"><%= objAsgRs("SumHits") %></td>
	<td class="treport_col" style="text-align: right;"><%= objAsgRs("SumVisits") %></td>
  </tr>
<% 
	' If data is grouped by path then show detail information
	if strAsgGroup = "path" then
		
		' Show detail information	
		if Len(strAsgDetails) > 0 AND objAsgRs("page_path") = strAsgDetails then
				
			Dim objAsgRs2
			Set objAsgRs2 = Server.CreateObject("ADODB.Recordset")
				
			strAsgSQL = "SELECT page_qs, page_hits, page_visits " &_
				"FROM " & ASG_TABLE_PREFIX & "page " &_
				"WHERE page_path = '" & strAsgDetails & "' "
				if strAsgMode = "month" then strAsgSQL = strAsgSQL & "AND page_period = '" & strAsgPeriod & "' " 
				strAsgSQL = strAsgSQL & "ORDER BY page_visits DESC, page_hits DESC "

%>
  <tr class="treport_rowdetails">
	<td class="treport_coldetails" style="text-align: center;">&nbsp;</td>
	<td class="treport_coldetails" style="text-align: center;" colspan="3">
		<!-- details -->
		<div id="tdetails">
			<div id="tdetails_bg">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
			  <tr>
				<td width="68%" class="tdetails_title"><%= TXT_page %></td>
				<td width="12%" class="tdetails_title"><%= TXT_pageviews_title %></td>
				<td width="12%" class="tdetails_title"><%= TXT_visits_title %></td>
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
					if objAsgRs2.EOF then
								
						' If it is a search query then show no results advise
						if Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 then
				
							' No current record for search terms		
							Response.Write(buildTableContNoRecord(3, "search"))
									
						' Else show general no record information
						else
					
							' No current record			
							Response.Write(buildTableContNoRecord(3, "standard"))
									
						end if
								
					else

						objAsgRs2.PageSize = detRecordsPerPage
						objAsgRs2.AbsolutePage = detpage
						
						for loopAdvDetDataSorting = 1 to detRecordsPerPage
							if not objAsgRs2.EOF then			

			  %>
			  <tr <%= buildTableContRollover("tdetails_row") %> >
				<td class="treport_col" style="text-align: left;"><%
					
					strAsgPage = objAsgRs("page_path") & "?" & objAsgRs2("page_qs")
					if Len(objAsgRs2("page_qs")) > 0 then strAsgPage = strAsgPage & "?" & objAsgRs2("page_qs")
							
					' Trim long strings
					strAsgTmpPage = searchTerms(stripValueTooLong(strAsgPage, 65, 30, 30), "page_path")
							
					' :: info - page trimmed ::
					if Len(strAsgPage) > 65 then
						Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[0],Style[1])"" onmouseout=""htm()"" /> ")
					end if
			
					' Link the page
					Response.Write("<a href=""" & Replace(strAsgPage, "[...]", "") & """ target=""_blank"" title=""" & TXT_gotoPage & "&nbsp;" & strAsgPage & """>")				
							
					' Write the page
					Response.Write(strAsgTmpPage & "</a>")				
				
					%></td>
				<td class="treport_col" style="text-align: right;"><%= objAsgRs2("page_hits") %></td>
				<td class="treport_col" style="text-align: right;"><%= objAsgRs2("page_visits") %></td>
			  </tr>
			  <%
						
							objAsgRs2.MoveNext
							end if  ' EOF
						next
					end if  ' EOF
				
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

	' Show detail information because grouped by domain
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
				
' :: Close tlayout :: MENUSECTION_Referers
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
	Response.Write(buildLayerGroup("none|path", "none|path"))

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
	Response.Write(buildLayerSearch("", "page_path|page_qs"))
				
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