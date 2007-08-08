<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<!--#include file="lib/functions_images.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


' ***** Search for a single querystring element in a list of comma separated items.
private function searchTypeElement(argList, argElement)

	Dim ii
	Dim aryList
	aryList = Split(argList, ",")

	for ii = 0 to Ubound(aryList)
		if Cint(argElement) = Cint(aryList(ii)) then exit for
	next
	
	' Return the function
	if ii > Ubound(aryList) then
		searchTypeElement = false
	else
		searchTypeElement = true
	end if

end function


'// WARNING! Program protection.
'	Changing default values may allow users to access the page.
Call checkPermission("True", "False", "False", appAsgSecurity)

' Include commons variable, declarations 
' and data filtering.
%><!--#include file="includes/variables.inc.asp" --><%

' By default select all external referers
Const STR_ASG_TYPE_DEFAULT = "4,5"
' Regexp pattern to chek new search engines
Const STR_ASG_DEBUG_SEARCHENGINES_PATTERN = "((http://.*search[\/|\?])|(\?\S*search=[^&]+|\?\S*query=[^&]+))"

' Other variables
Dim ii
Dim intAsgCounter
Dim strAsgTmpReferer
Dim strAsgDetails
strAsgDetails = Request.QueryString("details")
Dim strAsgGroup
strAsgGroup = formatSetting("group", "none")
Dim strAsgType, aryAsgType
strAsgType = formatSetting("type", STR_ASG_TYPE_DEFAULT)
Dim strAsgReportReferer
Dim strAsgReportDomain

' Split type values to get all type id and check them
if strAsgGroup <> STR_ASG_TYPE_DEFAULT then
	aryAsgType = Split(strAsgType, ",")
	' Clean the variable
	strAsgType = ""
	' Build the type string
	for ii = 0 to Ubound(aryAsgType)
		if IsNumeric(aryAsgType(ii)) then strAsgType = strAsgType & aryAsgType(ii) & ","
	next
	' Remove the trailing comma
	if Len(strAsgType) > 0 then strAsgType = Left(strAsgType, Len(strAsgType) - 1)
	' Check that the string is not empty
	if not Len(strAsgType) > 0 then strAsgType = STR_ASG_TYPE_DEFAULT
end if

' Create a tmp querystring without new values
strAsgAppend = appendToQuerystring("sortby||sortorder||details")


' Read sorting order from querystring
' Filter QS values and associate them 
' with their respective database fields
Select Case strAsgSortBy
	Case "visits"	strAsgSortByFld = formatSortingField("SumVisits", "SUM(referer_visits)", ASG_USE_MYSQL)
	Case "group"
		if strAsgGroup	= "none" then	
			strAsgSortByFld = "referer_url"
		elseif strAsgGroup	= "domain" then	
			strAsgSortByFld = "referer_domain"
		end if
	Case "date"		strAsgSortByFld = formatSortingField("MaxDate", "MAX(referer_last_access)", ASG_USE_MYSQL)
	Case else		strAsgSortByFld = formatSortingField("SumVisits", "SUM(referer_visits)", ASG_USE_MYSQL)
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
<script language="JavaScript" type="text/javascript" src="tip_info.js.asp"></script>

<div align="center">
	<div id="layout">
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Marketing & " &raquo;</span> " & MENUSECTION_Referers %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_Referers
Response.Write(builTableTlayout("", "open", MENUSECTION_Referers))

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
	<td width="5%"  class="treport_title">&nbsp;</td>
	<td width="55%" class="treport_title"><% 
	if strAsgGroup	= "none" then
		Response.Write(TXT_referer) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=group&amp;sortorder=DESC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_referer & "&nbsp;" & TXT_orderDesc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_referer & "&nbsp;" & TXT_orderDesc %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=group&amp;sortorder=ASC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_referer & "&nbsp;" & TXT_orderAsc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_referer & "&nbsp;" & TXT_orderAsc %>" /></a><%
	elseif strAsgGroup = "domain" then	
		Response.Write(TXT_domain) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=group&amp;sortorder=DESC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_domain & "&nbsp;" & TXT_orderDesc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_domain & "&nbsp;" & TXT_orderDesc %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=group&amp;sortorder=ASC&amp;details=" & strAsgDetails %>" title="<%= TXT_orderBy & "&nbsp;" & TXT_domain & "&nbsp;" & TXT_orderAsc %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= TXT_orderBy & "&nbsp;" & TXT_domain & "&nbsp;" & TXT_orderAsc %>" /></a><%
	end If
	%></td>
	<td width="23%" class="treport_title"><%= TXT_lastAccess %> 
	<% strOrderBy_message = Replace(TXT_orderBy_schema, "$field$", TXT_lastAccess) %>
		<a href="?<%= strAsgAppend & "&amp;sortby=date&amp;sortorder=DESC&amp;details=" & strAsgDetails %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=date&amp;sortorder=ASC&amp;details=" & strAsgDetails %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
	<td width="12%" class="treport_title"><%= TXT_visits %> 
		<a href="?<%= strAsgAppend & "&amp;sortby=visits&amp;sortorder=DESC&amp;details=" & strAsgDetails %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_desc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderDesc) %>" /></a>
		<a href="?<%= strAsgAppend & "&amp;sortby=visits&amp;sortorder=ASC&amp;details=" & strAsgDetails %>" title="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/order_asc.png" border="0" align="middle" alt="<%= Replace(strOrderBy_message, "$order$", TXT_orderAsc) %>" /></a></td>
  </tr>
<%

' Initialise SQL string to select data
strAsgSQL = "SELECT $var1$, referer_type, Max(referer_last_access) AS MaxDate, SUM(referer_visits) AS SumVisits " &_
	"FROM " & ASG_TABLE_PREFIX & "referer " 
' 'group' group by condition
if strAsgGroup = "none" then
	' Get referers grouping by Referer.
	' The group by condition is useless but useful to keep compatibility with Max() syntax.
	strAsgSQL = Replace(strAsgSQL, "$var1$", "referer_url")
elseif strAsgGroup = "domain" then
	' Get referers grouping by Domain.
	strAsgSQL = Replace(strAsgSQL, "$var1$", "referer_domain")
end If
' Referer type
strAsgSQL = strAsgSQL & "WHERE referer_type IN (" & strAsgType & ") "
' Month and Year
if strAsgMode = "month" then
	strAsgSQL = strAsgSQL & "AND referer_period = '" & strAsgPeriod & "' "
	strAsgSQL = searchFor(strAsgSQL, false)
elseif strAsgMode = "all" then 
	strAsgSQL = searchFor(strAsgSQL, false)
end if
' 'group' group by condition
if strAsgGroup = "none" then
	' Get referers grouping by Referer.
	' The group by condition is useless but useful to keep compatibility with Max() syntax.
	strAsgSQL = strAsgSQL & "GROUP BY referer_url "
elseif strAsgGroup = "domain" then
	' Get referers grouping by Domain.
	strAsgSQL = strAsgSQL & "GROUP BY referer_domain "
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
			
			if strAsgGroup = "none" then
				strAsgReportReferer = objAsgRs("referer_url")
			elseif strAsgGroup = "domain" then
				strAsgReportDomain = objAsgRs("referer_domain")	
			end if			
%>		  
  <tr <%= buildTableContRollover("treport_row") %> >
	<td class="treport_col" style="text-align: center;"><%= intAsgCounter %></td>
	<td class="treport_col" style="text-align: left;"><%
		
			if strAsgGroup = "none" then
				
				' Trim long strings
				strAsgTmpReferer = searchTerms(stripValueTooLong(strAsgReportReferer, 65, 30, 30), "referer_url")
				
				' :: warning - referer longer than 250 chr ::
				strAsgTmpReferer = Replace(strAsgTmpReferer, "[...]", "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_warning.png"" alt=""" & TXT_referer_longer250_warning & """ border=""0"" align=""middle"" />")
				
				' :: info - referer trimmed ::
				if Len(strAsgReportReferer) > 65 then
					Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[0],Style[1])"" onmouseout=""htm()"" /> ")
				end if
				
				if ASG_DEBUG_SEARCHENGINES AND objAsgRs("referer_type") <> 5 then
					' :: advice - may be a search engine ::
					if regexpTest(STR_ASG_DEBUG_SEARCHENGINES_PATTERN, strAsgReportReferer) then
						Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_advice.png"" alt=""" & TXT_referer_debugengine_advice & """ border=""0"" align=""middle"" /> ")
					end if
				end if

				' Link the referer page
				Response.Write("<a href=""" & Replace(strAsgReportReferer, "[...]", "") & """ target=""_blank"" title=""" & TXT_gotoPage & "&nbsp;" & strAsgReportReferer & """>")				
				
				' Write the referer
				Response.Write(strAsgTmpReferer & "</a>")				
	
			elseif strAsgGroup = "domain" then
						
				' Create a tmp querystring without new values
				strAsgAppend = appendToQuerystring("details||detpage")
	
				' Write an anchor
				Response.Write("<div id=""" & strAsgReportDomain & """></div>")
				
				' Write the detail link with the current detail value and the detpage set to 1
				Response.Write(vbCrLf & "<a href=""?" & strAsgAppend & "&amp;detpage=1&amp;details=" & strAsgReportDomain & "#" & strAsgReportDomain & """ title=""" & TXT_domain & """>")
				Response.Write(showIconDetails(strAsgReportDomain, strAsgDetails, TXT_domain) & "</a>&nbsp;")
				
				Response.Write(searchTerms(stripValueTooLong(strAsgReportDomain, 65, 30, 30), "referer_domain"))
				
			end If
				
	%></td>
	<td class="treport_col" style="text-align: center;"><%= formatDateTimeValue(objAsgRs("MaxDate"), "Date") & "&nbsp;" & TXT_time_at & "&nbsp;" & formatDateTimeValue(objAsgRs("MaxDate"), "Time") %></td>
	<td class="treport_col" style="text-align: right;"><%= objAsgRs("SumVisits") %></td>
  </tr>
<% 
	' If data is grouped by domain then show detail information
	if strAsgGroup = "domain" then
		
		' Show detail information	
		if Len(strAsgDetails) > 0 AND strAsgReportDomain = strAsgDetails then
				
			Dim objAsgRs2
			Set objAsgRs2 = Server.CreateObject("ADODB.Recordset")
				
			strAsgSQL = "SELECT referer_url, referer_type, referer_last_access, referer_visits " &_
				"FROM " & ASG_TABLE_PREFIX & "referer " &_
				"WHERE referer_domain = '" & strAsgDetails & "' "
				if strAsgMode = "month" then strAsgSQL = strAsgSQL & "AND referer_period = '" & strAsgPeriod & "' " 
				strAsgSQL = strAsgSQL & " ORDER BY referer_visits DESC "
		
%>
  <tr class="treport_rowdetails">
	<td class="treport_coldetails" style="text-align: center;">&nbsp;</td>
	<td class="treport_coldetails" style="text-align: center;" colspan="3">
		<!-- details -->
		<div id="tdetails">
			<div id="tdetails_bg">
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
			  <tr>
				<td width="68%" class="tdetails_title"><%= TXT_referer %></td>
				<td width="20%" class="tdetails_title"><%= TXT_lastAccess %></td>
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
						
						for loopAdvDetDataSorting = 1 to detRecordsPerPage
							if not objAsgRs2.EOF then			

			  %>
			  <tr <%= buildTableContRollover("tdetails_row") %> >
				<td class="treport_col" style="text-align: left;"><%
					
					' Trim long strings
					strAsgTmpReferer = searchTerms(stripValueTooLong(objAsgRs2("referer_url"), 65, 30, 30), "referer_url")
				
					' :: warning - referer longer than 250 chr ::
					strAsgTmpReferer = Replace(strAsgTmpReferer, "[...]", "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_warning.png"" alt=""" & TXT_referer_longer250_warning & """ border=""0"" align=""middle"" />")
					
					' :: info - referer trimmed ::
					if Len(objAsgRs2("referer_url")) > 65 then
						Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[0],Style[1])"" onmouseout=""htm()"" /> ")
					end if
					
					if ASG_DEBUG_SEARCHENGINES AND objAsgRs2("referer_type") <> 5 then
						' :: advice - may be a search engine ::
						if regexpEngineDebug(STR_ASG_DEBUG_SEARCHENGINES_PATTERN, objAsgRs2("referer_url")) then
							Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_advice.png"" alt=""" & TXT_referer_debugengine_advice & """ border=""0"" align=""middle"" /> ")
						end if
					end if
	
					' Link the referer page
					Response.Write("<a href=""" & Replace(objAsgRs2("referer_url"), "[...]", "") & """ target=""_blank"" title=""" & TXT_gotoPage & "&nbsp;" & objAsgRs2("referer_url") & """>")				
					
					' Write the referer
					Response.Write(strAsgTmpReferer & "</a>")				
				
					%></td>
				<td class="treport_col" style="text-align: center;"><%= formatDateTimeValue(objAsgRs2("referer_last_access"), "Date") %></td>
				<td class="treport_col" style="text-align: right;"><%= objAsgRs2("referer_visits") %></td>
			  </tr>
			  <%
						
							objAsgRs2.MoveNext
							end if
						next
					end if
				
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


	' :: tooltip - enable search engine debug ::
	if not ASG_DEBUG_SEARCHENGINES then
	
'	strAsgTmpLayer = "<div class=""divlayer""><p style=""text-align: justify;"" class=""fldlegendtitle"">"
'	strAsgTmpLayer = strAsgTmpLayer & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/tip.png"" alt=""tooltip"" border=""0"" align=""middle"" />&nbsp;" & TXT_Tooltip
'	strAsgTmpLayer = strAsgTmpLayer & "</p></div>"
'	Response.Write(strAsgTmpLayer)
	
	strAsgTmpLayer = "<p style=""text-align: justify;"">" & TXT_referer_tip1 & "</p>"
	
		' :: Create the layer ::
		Response.Write(buildLayer("layerType", "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/tip.png"" alt=""" & TXT_Tooltip & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Tooltip, "", strAsgTmpLayer))
		
	end if
				
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
	Response.Write(buildLayerGroup("none|domain", "none|domain"))
	
	strAsgTmpLayer = "<p>"
	for ii = 1 to 5
		strAsgTmpLayer = strAsgTmpLayer & "<input type=""checkbox"" name=""type"" value=""" & ii & """"
		if searchTypeElement(strAsgType, ii) then 
			' Check the DTD
			if InStr(STR_ASG_PAGE_DOCTYPE, "XHTML") > 0 then
				strAsgTmpLayer = strAsgTmpLayer & " checked=""checked"""
			else
				strAsgTmpLayer = strAsgTmpLayer & " checked"
			end if
		end if
		strAsgTmpLayer = strAsgTmpLayer & " /> " & Eval("TXT_referer_type" & ii) & "&nbsp;"
	next
	strAsgTmpLayer = strAsgTmpLayer & "</p>"
		
		' :: Create the layer ::
		Response.Write(buildLayer("layerType", LABEL_Type, "", strAsgTmpLayer))

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
	Response.Write(buildLayerSearch("", "referer_url|referer_domain"))
				
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