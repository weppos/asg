<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<!--#include file="lib/functions_count.asp" -->
<!--#include file="lib/functions_images.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'// WARNING! Program protection.
'	Changing default values may allow users to access the page.
Call checkPermission("True", "False", "False", appAsgSecurity)


' Dimension variables
Dim strLayerAdvDataSorting
Dim ii
Dim intAsgCounter
Dim strAsgDetails
strAsgDetails = Request.QueryString("details")
Dim blnDatabaseIsEmpty
blnDatabaseIsEmpty = false

Dim intAsgRecordCountHits
Dim blnAsgShowDetails		' Set true to expand visitor details
Dim strAsgActiveRange		' Hold the active time range of the selected visitor		

Dim lngAsgReportUserId
Dim strAsgReportCountry
Dim strAsgReportPage
Dim strAsgReportReferer
Dim strAsgReportSengine
Dim strAsgReportBrowser
Dim strAsgReportOs
Dim aryAsgReportVisitLength
Dim strAsgBuffer

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
<script language="JavaScript" type="text/javascript" src="tip_info.js.asp"></script>

<div align="center">
	<div id="layout">
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Visitors & " &raquo;</span> " & MENUSECTION_VisitorDetails %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_VisitorDetails
Response.Write(builTableTlayout("", "open", MENUSECTION_VisitorDetails))

'% >
'<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
'< %		  

' Initialise SQL string to select data
if ASG_USE_MYSQL then
	strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "user "
else
	strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "user "
end if
' Call the function to search into the database if there are enought information to do that
strAsgSQL = searchFor(strAsgSQL, true)
' Group information by following fields and order by the most recent date/time
if ASG_USE_MYSQL then
	strAsgSQL = strAsgSQL & "ORDER BY user_last_access DESC "
else
	strAsgSQL = strAsgSQL & "ORDER BY user_last_access DESC"
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
			Response.Write(buildTableContNoRecord(1, "search"))
					
		' Else show general no record information
		else
	
			' No current record			
			Response.Write(buildTableContNoRecord(1, "standard"))
					
		end If
				
	else
			
		objAsgRs.PageSize = RecordsPerPage
		objAsgRs.AbsolutePage = page
		intAsgCounter = (RecordsPerPage * (page-1))
				
		for loopAdvDataSorting = 1 to RecordsPerPage
					
			if not objAsgRs.EOF then			
				intAsgCounter = intAsgCounter + 1
				lngAsgReportUserId = Clng(objAsgRs("user_id"))
				strAsgReportCountry = objAsgRs("user_country")
				strAsgReportOs = objAsgRs("user_os")
				strAsgReportBrowser = objAsgRs("user_browser")
				strAsgReportReferer = objAsgRs("user_referer_url")
				strAsgReportSengine = objAsgRs("user_search_engine")
	
				' Write an anchor
				Response.Write(vbCrLf & "<div id=""user" & lngAsgReportUserId & """></div>")
		
	%>		  
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr class="treport_row">
	<td width="25%" class="treport_title" style="text-align: right;"><%= TXT_user_system %></td>
	<td class="treport_col">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>def/browser.asp?icon=<%= strAsgReportBrowser %>" alt="<%= strAsgReportBrowser %>" align="middle" />
		<%= searchTerms(strAsgReportBrowser, "user_browser", asgSearchfor, asgSearchIn) %> &middot;
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>def/os.asp?icon=<%= strAsgReportOs %>" alt="<%= strAsgReportOs %>" align="middle" />
		<%	
			Response.Write("" &_
			searchTerms(strAsgReportOs, "user_os", asgSearchfor, asgSearchIn) & " &middot; " &_
			searchTerms(objAsgRs("user_reso"), "user_reso", asgSearchfor, asgSearchIn) & " &middot; " &_
			searchTerms(objAsgRs("user_color"), "user_color", asgSearchfor, asgSearchIn) & "&nbsp;" & TXT_Bit) 
		%>
	</td>
  </tr>
  <tr class="treport_row">
	<td class="treport_title" style="text-align: right;"><%= TXT_IP %></td>
	<td class="treport_col"> 
<%
	if appAsgTrackCountry AND Len(strAsgReportCountry) > 0 AND strAsgReportCountry <> ASG_UNKNOWN then
		Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "flags/" & objAsgRs("user_country_2chr") & ".png"" alt=""" & strAsgReportCountry & """ align=""middle"" /> " &_
			"[ " & searchTerms(strAsgReportCountry, "user_country", asgSearchfor, asgSearchIn) & " ] ")
	end if

	Response.Write(objAsgRs("user_ip"))
%>
	</td>
  </tr>
<%
	if strAsgReportReferer <> ASG_UNKNOWN AND Len(strAsgReportReferer) > 0 then
%>
  <tr class="treport_row">
	<td class="treport_title" style="text-align: right;"><%= TXT_referer %></td>
	<td class="treport_col"><%
		if strAsgReportReferer = ASG_OWNSERVER then
			Response.Write(ASG_OWNSERVER)
		else
			Response.Write("<a href=""" & strAsgReportReferer & """ title=""" & TXT_gotoPage & """>" &_
				stripValueTooLong(strAsgReportReferer, 65, 30, 30) & "</a>")
		end if
	%></td>
  </tr>
<%
  	end if

	if objAsgRs("user_last_page") <> ASG_UNKNOWN AND Len(objAsgRs("user_last_page")) > 0 then
%>
  <tr class="treport_row">
	<td class="treport_title" style="text-align: right;"><%= TXT_currentpage %></td>
	<td class="treport_col">
		<a href="<% = objAsgRs("user_last_page") %>" title="<%= TXT_gotoPage %>">
		<%= stripValueTooLong(objAsgRs("user_last_page"), 65, 30, 30) %></a>
	</td>
  </tr>
<%
  	end if
		  	
	if Len(strAsgReportSengine) > 0 then
%>
  <tr class="treport_row">
	<td class="treport_title" style="text-align: right;"><%= TXT_search_engine %></td>
	<td class="treport_col">
	<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>def/engine.asp?icon=<%= strAsgReportSengine %>" alt="<%= strAsgReportSengine %>" align="middle" />
	[ <%= strAsgReportSengine  %> ] <%= objAsgRs("user_search_query") %></td>
  </tr>
<%
  	end if
	
	if ASG_VISITOR_ADVANCED then
		
		aryAsgReportVisitLength = convertSecondToTime(DateDiff("s", objAsgRs("user_first_access"), objAsgRs("user_last_access")))
		if aryAsgReportVisitLength(0) <> 0 OR aryAsgReportVisitLength(1) <> 0 OR aryAsgReportVisitLength(2) <> 0 then
			
			strAsgBuffer = TXT_visit_length_schema
			strAsgBuffer = Replace(strAsgBuffer, "$hours$", aryAsgReportVisitLength(2))
			strAsgBuffer = Replace(strAsgBuffer, "$minutes$", aryAsgReportVisitLength(1))
			strAsgBuffer = Replace(strAsgBuffer, "$seconds$", aryAsgReportVisitLength(0))
			strAsgBuffer = Replace(strAsgBuffer, "$hours_label$", TXT_hours)
			strAsgBuffer = Replace(strAsgBuffer, "$minutes_label$", TXT_minutes)
			strAsgBuffer = Replace(strAsgBuffer, "$seconds_label$", TXT_seconds)
%>
  <tr class="treport_row">
	<td class="treport_title" style="text-align: right;"><%= TXT_visit_length %></td>
	<td class="treport_col"><%= strAsgBuffer %></td>
  </tr>
<%
		end if
%>
  <tr class="treport_row">
	<td class="treport_title" style="text-align: right;"><%= TXT_user_agent %></td>
	<td class="treport_col"><%= objAsgRs("user_useragent") %></td>
  </tr>
<%
	end if

	' Get the total of hits
	intAsgRecordCountHits = Cint(objAsgRs("user_hits"))

	' Get the first visited page
	strAsgActiveRange = TXT_active_range_schema
	strAsgActiveRange = Replace(strAsgActiveRange, "$startDate$", formatDateTimeValue(objAsgRs("user_first_access"), "Date"))
	strAsgActiveRange = Replace(strAsgActiveRange, "$startTime$", formatDateTimeValue(objAsgRs("user_first_access"), "Time"))
	strAsgActiveRange = Replace(strAsgActiveRange, "$endDate$", formatDateTimeValue(objAsgRs("user_last_access"), "Date"))
	strAsgActiveRange = Replace(strAsgActiveRange, "$endTime$", formatDateTimeValue(objAsgRs("user_last_access"), "Time"))

%>
  <tr class="treport_row">
	<td class="treport_title" style="text-align: right;"><%= intAsgRecordCountHits & " " & TXT_pageviews %></td>
	<td class="treport_col"><% if Len(strAsgActiveRange) then Response.Write(strAsgActiveRange) %></td>
  </tr>
<%
		if Cint(strAsgDetails) = Cint(lngAsgReportUserId) OR blnAsgShowDetails then
				
			Dim objAsgRs2
			Set objAsgRs2 = Server.CreateObject("ADODB.Recordset")
				
			'Initialise SQL string to update values
			strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "detail " &_
				"WHERE detail_user_id = " & lngAsgReportUserId & " " &_
				"ORDER BY detail_date DESC "
					
%>
  <tr class="treport_rowdetails">
	<td class="treport_coldetails" style="text-align: center;" colspan="2">
		<!-- details -->
		<div id="tdetails">
			<div id="tdetails_bg">
			<table width="90%" align="center" cellpadding="1" cellspacing="1">
			  <tr>
				<td class="tdetails_title"><%= TXT_date %></td>
				<td class="tdetails_title"><%= TXT_page %></td>
			  </tr>
			  <% 
				
				' Set Rs properties
				if ASG_USE_MYSQL then
					objAsgRs2.CursorLocation = 3
				end if
				objAsgRs2.CursorType = 1
				objAsgRs2.LockType = 3
			
				objAsgRs2.Open strAsgSQL, objAsgConn
				do while not objAsgRs2.EOF
			  %>
			  <tr <%= buildTableContRollover("tdetails_row") %> >
				<td class="treport_col" style="text-align: center;"><%= formatDateTimeValue(objAsgRs2("detail_date"), "Date") & "&nbsp;" & TXT_time_at & "&nbsp;" & formatDateTimeValue(objAsgRs2("detail_date"), "Time") %></td>
				<td class="treport_col" style="text-align: left;">
				<%
				
				strAsgReportPage = objAsgRs2("detail_page_url")
				
				' Link the page
				Response.Write("<a href=""" & strAsgReportPage & """ title=""" & TXT_gotoPage & "&nbsp;" & strAsgReportPage & """>")				
				' Format the page text
				strAsgReportPage = stripValueTooLong(strAsgReportPage, 65, 30, 30)
				strAsgReportPage = searchTerms(strAsgReportPage, "detail_page_url", asgSearchfor, asgSearchIn)
				' Write the page
				Response.Write(strAsgReportPage & "</a>")				
					
				%>
				</td>
			  </tr>
			  <%
						
					objAsgRs2.MoveNext
				loop

				objAsgRs2.Close
				Set objAsgRs2 = Nothing

			%></table></div>
			</div>
		<!-- / details -->
	</td>
  </tr>
<%
		else
				
%>
  <tr class="treport_row">
	<td class="treport_col" style="text-align: right;" colspan="2"><%
		Response.Write(vbCrlf & "<a href=""asg_visitor.asp?page=" & Request.QueryString("page") & "&amp;details=" & lngAsgReportUserId & "#user" & lngAsgReportUserId & """ title=""" & TXT_details & """>" &_
			TXT_details & " <img src=""" & STR_ASG_SKIN_PATH_IMAGE & "arrow_small_dx.gif"" alt=""" & TXT_details & """ /></a>")
	%></td>
  </tr>
<%
		end if  ' details
%>
</table>
<%
			objAsgRs.MoveNext
			end if
		next
	
	end if	' .EOF

' 
'% ></table><%

' Advanced data sorting
strLayerAdvDataSorting = buildLayerAdvDataSorting()

objAsgRs.Close

' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing
				
' :: Close tlayout :: MENUSECTION_VisitorDetails
Response.Write(builTableTlayout("", "close", ""))


Response.Write(vbCrLf & "<br />")


' :: Open tlayout :: BARLABEL_DataView
Response.Write(builTableTlayout("rowNavy", "open", buildSwapDisplay("rowNavy", BARLABEL_DataView)))

	' Advanced data sorting layer
	Response.Write(strLayerAdvDataSorting)

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
'Response.Write(builTableTlayout("x-rowSearch", "open", buildSwapDisplay("rowSearch", BARLABEL_DataSearch)))

	' Row - Layers search
'	Response.Write(buildLayerSearch("", ""))
				
' :: Close tlayout :: BARLABEL_DataSearch
'Response.Write(builTableTlayout("", "close", ""))

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