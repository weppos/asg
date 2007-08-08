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

' Other variables
Dim intAsgCounter	' 
Dim blnAsgNoData
Dim loopAsgSerp		' 
Dim blnAsgDetails	' 
Dim aryAsgDetails	' 
Dim strAsgDetails
Dim blnAsgShowNopage

' -1 serp
if Request.QueryString("shownopage") = 1 then 
	blnAsgShowNopage = true
else
	blnAsgShowNopage = false
end if

' Serp details
strAsgDetails = Request.QueryString("serp")
if IsNumeric(strAsgDetails) AND Len(strAsgDetails) > 0 then
	strAsgDetails = Cint(strAsgDetails)
Else
	strAsgDetails = ""
End If

strAsgSortByFld = "query_visits"
strAsgSortOrder = "DESC"

'Set to false elaboration variables
blnAsgNoData = false
blnAsgDetails = false

' Call advanced data sorting configuration and variables
Call dimAdvDataSorting

' Carry on with the details
if strAsgDetails <> "" then
	
	strAsgDetails = Cint(strAsgDetails)
	aryAsgDetails = Split(strAsgDetails, "|")
	blnAsgDetails = true

else

	' Get SERP data from db
	if strAsgMode = "month" then 
		strAsgSQL = "SELECT DISTINCT query_serp_page " &_
			"FROM " & ASG_TABLE_PREFIX & "query " &_
			"WHERE query_period = '" & strAsgPeriod & "' "
		' Hide unknown serp
		if not blnAsgShowNopage then
			strAsgSQL = strAsgSQL & "AND query_serp_page <> -1 "
		end if
	elseif strAsgMode = "all" then 
		strAsgSQL = "SELECT DISTINCT query_serp_page " &_
		"FROM " & ASG_TABLE_PREFIX & "query "
		' Hide unknown serp
		if not blnAsgShowNopage then
			strAsgSQL = strAsgSQL & "WHERE query_serp_page <> -1 "
		end if
	End If
	' Order record the following field 
	strAsgSQL = strAsgSQL & "ORDER BY query_serp_page "

	' Get data from db and store information into a string
	objAsgRs.Open strAsgSQL, objAsgConn
	if not objAsgRs.EOF then
		Do While Not objAsgRs.EOF
			' Store data in a variable
			if Trim(Len(objAsgRs("query_serp_page"))) > 0 then strAsgDetails = strAsgDetails & objAsgRs("query_serp_page") & "|"
			objAsgRs.MoveNext
		Loop
		' Check if the recorset is .EOF
		If Not Trim(Len(strAsgDetails)) > 0 Then blnAsgNoData = true
	' .EOF
	else
		blnAsgNoData = true
	end if
	objAsgRs.Close

	if not blnAsgNoData then
		' Clean up the variable removing empty strings
		strAsgDetails = Left(strAsgDetails, Len(strAsgDetails) - 1)
		' Split the values
		aryAsgDetails = Split(strAsgDetails, "|")
		if IsNumeric(Request.QueryString("perpage")) AND Len(Request.QueryString("perpage")) > 0 then
			RecordsPerPage = Clng(Request.QueryString("perpage"))
		else
			RecordsPerPage = 10
		end if
	end if

' / Details
End If

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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Marketing & " &raquo;</span> " & MENUSECTION_Serp %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_Serp
Response.Write(builTableTlayout("", "open", MENUSECTION_SearchEngines))

' The recordset wasn't empty
if not blnAsgNoData then

	' Show TOP 5 of each serp
	for loopAsgSerp = 0 to Ubound(aryAsgDetails)
		if Len(Trim(aryAsgDetails(loopAsgSerp))) > 0 then

%>		
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
	<td width="5%"  class="treport_title">&nbsp;</td>
	<td width="58%"  class="treport_title"><% if not blnAsgDetails then Response.Write(TXT_pagetop & "&nbsp;") End If : Response.Write(TXT_Queries & "&nbsp;" & TXT_On  & "&nbsp;" & aryAsgDetails(loopAsgSerp) & "&deg;&nbsp;" & TXT_Page) %></td>
	<td width="25%"  class="treport_title"><%= TXT_search_engine %></td>
	<td width="25%"  class="treport_title"><%= TXT_visits %></td>
  </tr>
<%

'Initialise SQL string to select data
if strAsgMode = "month" then 
	strAsgSQL = "SELECT query_keyphrase, engine_name, query_serp_page, query_visits " &_
		"FROM " & ASG_TABLE_PREFIX & "query " &_
		"WHERE query_period = '" & strAsgPeriod & "' AND query_serp_page = " & aryAsgDetails(loopAsgSerp)
elseif strAsgMode = "all" then 
	strAsgSQL = "SELECT query_keyphrase, engine_name, query_serp_page, query_visits " &_
		"FROM " & ASG_TABLE_PREFIX & "query " &_
		"WHERE query_serp_page = " & aryAsgDetails(loopAsgSerp)
end if
' Call the function to search into the database if there are enought information to do that
strAsgSQL = searchFor(strAsgSQL, false)
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

		' 
		if blnAsgDetails then
			objAsgRs.PageSize = RecordsPerPage
			objAsgRs.AbsolutePage = page
			intAsgCounter = (RecordsPerPage * (page - 1))
		else
			intAsgCounter = 0
		end if
			
		objAsgRs.PageSize = RecordsPerPage
		objAsgRs.AbsolutePage = page
				
		for loopAdvDataSorting = 1 to RecordsPerPage
					
			if not objAsgRs.EOF then			
			intAsgCounter = intAsgCounter + 1
					
%>		  
  <tr <%= buildTableContRollover("treport_row") %> >
	<td class="treport_col" style="text-align: center;"><%= intAsgCounter %></td>
	<td class="treport_col" style="text-align: left;"><% if Len(objAsgRs("query_serp_page")) > 0 then Response.Write("&nbsp;<span class=""notetext"">[" & objAsgRs("query_serp_page") & "]</span>") %>&nbsp;<%= ShareWords(searchTerms(objAsgRs("query_keyphrase"), "query_keyphrase"), 40) %></td>
	<td class="treport_col" style="text-align: left;"><img src="<%= STR_ASG_SKIN_PATH_IMAGE %>def/engine.asp?icon=<%= objAsgRs("engine_name") %>" alt="<%= objAsgRs("engine_name") %>" align="middle" /> <%= searchTerms(objAsgRs("engine_name"), "engine_name") %></td>
	<td class="treport_col" style="text-align: right;"><%= objAsgRs("query_visits") %></td>
  </tr>
<%
	
			objAsgRs.MoveNext
			end if
		next
	end If
		
%>
  <tr <%= buildTableContRollover("treport_row") %> >
	<td class="treport_col" style="text-align: center;" colspan="4"><%
	
	' Create a tmp querystring without new values
	strAsgAppend = appendToQuerystring("serp")

	' Create the link to browse de reports...
	Response.Write(vbCrlf & "<a href=""serp.asp?" & strAsgAppend)
	' ...add the querystring values if it's necessary...
	if not blnAsgDetails then Response.Write("&amp;serp=" & aryAsgDetails(loopAsgSerp))
	'...finish the link and print data...
	if blnAsgDetails then 
	Response.Write(""" title=""" & TXT_pagetop & "&nbsp;" & TXT_Queries & """>")
	Response.Write(TXT_pagetop & "&nbsp;" & TXT_Queries)
	else
	Response.Write(""" title=""" & TXT_details & """>")
	Response.Write(TXT_details)
	end if
	'
	Response.Write(vbCrlf & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "arrow_small_dx.gif"" alt=""" & TXT_details & """ align=""middle"" border=""0"" /></a>")

	%>
	</td>
  </tr>
<%

%></table><%
' <br />

' Advanced data sorting
strLayerAdvDataSorting = buildLayerAdvDataSorting()

objAsgRs.Close

' / SERP values
End If
' / SERP Loop
Next


' .EOF
Else

	' Layout
	Response.Write(vbCrLf & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"">")
	Response.Write(vbCrLf & "<tr>")
	Response.Write(vbCrLf & "<td class=""treport_title"">&nbsp;</td>")
	Response.Write(vbCrLf & "</tr>")
		
		' If it is a search query then show no results advise
		if Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 then

			' No current record for search terms		
			Response.Write(buildTableContNoRecord(1, "search"))
					
		' Else show general no record information
		Else
	
			' No current record			
			Response.Write(buildTableContNoRecord(1, "standard"))
					
		End If

	Response.Write(vbCrLf & "</table>")

' / .EOF
End If

' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing
				
' :: Close tlayout :: MENUSECTION_SearchQueries
Response.Write(builTableTlayout("", "close", ""))


Response.Write(vbCrLf & "<br />")


' :: Open tlayout :: BARLABEL_DataView
Response.Write(builTableTlayout("rowNavy", "open", buildSwapDisplay("rowNavy", BARLABEL_DataView)))
			
	' Open the Navy form
	Response.Write(buildLayerForm("open"))
	
	' If the details are not empty show the advanced data sorting line
	if blnAsgDetails then
	
		' Advanced data sorting layer
		Response.Write(strLayerAdvDataSorting)
			
	end If
			
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
	Response.Write(buildLayerSearch("", "engine_name|query_keyphrase"))
				
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