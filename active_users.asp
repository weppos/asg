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
Call checkPermission("True", "True", "False", appAsgSecurity)

' Format date time information
Call formatTimeZone(dtmAsgNow, appAsgTimeZone)

Dim intAsgUsersOnline		' Holds the number of online users on the site
Dim intAsgCounter
Const INT_ASG_MINUTES = 10	' Holds the number of time range - minutes - 
							' to consider an user as active
							
intAsgCounter = 0

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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Main & " &raquo;</span> " & MENUSECTION_ActiveUsers %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_OnlineUsers
Response.Write(builTableTlayout("", "open", MENUSECTION_ActiveUsers))

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr>
	<td width="5%"  class="treport_title">&nbsp;</td>
	<td width="15%" class="treport_title"><%= TXT_IP %></td>
	<td width="15%" class="treport_title"><%= TXT_pageviews %></td>
	<td width="20%" class="treport_title"><%= TXT_lastAccess %></td>
	<td width="45%" class="treport_title">&nbsp;</td>
  </tr>
<%

Dim dtmOnlineTime
dtmOnlineTime = DateAdd("n", -INT_ASG_MINUTES, dtmAsgNow) : dtmOnlineTime = Year(dtmOnlineTime) & "/" & Month(dtmOnlineTime) & "/" & Day(dtmOnlineTime) & " " & Hour(dtmOnlineTime) & "." & Minute(dtmOnlineTime) & "." & Second(dtmOnlineTime)

'Initialise SQL string to count online users
if ASG_USE_MYSQL then
	strAsgSQL = "SELECT user_ip, user_last_access, user_hits, user_search_engine, user_search_query, user_referer_url, user_last_page  " &_
		"FROM " & ASG_TABLE_PREFIX & "user " &_
		"WHERE user_last_access > '" & dtmOnlineTime & "' " &_
		"ORDER BY user_last_access DESC"
else
	strAsgSQL = "SELECT user_ip, user_last_access, user_hits, user_search_engine, user_search_query, user_referer_url, user_last_page  " &_
		"FROM " & ASG_TABLE_PREFIX & "user " &_
		"WHERE user_last_access > '" & dtmOnlineTime & "' " &_
		"ORDER BY user_last_access DESC"
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
			Response.Write(buildTableContNoRecord(5, "search"))
					
		' Else show general no record information
		else
	
			' No current record			
			Response.Write(buildTableContNoRecord(5, Replace(TXT_Nodata_activeusers, "$var1$", INT_ASG_MINUTES)))
					
		end if
				
	else
			
		do while not objAsgRs.EOF
		intAsgCounter = intAsgCounter + 1
					
%>		  
  <tr <%= buildTableContRollover("treport_row") %> >
	<td class="treport_col" style="text-align: center;"><%= intAsgCounter %></td>
	<td class="treport_col" style="text-align: center;"><%= objAsgRs("user_ip") %></td>
	<td class="treport_col" style="text-align: center;"><%= Cint(objAsgRs("user_hits")) & "&nbsp;" & TXT_pageviews %></td>
	<td class="treport_col" style="text-align: center;"><%= formatDateTimeValue(objAsgRs("user_last_access"), "Date") & "&nbsp;" & TXT_time_at & "&nbsp;" & formatDateTimeValue(objAsgRs("user_last_access"), "Time") %></td>
	<td class="treport_col" style="text-align: left;"><%
	
	if Len(objAsgRs("user_last_page")) > 0 then
		Response.Write("<a href=""" & objAsgRs("user_last_page") & """ title=""" & objAsgRs("user_last_page") & """>")
		Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "menu/www.png"" alt=""" & TXT_page & """ align=""middle"" border=""0"" /></a> ")
	end if
	
	if Len(objAsgRs("user_referer_url")) > 0 then
		Response.Write("<a href=""" & objAsgRs("user_referer_url") & """ title=""" & objAsgRs("user_referer_url") & """>")
		Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "menu/referer.png"" alt=""" & TXT_referer & """ align=""middle"" border=""0"" /></a> ")
	end if
	
	if Len(objAsgRs("user_search_engine")) > 0 then
		Response.Write("<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "def/engine.asp?icon=" & objAsgRs("user_search_engine") & """ alt=""" & objAsgRs("user_search_engine") & """ align=""middle"" />")
		Response.Write(" [" & objAsgRs("user_search_query")  & "] ")
	end if
	
	%></td>
  </tr>
<%
				
		objAsgRs.MoveNext
		loop
	end if

%></table><%

objAsgRs.Close

' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing
				
' :: Close tlayout :: MENUSECTION_MENUSECTION_ActiveUsers
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
	Response.Write(buildLayerSearch("", "user_browser|user_os|user_color|user_reso"))
				
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