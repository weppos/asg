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


Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing
Response.Redirect("asg_visitor.asp")

'-----------------------------------------------------------------------------------------
' Query Daily table to get hits and visits data and information	
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	
' Comment:			
'-----------------------------------------------------------------------------------------
Function getDailyValue(aryValue)
	
	' Build the query
	strAsgSQL = "SELECT Hits, Visits FROM " & ASG_TABLE_PREFIX & "daily WHERE "
	if ASG_USE_MYSQL then
		strAsgSQL = strAsgSQL & "Data = '" & dtmAsgDateValue(aryValue) & "'"
	else
		strAsgSQL = strAsgSQL & "Data = #" & dtmAsgDateValue(aryValue) & "#"
	end if
	
	' Open Rs
	objAsgRs.Open strAsgSQL, objAsgConn
		' If there are tracking information then valorize the variables
		if not objAsgRs.EOF then
			intAsgVisitsValue(aryValue) = objAsgRs("Visits")
			intAsgHitsValue(aryValue) = objAsgRs("Hits")
		' If the Rs is empty then set the variables to 0
		else
			intAsgVisitsValue(aryValue) = 0
			intAsgHitsValue(aryValue) = 0
		end If
	' Close Rs
	objAsgRs.Close

End Function

Function getMonthlyValue(aryValue)
	
	' Build the query
	strAsgSQL = "SELECT SUM(Hits) AS SumHits, SUM(Visits) AS SumVisits FROM " & ASG_TABLE_PREFIX & "daily WHERE Mese = '" & dtmAsgDateValue(aryValue) & "' GROUP BY Mese "
	
	' Open Rs
	objAsgRs.Open strAsgSQL, objAsgConn
		' If there are tracking information then valorize the variables
		if not objAsgRs.EOF then
			intAsgVisitsValue(aryValue) = objAsgRs("SumVisits")
			intAsgHitsValue(aryValue) = objAsgRs("SumHits")
		' If the Rs is empty then set the variables to 0
		else
			intAsgVisitsValue(aryValue) = 0
			intAsgHitsValue(aryValue) = 0
		end If
	' Close Rs
	objAsgRs.Close

End Function

'Faccio prima a richiamare la stessa funzione in config
'e passargli i nuovi parametri!
Call formatTimeZone(dtmAsgNow, appAsgTimeZone)


' Declare variables
Dim dtmAsgDateValue(4)		'Holds the dates from the most recent to the oldest
							'1 today - 2 yesterday - 3 this month - 4 last month
Dim intAsgHitsValue(4)		'Holds hits data and information from the most recent to the oldest
Dim intAsgVisitsValue(4)	'Holds visits data and information from the most recent to the oldest
Dim intAsgUsersOnline		' Holds the number of online users on the site


' Get the date of today
dtmAsgDateValue(1) = dtmAsgDate
' Get the date of yesterday and convert in a proper date format
dtmAsgDateValue(2) = DateAdd("d", -1, dtmAsgDate) : dtmAsgDateValue(2) = Year(dtmAsgDateValue(2)) & "/" & Month(dtmAsgDateValue(2)) & "/" & Day(dtmAsgDateValue(2))
' Get the value of the current month
dtmAsgDateValue(3) = Cstr(Month(dtmAsgDate))				'Convert to string to be able to use them in the query
' Get the value of the last month
dtmAsgDateValue(4) = Cstr(Month(DateAdd("m", -1, dtmAsgDate)))	'Convert to string to be able to use them in the query
' Format values returnin a 2 chr format
If Len(dtmAsgDateValue(3)) < 2 Then dtmAsgDateValue(3) = "0" & dtmAsgDateValue(3)
If Len(dtmAsgDateValue(4)) < 2 Then dtmAsgDateValue(4) = "0" & dtmAsgDateValue(4)

' Add the year value
' Check the "last month" value to add the right year at the end of the string
' We are on january. Put the last year at the end
if Cstr(dtmAsgDateValue(3)) = "01" then
	dtmAsgDateValue(4) = dtmAsgDateValue(4) & "-" & (dtmAsgYear - 1)
' Month anf last month are on the same year
else
	dtmAsgDateValue(4) = dtmAsgDateValue(4) & "-" & dtmAsgYear
end if
dtmAsgDateValue(3) = dtmAsgDateValue(3) & "-" & dtmAsgYear

'Today report
Call getDailyValue(1)

'Yesterday report
Call getDailyValue(2)

'This month
Call getMonthlyValue(3)

'Last month
Call getMonthlyValue(4)

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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Main & " &raquo;</span> " & MENUSECTION_Summary %></div>
		<div id="layout_content">

<%
			  
' TableBar			
'Call buildTableBar(MENUSECTION_Summary, MENUGROUP_Main)
	
' 
'Response.Write(vbCrLf & "<div class=""table_layout"">")
		  
%>
		<table width="95%" border="0" align="center" cellpadding="1" cellspacing="1">
		<tr valign="top"><td width="48%">

		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
          <tr bgcolor="<%= STR_ASG_SKIN_TABLE_TITLE_BGCOLOUR %>"class="normaltitle">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_TITLE_BGIMAGE %>" colspan="2" align="center" height="16"><%= UCase(TXT_BoxTitle_TrafficSummary) %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td width="70%"><span class="notetext"><%= TXT_pageviews & "&nbsp;" & TXT_Today %></span></td>
            <td width="30%"><%= intAsgHitsValue(1) %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><span class="notetext"><%= TXT_visits & "&nbsp;" & TXT_Today %></span></td>
            <td><%= intAsgVisitsValue(1) %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><%= TXT_pageviews & "&nbsp;" & TXT_Yesterday %></td>
            <td><%= intAsgHitsValue(2) %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><%= TXT_visits & "&nbsp;" & TXT_Yesterday %></td>
            <td><%= intAsgVisitsValue(2) %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><span class="notetext"><%= TXT_pageviews & "&nbsp;" & TXT_ThisMonth %></span></td>
            <td><%= intAsgHitsValue(3) %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><span class="notetext"><%= TXT_visits & "&nbsp;" & TXT_ThisMonth %></span></td>
            <td><%= intAsgVisitsValue(3) %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><%= TXT_pageviews & "&nbsp;" & TXT_LastMonth %></td>
            <td><%= intAsgHitsValue(4) %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><%= TXT_visits & "&nbsp;" & TXT_LastMonth %></td>
            <td><%= intAsgVisitsValue(4) %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><%= TXT_BeginningOfStats %></td>
            <td><%= Day(appAsgProgramSetup) & "&nbsp;" & aryAsgMonth(1,Month(appAsgProgramSetup)) & "&nbsp;" & Year(appAsgProgramSetup)  %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><span class="notetext"><%= MENUSECTION_OnlineUsers %></span></td>
            <td><%= "" %></td>
          </tr>
		  <%
			  
			'// Row - End table spacer			
			Call buildTableContEndSpacer(2)
			  
		  %>
		</table><br />
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
          <tr bgcolor="<%= STR_ASG_SKIN_TABLE_TITLE_BGCOLOUR %>" class="normaltitle">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_TITLE_BGIMAGE %>" colspan="2" align="center" height="16"><%= UCase(TXT_BoxTitle_TrafficSummary_Year) %></td>
          </tr>
<%

' Get yearly traffic information from database
strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "counter ORDER BY counter_periody DESC "
objAsgRs.Open strAsgSQL, objAsgConn

if objAsgRs.EOF then

%>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td width="70%"><span class="notetext"><%= TXT_pageviews & "&nbsp;" & dtmAsgYear %></span></td>
            <td width="30%"><%= appAsgStartHits %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><span class="notetext"><%= TXT_visits & "&nbsp;" & dtmAsgYear %></span></td>
            <td><%= appAsgStartVisits %></td>
          </tr>
<%

else
	do while not objAsgRs.EOF
%>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td width="70%"><%

			' format current year values
			If Cint(objAsgRs("counter_periody")) = Cint(dtmAsgYear) then Response.Write("<span class=""notetext"">")
			Response.Write(TXT_pageviews & "&nbsp;" & objAsgRs("counter_periody")) 
			If Cint(objAsgRs("counter_periody")) = Cint(dtmAsgYear) then Response.Write("</span>")
			
			%></td>
            <td width="30%"><% Response.Write(objAsgRs("counter_hits") + appAsgStartHits) %></td>
          </tr>
		  <tr <%= buildTableContRollover("table_cont_row") %>>
            <td><%
			
			' format current year values
			If Cint(objAsgRs("counter_periody")) = Cint(dtmAsgYear) then Response.Write("<span class=""notetext"">")
			Response.Write(TXT_visits & "&nbsp;" & objAsgRs("counter_periody")) 
			If Cint(objAsgRs("counter_periody")) = Cint(dtmAsgYear) then Response.Write("</span>")
			
			%></td>
            <td><% Response.Write(objAsgRs("counter_visits") + appAsgStartVisits) %></td>
          </tr>
<%
	objAsgRs.MoveNext
	Loop

end if
	
objAsgRs.Close

Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

					
			'// Row - End table spacer			
			Call buildTableContEndSpacer(2)
	
%>		  
		</table><br />
		
		</td><td width="4%" >
		</td><td width="48%">
		
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= STR_ASG_SKIN_TABLE_TITLE_BGCOLOUR %>" align="center" class="normaltitle">
            <td colspan="2" align="center" background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_TITLE_BGIMAGE %>" height="16"><%= UCase(TXT_BoxTitle_TrafficSummary_Average) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>" width="70%"><span class="notetext"><%= TXT_pageviews & "&nbsp;" & TXT_Today & "&nbsp;" & TXT_PerHour %></span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>" width="30%"><%= MediaGiorno(intAsgHitsValue(1), 0, 1) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><span class="notetext"><%= TXT_visits & "&nbsp;" & TXT_Today  & "&nbsp;" & TXT_PerHour %></span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaGiorno(intAsgVisitsValue(1), 0, 1) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= TXT_pageviews & "&nbsp;" & TXT_Yesterday & "&nbsp;" & TXT_PerHour %></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaGiorno(intAsgHitsValue(2), 0, 2) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= TXT_visits & "&nbsp;" & TXT_Yesterday & "&nbsp;" & TXT_PerHour %></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaGiorno(intAsgVisitsValue(2), 0, 2) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><span class="notetext"><%= TXT_pageviews & "&nbsp;" & TXT_ThisMonth & "&nbsp;" & TXT_PerHour %></span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaMese(intAsgHitsValue(3), 1, 1) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><span class="notetext"><%= TXT_visits & "&nbsp;" & TXT_ThisMonth & "&nbsp;" & TXT_PerHour %></span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaMese(intAsgVisitsValue(3), 1, 1) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= TXT_pageviews & "&nbsp;" & TXT_LastMonth & "&nbsp;" & TXT_PerHour %></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaMese(intAsgHitsValue(4), 1, 2) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= TXT_visits & "&nbsp;" & TXT_LastMonth & "&nbsp;" & TXT_PerHour %></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaMese(intAsgVisitsValue(4), 1, 2) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><span class="notetext"><%= TXT_pageviews & "&nbsp;" & TXT_ThisMonth & "&nbsp;" & TXT_PerDay %></span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaMese(intAsgHitsValue(3), 2, 1) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><span class="notetext"><%= TXT_visits & "&nbsp;" & TXT_ThisMonth & "&nbsp;" & TXT_PerDay %></span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaMese(intAsgVisitsValue(3), 2, 1) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= TXT_pageviews & "&nbsp;" & TXT_LastMonth & "&nbsp;" & TXT_PerDay %></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaMese(intAsgHitsValue(4), 2, 2) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= TXT_visits & "&nbsp;" & TXT_LastMonth & "&nbsp;" & TXT_PerDay %></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= MediaMese(intAsgVisitsValue(4), 2, 2) %></td>
          </tr>
		  <%
					
			'// Row - End table spacer			
			Call buildTableContEndSpacer(2)
	
		  %>
		</table><br />
		
		</td></tr>
		<tr valign="top"><td width="100%" colspan="3"><br />

		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
          <tr bgcolor="<%= STR_ASG_SKIN_TABLE_TITLE_BGCOLOUR %>"class="normaltitle">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_TITLE_BGIMAGE %>" colspan="4" align="center" height="16"><%= UCase(TXT_ServerInfo) %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>" width="20%"><span class="notetext"><%= TXT_IISversion %></span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>" width="30%"><%= Request.ServerVariables("SERVER_SOFTWARE") %></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>" width="20%"><span class="notetext"><%= TXT_ServerName %></span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>" width="30%"><%= Request.ServerVariables("SERVER_NAME") %></td>
          </tr>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><span class="notetext"><%= TXT_ProtocolVersion %></span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= Request.ServerVariables("SERVER_PROTOCOL") %></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><span class="notetext">VBScript Engine</span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%= getScriptEngineInfo() %></td>
          </tr>
		  <%

			'Link a completo se loggato
			If Session("asgLogin") = "Logged" Then

		  %>
		  <tr class="normaltext" bgcolor="<%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>">
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><span class="notetext"><%= TXT_YourIpIs %></span></td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>"><%

				'Filter IP
				'// Link PopUp
				Response.Write(vbCrLf & "						<a href=""JavaScript:openWin('popup_filter_ip.asp?IP=" & Request.ServerVariables("REMOTE_ADDR") & "','Filter','toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=200')"" title=""" & TXT_Filtered_IPs & """>" & Request.ServerVariables("REMOTE_ADDR"))
				'// Chiudi Link PopUp
				Response.Write("</a>") 

				'Icona Filter IP
				Call ShowIconFilterIp(Request.ServerVariables("REMOTE_ADDR"))

			%>
			</td>
            <td background="<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>" colspan="2" align="right">
				<a href="sysinfo.asp" title="<%= MENUSECTION_ServerVariables_descr %>"><%= TXT_FullVersion %> <img src="<%= STR_ASG_SKIN_PATH_IMAGE %>arrow_small_dx.gif" alt="<%= MENUSECTION_ServerVariables_descr %>" align="middle" border="0" /></a>
			</td>
          </tr>
		  <%

			End If

			'// Row - End table spacer			
			Call buildTableContEndSpacer(4)
	
		  %>
		</table><br />

		</td></tr>
		</table>

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