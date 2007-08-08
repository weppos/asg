<% @LANGUAGE="VBSCRIPT" %>
<!--#include file="w2k3_config.asp" -->
<!--#include file="lib/functions_count.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'// WARNING! Program protection.
'	Changing default values may allow users to access the page.
Call checkPermission("False", "False", "False", appAsgSecurity)


Dim blnAsgIsIncluded		' Set to true if the PC is included into monitoring action
Dim strAsgTmpIPs

' Exclude the pc
if Request.QueryString("act") = "excludepc" Then
	Response.Cookies(ASG_COOKIE_PREFIX & "exitcount") = "excludepc"
	Response.Cookies(ASG_COOKIE_PREFIX & "exitcount").Expires = dateAdd("yyyy", 1, date)

' Include the pc
elseIf Request.QueryString("act") = "includepc" Then
	Response.Cookies(ASG_COOKIE_PREFIX & "exitcount") = ""
	Response.Cookies(ASG_COOKIE_PREFIX & "exitcount").Expires = dateAdd("yyyy", -1, date)

' Update settings
elseIf Request.QueryString("act") = "upd" Then

	appAsgFilteredIPs = Trim(Request.Form("filteredIP"))
	' Trim spaces
	appAsgFilteredIPs = Replace(appAsgFilteredIPs, " ", "")

	' Initialise SQL string to update the table
	strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "config SET conf_filtered_ips = '" & appAsgFilteredIPs & "'"
	objAsgConn.Execute(strAsgSQL)
	
	' If application variables are used update them
	if blnApplicationConfig then
						
		' Update IPs variable
		Application(ASG_APPLICATION_PREFIX & "FilteredIPs") = appAsgFilteredIPs

		' Refresh application
		Application(ASG_APPLICATION_PREFIX & "Config") = False
	
	end If
	
	' Reset objects
	' Set objAsgRs = Nothing
	' objAsgConn.Close
	' Set objAsgConn = Nothing
	
	' Redirect to refresh new values
	' Response.Redirect("settings_exitcount.asp")

end If

' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

' 
strAsgTmpIPs = appAsgFilteredIPs

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
<script language="JavaScript" type="text/javascript" src="tip_idea.js.asp"></script>

<div align="center">
	<div id="layout">
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Administration & " &raquo; " & MENUSECTION_General & " &raquo;</span> " & MENUSECTION_TrackingExclusion %></div>
		<div id="layout_content">

<form action="settings_exitcount.asp?act=upd" name="frmSettings" method="post">
<%

' :: Open tlayout :: MENUSECTION_TrackingExclusion
Response.Write(builTableTlayout("", "open", MENUSECTION_TrackingExclusion))
	
	
	' Update completed
	if Len(Request.QueryString("act")) > 0 then Response.Write("<p style=""text-align: center;""><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Update_Completed & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Update_Completed & "</p></p>")
	
	' Check current status
	if Request.Cookies(ASG_COOKIE_PREFIX & "exitcount") = "excludepc" then
		blnAsgIsIncluded = false
	else
		blnAsgIsIncluded = true
	end if

	' Filter by IP
	strAsgTmpLayer = "<p>"
	if exitCountByIP(Request.ServerVariables("REMOTE_ADDR")) then
		strAsgTmpLayer = strAsgTmpLayer & Replace(TXT_Exclmex, "$v1$", "<span class=""notetext"">" & TXT_Excluded & "</span>")
	else
		strAsgTmpLayer = strAsgTmpLayer & Replace(TXT_Exclmex, "$v1$", "<span class=""notetext"">" & TXT_Included & "</span>")
	end if
	strAsgTmpLayer = strAsgTmpLayer & "</p>"
	strAsgTmpLayer = strAsgTmpLayer & "<p>" & TXT_Filtered_IPs & ": <input type=""text"" name=""filteredIP"" value=""" & strAsgTmpIPs & """ size=""60"" maxlength=""200"" />&nbsp;"
	strAsgTmpLayer = strAsgTmpLayer & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/tip.png"" alt=""" & TXT_Tip & """ border=""0"" align=""middle"" onmouseover=""stm(Idea[0],Style[7])"" onmouseout=""htm()"" />"
	strAsgTmpLayer = strAsgTmpLayer & "</p>"
		
		' :: Create the layer ::
		Response.Write(buildLayer("layerFilterIp", TXT_ExitByIP, "", strAsgTmpLayer))
	
	' Filter by Cookie
	strAsgTmpLayer = "<p>"
	if blnAsgIsIncluded then
		strAsgTmpLayer = strAsgTmpLayer & Replace(TXT_Exclmex, "$v1$", "<span class=""notetext"">" & TXT_Included & "</span>")
	else
		strAsgTmpLayer = strAsgTmpLayer & Replace(TXT_Exclmex, "$v1$", "<span class=""notetext"">" & TXT_Excluded & "</span>")
	end if
	strAsgTmpLayer = strAsgTmpLayer & "</p><p>"
	' 
	if blnAsgIsIncluded then
		strAsgTmpLayer = strAsgTmpLayer & "<a href=""settings_exitcount.asp?act=excludepc"" title=""" & Replace(TXT_Exclpc, "$v1$", TXT_Exclude) & """><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/exclude.png"" alt=""" & Replace(TXT_Exclpc, "$v1$", TXT_Exclude) & """ border=""0"" align=""middle"" />&nbsp;" & Replace(TXT_Exclpc, "$v1$", TXT_Exclude) & "</a>"
	else
		strAsgTmpLayer = strAsgTmpLayer & "<a href=""settings_exitcount.asp?act=includepc"" title=""" & Replace(TXT_Exclpc, "$v1$", TXT_Include) & """><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/include.png"" alt=""" & Replace(TXT_Exclpc, "$v1$", TXT_Include) & """ border=""0"" align=""middle"" />&nbsp;" & Replace(TXT_Exclpc, "$v1$", TXT_Include) & "</a>"
	end if
	strAsgTmpLayer = strAsgTmpLayer & "</p>"
		
		' :: Create the layer ::
		Response.Write(buildLayer("layerFilterCookie", TXT_ExitByCookie, "", strAsgTmpLayer))
	
	' Submit form area
	Response.Write("<div class=""submitarea""><input type=""submit"" name=""settings"" value=""" & TXT_Update & """ /></div>")
	
' :: Close tlayout :: MENUSECTION_TrackingExclusion
Response.Write(builTableTlayout("", "close", ""))

%>
</form>

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