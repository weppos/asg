<% @LANGUAGE="VBSCRIPT" %>
<!--#include file="w2k3_config.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'// WARNING! Program protection.
'	Changing default values may allow users to access the page.
Call checkPermission("False", "False", "False", appAsgSecurity)


' Execute the update
if Request.QueryString("act") = "upd" then

	Dim strAsgPswNew
	Dim strAsgPswConf
	Dim blnAsgError
	Dim blnAsgPswIns
	
	blnAsgError = false
	blnAsgPswIns = false
	
	strAsgPswNew = Trim(Request.Form("newpsw"))
	strAsgPswConf = Trim(Request.Form("confpsw"))
	
	if IsNumeric(Request.Form("security")) then appAsgSecurity = CInt(Request.Form("security"))
	
	strAsgPswNew = CleanInput(strAsgPswNew)
	strAsgPswConf = CleanInput(strAsgPswConf)
	
	if "[]" & strAsgPswNew <> "[]" then
	
		if strAsgPswNew = strAsgPswConf then
			blnAsgPswIns = true
		else
			blnAsgError = true
		end if
	
	end If
	
	' No errors
	if blnAsgError = false then
	
		if blnAsgPswIns = true then
			strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "config SET conf_site_psw = '" & strAsgPswNew & "', conf_security_level = " & appAsgSecurity & ""
		else
			strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "config SET conf_security_level = " & appAsgSecurity & ""
		end if
		
		objAsgConn.Execute(strAsgSQL)
	
		' If application variables are used update them
		if blnApplicationConfig then
					
			' Update password variable
			if blnAsgPswIns then Application(ASG_APPLICATION_PREFIX & "SitePsw") = strAsgPswNew
			Application(ASG_APPLICATION_PREFIX & "Security") = CInt(appAsgSecurity)
			
			' Refresh application
			Application(ASG_APPLICATION_PREFIX & "Config") = false
		
		end if
		
		' Reset objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing
		
		' 
		Response.Redirect("settings_security.asp?act=updmex")
	
	end if

end if

' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Administration & " &raquo; " & MENUSECTION_General & " &raquo;</span> " & MENUSECTION_Security %></div>
		<div id="layout_content">

<form action="settings_security.asp?act=upd" name="frmSecurity" method="post">
<%

' :: Open tlayout :: MENUSECTION_TrackingExclusion
Response.Write(builTableTlayout("", "open", MENUSECTION_TrackingExclusion))
	
	
	' Update completed
	if Request.QueryString("act") = "updmex" then Response.Write("<p style=""text-align: center;""><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Update_Completed & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Update_Completed & "</p>")
	
	' Change password
	if blnAsgError then 
		strAsgTmpLayer = "<p>"
		strAsgTmpLayer = strAsgTmpLayer & "<span class=""errortext""><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/no.png"" alt=""" & TXT_Error_PasswordNotMatching & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Error_PasswordNotMatching & "</span>" 
		strAsgTmpLayer = strAsgTmpLayer & "</p>" 
	end if 
	
	strAsgTmpLayer = strAsgTmpLayer & "<table align=""center"" border=""0"" cellspacing=""1"" cellpadding=""3"">" 
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right"">" & TXT_Password_new & ":&nbsp;</td><td><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[5],Style[1])"" onmouseout=""htm()"" />&nbsp;<input type=""password"" name=""newpsw"" value="""" size=""20"" maxlength=""20"" /></td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right"">" & TXT_Password_confirm & ":&nbsp;</td><td><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[5],Style[1])"" onmouseout=""htm()"" />&nbsp;<input type=""password"" name=""confpsw"" value="""" size=""20"" maxlength=""20"" /></td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "</table>"

		' :: Create the layer ::
		Response.Write(buildLayer("layerPassword", TXT_Entry_password, "", strAsgTmpLayer))
	
	' Protection level
	strAsgTmpLayer = "<table align=""center"" border=""0"" cellspacing=""1"" cellpadding=""3"">" 
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right"">&nbsp;</td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td style=""text-align: left;""><input type=""radio"" name=""security"" value=""0"""
		if appAsgSecurity = 0 then strAsgTmpLayer = strAsgTmpLayer & " checked"
		 strAsgTmpLayer = strAsgTmpLayer & " />&nbsp;<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[2],Style[1])"" onmouseout=""htm()"" />&nbsp;" & TXT_Seclevel_None & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right"">" & TXT_Seclevel & ":&nbsp;</td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td style=""text-align: left;""><input type=""radio"" name=""security"" value=""1"""
		if appAsgSecurity = 1 then strAsgTmpLayer = strAsgTmpLayer & " checked"
		 strAsgTmpLayer = strAsgTmpLayer & " />&nbsp;<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[3],Style[1])"" onmouseout=""htm()"" />&nbsp;" & TXT_Seclevel_Limited & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right"">&nbsp;</td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td style=""text-align: left;""><input type=""radio"" name=""security"" value=""2"""
		if appAsgSecurity = 2 then strAsgTmpLayer = strAsgTmpLayer & " checked"
		 strAsgTmpLayer = strAsgTmpLayer & " />&nbsp;<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[4],Style[1])"" onmouseout=""htm()"" />&nbsp;" & TXT_Seclevel_Full & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "</table>" 
		
		' :: Create the layer ::
		Response.Write(buildLayer("layerLevel", TXT_StatsProtection, "", strAsgTmpLayer))
	
	' Submit form area
	Response.Write("<div class=""submitarea""><input type=""submit"" name=""settings"" value=""" & TXT_Update & """ /></div>")

' :: Open tlayout :: MENUSECTION_TrackingExclusion
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