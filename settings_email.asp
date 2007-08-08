<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<!--#include file="lib/utils.email.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'// WARNING! Program protection.
'	Changing default values may allow users to access the page.
Call checkPermission("False", "False", "False", appAsgSecurity)

Dim ii
Dim aryAsgMail

' Update
if Request.QueryString("act") = "upd" then

	' Get data from the form
	appAsgEmailAddress = Trim(Request.Form("address"))
	appAsgEmailServer = Trim(Request.Form("server"))
	appAsgEmailComponent = Trim(Request.Form("component"))

	' Initialise SQL string to update values
	strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "config SET " &_
	"conf_email_address = '" & appAsgEmailAddress & "', " &_
	"conf_email_component = '" & appAsgEmailComponent & "', " &_
	"conf_email_server = '" & appAsgEmailServer & "' "

	' Execute the update
	' Response.Write(strAsgSQL) : Response.End()
	objAsgConn.Execute(strAsgSQL)
	
	' If application variables are used update them
	if blnApplicationConfig then
						
		' Update application variables
		Application(ASG_APPLICATION_PREFIX & "EmailAddress") = appAsgEmailAddress
		Application(ASG_APPLICATION_PREFIX & "EmailServer") = appAsgEmailServer
		Application(ASG_APPLICATION_PREFIX & "EmailComponent") = appAsgEmailComponent

		' Refresh application
		Application(ASG_APPLICATION_PREFIX & "Config") = false
	
	end if
	
	' Reset objects
	Set objAsgRs = Nothing
	objAsgConn.Close
	Set objAsgConn = Nothing
	Response.Redirect("settings_email.asp?act=updmex")

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

<script type="text/javascript" language="javascript">
<!--

function testMailObject(mailobject) {
	openWin('email_test.asp?obj=' + mailobject, '', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=425');
}

// -->
</script>
</head>

<body>
<!--#include file="includes/header.asp" -->

<div align="center">
	<div id="layout">
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Administration & " &raquo; " & MENUSECTION_Email & " &raquo;</span> " & MENUSECTION_Config %></div>
		<div id="layout_content">

<form action="settings_email.asp?act=upd" name="frmSettings" method="post">
<%

' :: Open tlayout :: MENUSECTION_Config
Response.Write(builTableTlayout("", "open", MENUSECTION_Config))
	
	' Update completed
	if Request.QueryString("act") = "updmex" then Response.Write("<p style=""text-align: center;""><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Update_Completed & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Update_Completed & "</p>")

	' Get the email component list
	aryAsgMail = mail_components()
	 
	strAsgTmpLayer = "<table align=""center"" border=""0"" cellspacing=""1"" cellpadding=""3"">" &_
	"<tr><td align=""right"">" & TXT_Email_object & ":&nbsp;</td><td align=""left"">" &_
	"<select name=""component"">"
	for ii = 0 to Ubound(aryAsgMail)
		strAsgTmpLayer = strAsgTmpLayer & "<option name=""" & aryAsgMail(ii) & """"
		if aryAsgMail(ii) = appAsgEmailComponent then 
			strAsgTmpLayer = strAsgTmpLayer & " selected=""selected"""
		end if
		strAsgTmpLayer = strAsgTmpLayer & ">" & aryAsgMail(ii) & "</option>"
	next
	strAsgTmpLayer = strAsgTmpLayer & "</select> <input type=""button"" name=""runtest"" value=""" & BUTTON_Object_test & """ onclick=""testMailObject(document.frmSettings.component.options[document.frmSettings.component.selectedIndex].value);"" />" &_
	"<tr><td align=""right"">" & TXT_Email_server & ":&nbsp;</td><td align=""left""><input type=""text"" name=""server"" value=""" & appAsgEmailServer & """ size=""40"" maxlength=""70"" /></td></tr>" &_
	"<tr><td align=""right"">" & TXT_Email_address & ":&nbsp;</td><td align=""left""><input type=""text"" name=""address"" value=""" & appAsgEmailAddress & """ size=""40"" maxlength=""70"" /></td></tr>" &_
	"</table>"

		' :: Create the layer ::
		Response.Write(buildLayer("layerEmail", LABEL_Settings_emailserver, "", strAsgTmpLayer))
	
	' Submit form area
	Response.Write("<div class=""submitarea""><input type=""submit"" name=""settings"" value=""" & TXT_Update & """ /></div>")

' :: Open tlayout :: MENUSECTION_Config
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