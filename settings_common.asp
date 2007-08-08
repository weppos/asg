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
Call checkPermission("False", "False", "False", appAsgSecurity)

Dim ii
Dim blnAsgShowMonitString		' Set to true to show monitoring string

' Get settings from querystring
if Request.QueryString("monitstring") = 1 then
	blnAsgShowMonitString = true
else
	blnAsgShowMonitString = false
end if


' Update
If Request.QueryString("act") = "upd" then

	' Get data from the form
	appAsgSiteURL = Trim(Request.Form("URLremote"))
		If Left(appAsgSiteURL, 7) <> "http://" Then appAsgSiteURL = "http://" & appAsgSiteURL
		If Right(appAsgSiteURL, 1) <> "/" Then appAsgSiteURL = appAsgSiteURL & "/"
	appAsgSiteName = filterSQLinput(Request.Form("sitename"), true, true)
	
	appAsgStartHits = Clng(Trim(Request.Form("starthits")))
	appAsgStartVisits = Clng(Trim(Request.Form("startvisits")))
	appAsgTimeZone = Request.Form("serverTimeZonePosition") & Request.Form("serverTimeZoneValue")
	
	' Filter data
	' Change from Bolean to Int() type to use data with MySQL
	appAsgRefererServer = Abs(Cint(Cbool(Request.Form("optRefserver"))))
	appAsgDebugIcon = Abs(Cint(Cbool(Request.Form("optCheckicon"))))
	
	appAsgTrackReferer = Abs(Cint(Cbool(Request.Form("monitReferer"))))
	appAsgTrackDaily = Abs(Cint(Cbool(Request.Form("monitDaily"))))
	appAsgTrackIP = Abs(Cint(Cbool(Request.Form("monitIP"))))
	appAsgTrackHourly = Abs(Cint(Cbool(Request.Form("monitHourly"))))
	appAsgTrackSystem = Abs(Cint(Cbool(Request.Form("monitSystem"))))
	appAsgTrackLang = Abs(Cint(Cbool(Request.Form("monitLang"))))
	appAsgTrackPages = Abs(Cint(Cbool(Request.Form("monitPage"))))
	appAsgTrackEngine = Abs(Cint(Cbool(Request.Form("monitEngine"))))
	appAsgTrackCountry = Abs(Cint(Cbool(Request.Form("monitCountry"))))

	' Initialise SQL string to update values
	strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "config SET " &_
	"conf_site_name = '" & appAsgSiteName & "', " &_
	"conf_site_url = '" & appAsgSiteURL & "', " &_
	"conf_start_hits = " & appAsgStartHits & ", " &_
	"conf_start_visits = " & appAsgStartVisits & ", " &_
	"conf_time_zone = '" & appAsgTimeZone & "', " &_
	"conf_referer_server = " & appAsgRefererServer & ", " &_
	"track_referer = " & appAsgTrackReferer & ", " &_
	"track_daily = " & appAsgTrackDaily & ", " &_
	"track_ip = " & appAsgTrackIP & ", " &_
	"track_hourly = " & appAsgTrackHourly & ", " &_
	"track_system = " & appAsgTrackSystem & ", " &_
	"track_lang = " & appAsgTrackLang & ", " &_
	"track_page = " & appAsgTrackPages & ", " &_
	"track_engine = " & appAsgTrackEngine & ", " &_
	"track_country = " & appAsgTrackCountry & ", " &_
	"conf_debug_icon = " & appAsgDebugIcon & " "

	' Execute the update
	' Response.Write(strAsgSQL) : Response.End()
	objAsgConn.Execute(strAsgSQL)
	
	' If application variables are used update them
	if blnApplicationConfig then
						
		' Update application variables
		Application(ASG_APPLICATION_PREFIX & "site_name") = appAsgSiteName
		Application(ASG_APPLICATION_PREFIX & "site_url") = appAsgSiteURL
		Application(ASG_APPLICATION_PREFIX & "StartHits") = CLng(appAsgStartHits)
		Application(ASG_APPLICATION_PREFIX & "StartVisits") = CLng(appAsgStartVisits)
		Application(ASG_APPLICATION_PREFIX & "TimeZone") = appAsgTimeZone
		Application(ASG_APPLICATION_PREFIX & "RefererServer") = CBool(appAsgRefererServer)
		Application(ASG_APPLICATION_PREFIX & "DebugIcon") = CBool(appAsgDebugIcon)
		Application(ASG_APPLICATION_PREFIX & "TrackReferer") = CBool(appAsgTrackReferer)
		Application(ASG_APPLICATION_PREFIX & "TrackDaily") = CBool(appAsgTrackDaily)
		Application(ASG_APPLICATION_PREFIX & "TrackIP") = CBool(appAsgTrackIP)
		Application(ASG_APPLICATION_PREFIX & "TrackHourly") = CBool(appAsgTrackHourly)
		Application(ASG_APPLICATION_PREFIX & "TrackSystem") = CBool(appAsgTrackSystem)
		Application(ASG_APPLICATION_PREFIX & "TrackLang") = CBool(appAsgTrackLang)
		Application(ASG_APPLICATION_PREFIX & "TrackPages") = CBool(appAsgTrackPages)
		Application(ASG_APPLICATION_PREFIX & "TrackEngine") = CBool(appAsgTrackEngine)
		Application(ASG_APPLICATION_PREFIX & "TrackCountry") = CBool(appAsgTrackCountry)

		' Refresh application
		Application(ASG_APPLICATION_PREFIX & "Config") = false
	
	end if
	
	' Reset objects
	Set objAsgRs = Nothing
	objAsgConn.Close
	Set objAsgConn = Nothing
	
	' 
	Response.Redirect("settings_common.asp?act=updmex")

End If

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

<div align="center">
	<div id="layout">
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Administration & " &raquo; " & MENUSECTION_General & " &raquo;</span> " & MENUSECTION_Config %></div>
		<div id="layout_content">

<form action="settings_common.asp?act=upd" name="frmSettings" method="post">
<%

' :: Open tlayout :: MENUSECTION_Config
Response.Write(builTableTlayout("", "open", MENUSECTION_Config))
	
	' Update completed
	if Request.QueryString("act") = "updmex" then Response.Write("<p style=""text-align: center;""><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Update_Completed & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Update_Completed & "</p>")

	' Main settings
	strAsgTmpLayer = "<table align=""center"" border=""0"" cellspacing=""1"" cellpadding=""3"">" 
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right"">" & TXT_SiteName & ":&nbsp;</td><td align=""left""><input type=""text"" name=""sitename"" value=""" & appAsgSiteName & """ size=""60"" maxlength=""140"" /></td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right"">" & TXT_SiteURL & ":&nbsp;</td><td align=""left""><input type=""text"" name=""URLremote"" value="""  
		if "[]" & appAsgSiteURL = "[]" then strAsgTmpLayer = strAsgTmpLayer & "http://" else strAsgTmpLayer = strAsgTmpLayer & appAsgSiteURL
		strAsgTmpLayer = strAsgTmpLayer & """ size=""60"" maxlength=""140"" /></td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right"">" & TXT_StartVisits & ":&nbsp;</td><td align=""left""><input type=""text"" name=""startvisits"" value=""" & appAsgStartVisits & """ size=""10"" maxlength=""8"" /></td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right"">" & TXT_StartHits & ":&nbsp;</td><td align=""left""><input type=""text"" name=""starthits"" value=""" & appAsgStartHits & """ size=""10"" maxlength=""8"" /></td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "</table>"

		' :: Create the layer ::
		Response.Write(buildLayer("layerConfig", LABEL_Settings_site, "", strAsgTmpLayer))

	strAsgTmpLayer = "<table align=""center"" border=""0"" cellspacing=""1"" cellpadding=""3"">" 
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right"">" & TXT_Datetime_offset & "</td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left""><select name=""serverTimeZonePosition"">"
		strAsgTmpLayer = strAsgTmpLayer & "<option value=""+"" "
		if Left(appAsgTimeZone, 1) = "+" then strAsgTmpLayer = strAsgTmpLayer & "selected" 
		strAsgTmpLayer = strAsgTmpLayer & ">+</option>"
		strAsgTmpLayer = strAsgTmpLayer & "<option value=""-"" "
		if Left(appAsgTimeZone, 1) = "-" then strAsgTmpLayer = strAsgTmpLayer & "selected" 
		strAsgTmpLayer = strAsgTmpLayer & ">-</option>"
		strAsgTmpLayer = strAsgTmpLayer & "</select>"
		
		strAsgTmpLayer = strAsgTmpLayer & "<select name=""serverTimeZoneValue"">"
		for ii = 0 to 23
		strAsgTmpLayer = strAsgTmpLayer & "<option value=""" & ii & """ "
			if Cint(Right(appAsgTimeZone, Len(appAsgTimeZone) - 1)) = ii then strAsgTmpLayer = strAsgTmpLayer & "selected"
			strAsgTmpLayer = strAsgTmpLayer & ">" & ii & "</option>"
		next
		strAsgTmpLayer = strAsgTmpLayer & "</select>&nbsp;" & TXT_Datetime_offsetbetw
		strAsgTmpLayer = strAsgTmpLayer & "<br />" & TXT_Datetime_servernow & ":&nbsp;<span class=""notetext"">" & Now() & "</span>"
		
		strAsgTmpLayer = strAsgTmpLayer & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "</table>"

		' :: Create the layer ::
		Response.Write(buildLayer("layerDateTime", LABEL_Settings_datetime, "", strAsgTmpLayer))

	'
	strAsgTmpLayer = "<table align=""center"" border=""0"" cellspacing=""1"" cellpadding=""3"">" 
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""monitReferer"" value=""True"" "
		if appAsgTrackReferer then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & MENUSECTION_Referers & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""monitDaily"" value=""True"" "
		if appAsgTrackDaily then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & MENUSECTION_DailyReports & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""monitHourly"" value=""True"" "
		if appAsgTrackHourly then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & MENUSECTION_HourlyReports & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""monitIP"" value=""True"" "
		if appAsgTrackIP then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & MENUSECTION_IpAddresses & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""monitSystem"" value=""True"" "
		if appAsgTrackSystem then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & MENUSECTION_Systems & ": " & MENUSECTION_Browsers & ", " & MENUSECTION_OS & ", " & MENUSECTION_ResoBit & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""monitLang"" value=""True"" "
		if appAsgTrackLang then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & MENUSECTION_BrowsersLang & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""monitPage"" value=""True"" "
		if appAsgTrackPages then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & MENUSECTION_VisitedPages & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""monitEngine"" value=""True"" "
		if appAsgTrackEngine then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & MENUSECTION_SearchEngines & " " & TXT_And & " " & MENUSECTION_SearchQueries & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""monitCountry"" value=""True"" "
		if appAsgTrackCountry then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & MENUSECTION_Countries & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "</table>"

		' :: Create the layer ::
		Response.Write(buildLayer("layerMonitEnable", LABEL_Settings_tracking, "", strAsgTmpLayer))

	'
	strAsgTmpLayer = "<table align=""center"" border=""0"" cellspacing=""1"" cellpadding=""3"">" 
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""optCheckicon"" value=""True"" "
		if appAsgDebugIcon then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & TXT_Debug_icons & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "<tr><td align=""right""><input type=""checkbox"" name=""optRefserver"" value=""True"" "
		if appAsgRefererServer then strAsgTmpLayer = strAsgTmpLayer & "checked"
		strAsgTmpLayer = strAsgTmpLayer & " /></td>"
	strAsgTmpLayer = strAsgTmpLayer & "<td align=""left"">" & TXT_Option_refserver & "</td></tr>"
	strAsgTmpLayer = strAsgTmpLayer & "</table>"

		' :: Create the layer ::
		Response.Write(buildLayer("layerCheckIcon", LABEL_Settings_misc, "", strAsgTmpLayer))
	
	' Submit form area
	Response.Write("<div class=""submitarea""><input type=""submit"" name=""settings"" value=""" & TXT_Update & """ /></div>")

' Show monitoring string
if blnAsgShowMonitString then

	' Monitoring string
	Response.Write("<a name=""monitstring""></a>")
	strAsgTmpLayer = "<textarea name=""monitstring"" cols=""80"" rows=""3"">" &_
						"&lt;script type=""text/javascript"" language=""JavaScript"" " &_
						"src=""http://" & Request.ServerVariables("HTTP_HOST") & Left(Request.ServerVariables("URL"), InStrRev(Request.ServerVariables("URL"), "/")-1) & "/stats.js.asp"" " &_
						"&gt; &lt;/script&gt;</textarea>"

		' :: Create the layer ::
		Response.Write(buildLayer("layerString", LABEL_Monitstring, "", strAsgTmpLayer))

end if

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