<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="config.asp" -->
<%

'/**
' * ASP Stats Generator - Powerful and reliable ASP website counter
' *
' * This file is part of the ASP Stats Generator package.
' * (c) 2003-2008 Simone Carletti <weppos@weppos.net>, All Rights Reserved
' *
' * 
' * COPYRIGHT AND LICENSE NOTICE
' *
' * The License allows you to download, install and use one or more free copies of this program 
' * for private, public or commercial use.
' * 
' * You may not sell, repackage, redistribute or modify any part of the code or application, 
' * or represent it as being your own work without written permission from the author.
' * You can however modify source code (at your own risk) to adapt it to your specific needs 
' * or to integrate it into your site. 
' *
' * All links and information about the copyright MUST remain unchanged; 
' * you can modify or remove them only if expressly permitted.
' * In particular the license allows you to change the application logo with a personal one, 
' * but it's absolutly denied to remove copyright information,
' * including, but not limited to, footer credits, inline credits metadata and HTML credits comments.
' *
' * For the full copyright and license information, please view the LICENSE.htm
' * file that was distributed with this source code.
' *
' * Removal or modification of this copyright notice will violate the license contract.
' *
' *
' * @category        ASP Stats Generator
' * @package         ASP Stats Generator
' * @author          Simone Carletti <weppos@weppos.net>
' * @copyright       2003-2008 Simone Carletti
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id$
' */
 
'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


'// ATTENZIONE! Protezione statistiche.
'	Modificare solo se necessario e se sicuri.
'	Impostazioni errate possono compromettere la privacy.
Call AllowEntry("False", "False", "False", intAsgSecurity)


'Inserimento record
If Request.Form("Settings") = strAsgTxtUpdate AND Request.QueryString("act") = "upd" Then

	'Pulisci e controlla URL
	strAsgSiteURLremote = Trim(Request.Form("URLremote"))
		If Left(strAsgSiteURLremote, 7) <> "http://" Then strAsgSiteURLremote = "http://" & strAsgSiteURLremote
		If Right(strAsgSiteURLremote, 1) <> "/" Then strAsgSiteURLremote = strAsgSiteURLremote & "/"
	strAsgSiteURLlocal = Trim(Request.Form("URLlocal"))
		If Left(strAsgSiteURLlocal, 7) <> "http://" Then strAsgSiteURLlocal = "http://" & strAsgSiteURLlocal
		If Right(strAsgSiteURLlocal, 1) <> "/" Then strAsgSiteURLlocal = strAsgSiteURLlocal & "/"
	'Richiama dati da Form
	strAsgSiteName = FilterSQLInput(Trim(Server.HTMLEncode(Request.Form("SiteName"))))
	strAsgSiteEmail = Trim(Request.Form("SiteEmail"))
	
	strAsgStartHits = Clng(Trim(Request.Form("StartHits")))
	strAsgStartVisits = Clng(Trim(Request.Form("StartVisits")))
	'Momentaneamente disabilitato per incongruenza periodi accavallati
	'strAsgTimeZone = Request.Form("serverTimeZonePosition") & Request.Form("serverTimeZoneValue") & "|" & Request.Form("gmtTimeZonePosition") & Request.Form("gmtTimeZoneValue")
	strAsgTimeZone = Request.Form("serverTimeZonePosition") & Request.Form("serverTimeZoneValue") & "|+0"
	
	'Filter data
	'Change from Bolean to Int() type to use data with MySQL
	blnRefererServer = CInt(CBool(Request.Form("RefererServer")))
	blnStripPathQS = CInt(CBool(Request.Form("strAsgIPPathQS")))
	blnMonitReferer = CInt(CBool(Request.Form("MonitReferer")))
	blnMonitDaily = CInt(CBool(Request.Form("MonitDaily")))
	blnMonitIP = CInt(CBool(Request.Form("MonitIP")))
	blnMonitHourly = CInt(CBool(Request.Form("MonitHourly")))
	blnMonitSystem = CInt(CBool(Request.Form("MonitSystem")))
	blnMonitLanguages = CInt(CBool(Request.Form("MonitLanguages")))
	blnMonitPages = CInt(CBool(Request.Form("MonitPages")))
	blnMonitEngine = CInt(CBool(Request.Form("MonitEngine")))
	blnMonitCountry = CInt(CBool(Request.Form("MonitCountry")))

	blnAsgCheckIcon = CInt(CBool(Request.Form("CheckIcon")))
	

	'Initialise SQL string to update values
	strAsgSQL = "UPDATE "&strAsgTablePrefix&"config SET " &_
	"Sito_Nome = '" & strAsgSiteName & "', " &_
	"Sito_URL_Remoto = '" & strAsgSiteURLremote & "', " &_
	"Sito_URL_Locale = '" & strAsgSiteURLlocal & "', " &_
	"Sito_Email = '" & strAsgSiteEmail & "', " &_
	"Start_Hits = " & strAsgStartHits & ", " &_
	"Start_Visits = " & strAsgStartVisits & ", " &_
	"Time_Zone = '" & strAsgTimeZone & "', " &_
	"Opt_Referer_Server = " & blnRefererServer & ", " &_
	"Opt_Strip_Path_QS = " & blnStripPathQS & ", " &_
	"Opt_Monit_Referer = " & blnMonitReferer & ", " &_
	"Opt_Monit_Daily = " & blnMonitDaily & ", " &_
	"Opt_Monit_IP = " & blnMonitIP & ", " &_
	"Opt_Monit_Hourly = " & blnMonitHourly & ", " &_
	"Opt_Monit_System = " & blnMonitSystem & ", " &_
	"Opt_Monit_Languages = " & blnMonitLanguages & ", " &_
	"Opt_Monit_Pages = " & blnMonitPages & ", " &_
	"Opt_Monit_Engine = " & blnMonitEngine & ", " &_
	"Opt_Monit_Country = " & blnMonitCountry & ", " &_
	"Opt_Check_Icon = " & blnAsgCheckIcon & " "

	'Execute the update
	objAsgConn.Execute(strAsgSQL)
	'Response.Write(strAsgSQL) : Response.End()
	
	'Se si utilizzano le variabili Application aggiornale
	If blnApplicationConfig Then
						
		'Aggiorna Variabili Application
		Application("strAsgSiteName") = strAsgSiteName
		Application("strAsgSiteURLremote") = strAsgSiteURLremote
		Application("strAsgSiteURLlocal") = strAsgSiteURLlocal
		Application("strAsgSiteEmail") = strAsgSiteEmail
	
		Application("strAsgStartHits") = CLng(strAsgStartHits)
		Application("strAsgStartVisits") = CLng(strAsgStartVisits)
		Application("strAsgTimeZone") = strAsgTimeZone
	
		Application("blnRefererServer") = CBool(blnRefererServer)
		Application("blnStripPathQS") = CBool(blnStripPathQS)
		Application("blnMonitReferer") = CBool(blnMonitReferer)
		Application("blnMonitDaily") = CBool(blnMonitDaily)
		Application("blnMonitIP") = CBool(blnMonitIP)
		Application("blnMonitHourly") = CBool(blnMonitHourly)
		Application("blnMonitSystem") = CBool(blnMonitSystem)
		Application("blnMonitLanguages") = CBool(blnMonitLanguages)
		Application("blnMonitPages") = CBool(blnMonitPages)
		Application("blnMonitEngine") = CBool(blnMonitEngine)
		Application("blnMonitCountry") = CBool(blnMonitCountry)

		Application("blnAsgCheckIcon") = CBool(blnAsgCheckIcon)

		'Forza il ricalcolo delle Application
		Application("blnConfig") = False
	
	End If
	
	'Reset Server Objects
	Set objAsgRs = Nothing
	objAsgConn.Close
	Set objAsgConn = Nothing
	
	'Reindirizza per rivalorizzare dati
	Response.Redirect("settings_common.asp")

End If

'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | powered by ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="copyright" content="Copyright (C) 2003-2008 Carletti Simone, All Rights Reserved" />
<meta name="generator" content="ASP Stats Generator <%= strAsgVersion %>" /> <!-- leave this for stats -->

<!--#include file="asg-includes/layout/head.asp" -->

<!--
  ASP Stats Generator (release <%= strAsgVersion %>) is a free software package
  completely written in ASP programming language, for real time visitor tracking.
  Get your own copy for free at http://www.asp-stats-com/ !
-->

</head>

<!--#include file="asg-includes/layout/header.asp" -->
		<form action="settings_common.asp?act=upd" name="frmImpostazioni" method="post">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtGeneralSettings) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtConfigSettings) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right" width="30%"><%= strAsgTxtSiteName %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left" width="70%">&nbsp;<input type="text" name="SiteName" value="<%= strAsgSiteName %>" size="60" maxlength="140" /></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtSiteURLlocal %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<input type="text" name="URLlocal" value="<% If "[]" & strAsgSiteURLlocal = "[]" Then Response.Write("http://") Else Response.Write(strAsgSiteURLLocal) %>" size="60" maxlength="140" /></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtSiteURLremote %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<input type="text" name="URLremote" value="<% If "[]" & strAsgSiteURLremote = "[]" Then Response.Write("http://") Else Response.Write(strAsgSiteURLremote) %>" size="60" maxlength="140" /></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtSiteEmail %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<input type="text" name="SiteEmail" value="<%= strAsgSiteEmail %>" size="60" maxlength="140" /></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtCountSettings) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%" align="right"><%= strAsgTxtStartVisits %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="70%" align="left">&nbsp;<input type="text" name="StartVisits" value="<%= strAsgStartVisits %>" size="10" maxlength="8" /></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtStartHits %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<input type="text" name="StartHits" value="<%= strAsgStartHits %>" size="10" maxlength="8" /></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtDateSettings) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%" align="right"><%= strAsgTxtTimeZoneOffSet %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="70%" align="left">&nbsp;<select name="serverTimeZonePosition" class="normalform">
					<option value="+" <% If Left(aryAsgTimeZone(0), 1) = "+" Then Response.Write("selected") %>>+</option>
					<option value="-" <% If Left(aryAsgTimeZone(0), 1) = "-" Then Response.Write("selected") %>>-</option>
				</select>
				<select name="serverTimeZoneValue" class="normalform">
					<% For looptmp = 0 to 23 %>
					<option value="<%= looptmp %>" <% If Cint(Right(aryAsgTimeZone(0), Len(aryAsgTimeZone(0))-1)) = looptmp Then Response.Write("selected") %>><%= looptmp %></option>
					<% Next %>
				</select>
				<%= strAsgTxtOffSetClientServer %><br />&nbsp;&nbsp;<%= strAsgTxtServerDateTimeAre & ":&nbsp;<span class=""notetext"">" & Now() & "</span>" %>
			</td>
		  </tr>
		  <!-- Momentaneamente disabilitato per incongruenza periodi accavallati!
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtTimeZoneOffSet %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<select name="gmtTimeZonePosition" class="normalform">
					<option value="+" <% If Left(aryAsgTimeZone(1), 1) = "+" Then Response.Write("selected") %>>+</option>
					<option value="-" <% If Left(aryAsgTimeZone(1), 1) = "-" Then Response.Write("selected") %>>-</option>
				</select>
				<select name="gmtTimeZoneValue" class="normalform">
					<% For looptmp = 0 to 23 %>
					<option value="<%= looptmp %>" <% If Cint(Right(aryAsgTimeZone(1), Len(aryAsgTimeZone(1))-1)) = looptmp Then Response.Write("selected") %>><%= looptmp %></option>
					<% Next %>
				</select>
				<%= strAsgTxtOffSetGMTtoUser %><br />&nbsp;&nbsp;<%= strAsgTxtReportDateTimeAre & ":&nbsp;<span class=""notetext"">" & FormatOutTimeZone(dtmAsgNow, "Date") & "&nbsp;" & FormatOutTimeZone(dtmAsgNow, "Time") & "</span>" %>
			</td>
		  </tr>
		  -->
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtMonitSettings) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtMonitOptions %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">
				&nbsp;<input type="checkbox" name="RefererServer" value="True" <% If blnRefererServer Then Response.Write "checked" %> /> <%= strAsgTxtCountServerAsReferer %><br />
				&nbsp;<input type="checkbox" name="strAsgIPPathQS" value="True" <% If blnStripPathQS Then Response.Write "checked" %> /> <%= strAsgTxtstrAsgIPPathQS %>
			</td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtEnableMonit %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">
				&nbsp;<input type="checkbox" name="MonitReferer" value="True" <% If blnMonitReferer Then Response.Write "checked" %> /> <%= strAsgTxtReferer %><br />
				&nbsp;<input type="checkbox" name="MonitDaily" value="True" <% If blnMonitDaily Then Response.Write "checked" %> /> <%= strAsgTxtDailyMonit %><br />
				&nbsp;<input type="checkbox" name="MonitHourly" value="True" <% If blnMonitHourly Then Response.Write "checked" %> /> <%= strAsgTxtHourlyMonit %><br />
				&nbsp;<input type="checkbox" name="MonitIP" value="True" <% If blnMonitIP Then Response.Write "checked" %> /> <%= strAsgTxtIPAddress %> <br />
				&nbsp;<input type="checkbox" name="MonitSystem" value="True" <% If blnMonitSystem Then Response.Write "checked" %> /> <%= strAsgTxtSystems & ": " & strAsgTxtBrowser & ", " & strAsgTxtOS & ", " & strAsgTxtColor & ", " & strAsgTxtReso %><br />
				&nbsp;<input type="checkbox" name="MonitLanguages" value="True" <% If blnMonitLanguages Then Response.Write "checked" %> /> <%= strAsgTxtBrowserLanguages %><br />
				&nbsp;<input type="checkbox" name="MonitPages" value="True" <% If blnMonitPages Then Response.Write "checked" %> /> <%= strAsgTxtHits %><br />
				&nbsp;<input type="checkbox" name="MonitEngine" value="True" <% If blnMonitEngine Then Response.Write "checked" %> /> <%= strAsgTxtSearchEngine & " " & strAsgTxtAnd & " " & strAsgTxtSearchQuery %><br />
				&nbsp;<input type="checkbox" name="MonitCountry" value="True" <% If blnMonitCountry Then Response.Write "checked" %> /> <%= strAsgTxtCountry %>
			</td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtProgramTools) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">
				&nbsp;<input type="checkbox" name="CheckIcon" value="True" <% If blnAsgCheckIcon Then Response.Write "checked" %> /> <%= strAsgTxtReportUnknownIcons %>
			</td>
		  </tr>
		  <%
					
			'// Row - End table spacer			
			Call BuildTableContEndSpacer(2)
	
		  %>
		  <tr class="normaltitle">
			<td colspan="2" align="center"><br /><input type="submit" name="Settings" value="<%= strAsgTxtUpdate %>" /></td>
		  </tr>
		</table><br />
		</form>

		<!-- write an anchor. It could be usefult in future -->
		<a name="monitoringstring"></a>
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtUsingApplication) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center" colspan="2">
				<%= strAsgTxtPageMonitoringString %><br />
				<textarea name="monitoringstring" cols="80" rows="3">&lt;script type="text/javascript" language="JavaScript" src="http://<%= Request.ServerVariables("HTTP_HOST") & Left(Request.ServerVariables("URL"), InStrRev(Request.ServerVariables("URL"), "/")-1) %>/stats_js.asp"&gt; &lt;/script&gt;
				</textarea>
			</td>
		  </tr>
		</table><br />
<%

' Footer
Response.Write(vbCrLf & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
'// Row - Footer Border Line
Call BuildFooterBorderLine()

' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write("<tr align=""center"" valign=""middle"">")
Response.Write("<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer""><a href=""http://www.asp-stats.com/"" class=""linkfooter"" title=""ASP Stats Generator Homepage"">ASP Stats Generator</a> [" & strAsgVersion & "] - &copy; 2003-2008 <a href=""http://www.weppos.com/"" class=""linkfooter"" title=""Weppos.com Homepage"">weppos</a>")
if blnAsgElabTime then Response.Write(asgElabtime())
Response.Write("</td>")
Response.Write("</tr>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write("</table>")

%>
<!--#include file="asg-includes/layout/footer.asp" -->

</body></html>
