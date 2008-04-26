<%@ LANGUAGE="VBSCRIPT" %>
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


Dim strAsgConfigOk
strAsgConfigOk = False

If Request.Form("imposta") = strAsgTxtUpdate Then
	
	'Pulisci e controlla URL
	strAsgSiteURLremote = Trim(Request.Form("URLremote"))
		If Left(strAsgSiteURLremote, 7) <> "http://" Then strAsgSiteURLremote = "http://" & strAsgSiteURLremote
		If Right(strAsgSiteURLremote, 1) <> "/" Then strAsgSiteURLremote = strAsgSiteURLremote & "/"
	strAsgSiteURLlocal = Trim(Request.Form("URLlocal"))
		If Left(strAsgSiteURLlocal, 7) <> "http://" Then strAsgSiteURLlocal = "http://" & strAsgSiteURLlocal
		If Right(strAsgSiteURLlocal, 1) <> "/" Then strAsgSiteURLlocal = strAsgSiteURLlocal & "/"

	'Richiama dati da Form
	strAsgSiteName = Trim(Server.HTMLEncode(Request.Form("SiteName")))
	strAsgSiteEmail = Trim(Request.Form("SiteEmail"))
	
	strAsgStartHits = Clng(Trim(Request.Form("StartHits")))
	strAsgStartVisits = Clng(Trim(Request.Form("StartVisits")))
	strAsgTimeZone = Request.Form("serverTimeZonePosition") & Request.Form("serverTimeZoneValue") & "|+0"
	
	'Pulisci Password
	If Len(Trim(Request.Form("Password"))) > 0 Then strAsgSitePsw = CleanInput(Trim(Request.Form("Password")))
	'Richiama varie
	If IsNumeric(Request.Form("Protezione")) Then intAsgSecurity = CInt(Request.Form("Protezione"))


	strAsgSQL = "SELECT TOP 1 * FROM "&strAsgTablePrefix&"Config"

	objAsgRs.CursorType = 3
	objAsgRs.LockType = 3
	objAsgRs.Open strAsgSQL, objAsgConn
	
	If objAsgRs.EOF Then 
		objAsgRs.AddNew
	End If
	
	With objAsgRs

		.Fields("Sito_Nome") = strAsgSiteName
		.Fields("Sito_URL_Remoto") = strAsgSiteURLremote
		.Fields("Sito_URL_Locale") = strAsgSiteURLlocal
		.Fields("Sito_Email") =strAsgSiteEmail
		.Fields("Image") = "stats.gif"

		.Fields("Start_Stats") = Now()
		
		If IsNumeric(Request.Form("StartHits")) Then .Fields("Start_Hits") = Clng(Trim(Request.Form("StartHits")))
		If IsNumeric(Request.Form("StartVisits")) Then .Fields("Start_Visits") = Clng(Trim(Request.Form("StartVisits")))

		.Fields("Filter_IP") = Replace(Request.Form("FilterIP"), " ", "")
		.Fields("Time_Zone") = strAsgTimeZone
		
		.Fields("Opt_Referer_Server") = CBool(Request.Form("RefererServer"))
		.Fields("Opt_Strip_Path_QS") = CBool(Request.Form("strAsgIPPathQS"))
		.Fields("Opt_Monit_Referer") = CBool(Request.Form("MonitReferer"))
		.Fields("Opt_Monit_Daily") = CBool(Request.Form("MonitDaily"))
		.Fields("Opt_Monit_IP") = CBool(Request.Form("MonitIP"))
		.Fields("Opt_Monit_Hourly") = CBool(Request.Form("MonitHourly"))
		.Fields("Opt_Monit_System") = CBool(Request.Form("MonitSystem"))
		.Fields("Opt_Monit_Languages") = CBool(Request.Form("MonitLanguages"))
		.Fields("Opt_Monit_Pages") = CBool(Request.Form("MonitPages"))
		.Fields("Opt_Monit_Engine") = CBool(Request.Form("MonitEngine"))
		.Fields("Opt_Monit_Country") = CBool(Request.Form("MonitCountry"))
		.Fields("Opt_Check_Icon") = True
		.Fields("Opt_Check_Update") = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)

		.Fields("Sito_PSW") = strAsgSitePsw
		.Fields("Stats_Protezione") = intAsgSecurity
		
		.Update
		.Requery
	End With
	
	strAsgConfigOk = True
	objAsgRs.Close

End If

'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Installazione ASP Stats Generator | Versione <%= ASG_VERSION %></title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<style type="text/css">
<!--
body, div, p, tr, td {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10pt;
	font-weight: normal;
	color: #000000;
}

.menutitle {
	color: #0066CC;
	line-height: 140%;
}

a {
	color: #0066CC;
	text-decoration: none;
}
a:hover {
	color: #0066CC;
	text-decoration: underline;
}
a:visited  {
	color: #0066CC;
	text-decoration: none;
}
a:visited:hover {
	color: #0066CC;
	text-decoration: underline;
}

.normaltitle, h2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	color: #0066CC;
	font-weight: bold;
	line-height: 140%;
	text-align: center;
	font-variant: small-caps;
}

h3 {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 14px;
	color: #0066CC;
	font-weight: bold;
	line-height: 120%;
	font-variant: small-caps;
}

#content {
	padding-left: 10px;
	padding-right: 10px;
}
-->
</style>
<% If Request.Form("imposta") <> strAsgTxtUpdate Then %>
<script language="Javascript" type="text/javascript">
// Verifica campi setup
function controlla(form) { 
	
	if (form.SiteName.value == "") { 
		alert('Inserire un nome per il Sito'); 
		return false; 
		} 
	
	if (form.URLremote.value == "http://") { 
		alert("Inserire un indirizzo web valido per il sito"); 
		return false; 
		} 
	
	if (form.SiteEmail.value == "") { 
		alert("Inserire una Email di riferimento per il sito"); 
		return false; 
		} 
	
	if (form.Password.value == "") { 
		alert("Inserire una password di protezione valida"); 
		return false; 
		} 
	
	return true; 
	} 
</script>
<% End If %>
</head>

<body>

<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr><td height="50" valign="middle"><img src="images/asg_setupfile.gif" width="276" height="28" border="0" alt="ASP Stats Generator" /></td></tr>
  <tr><td height="2" bgcolor="#FDD353"></td></tr>
  <tr><td height="15" valign="middle" align="right" class="x-smalltext">Installazione ASP Stats Generator - v. <%= ASG_VERSION %></td></tr>
</table>
<% If Request.Form("imposta") = strAsgTxtUpdate AND strAsgConfigOk Then %>
<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="100%" colspan="2" align="center"><br />
	<% If infoAsgTypeLanguage = "italiano" Then %>
		La procedura di configurazione è stata completata con successo.<br />
		L'applicazione è ora attiva.<br /><br />
		Ricordarsi di inserire nelle pagine da monitorare la stringa di monitoraggio <br />
		come riportato nella documentazione di aiuto. <br />
	<% Else %>
		The configuration has been succesfully completed.<br />
		Application is now active.<br /><br />
		Please remember to put monitoring string into pages you would like to track <br />
		as written in ufficial documentation. <br />
	<% End If %>
	</td>
  </tr>
</table>
<% ElseIf Request.Form("imposta") = strAsgTxtUpdate AND strAsgConfigOk = False Then %>
<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center">
    <td width="100%" colspan="2"><br />
	<% If infoAsgTypeLanguage = "italiano" Then %>
		Si sono verificati errori imprevisti nella configurazione.<br />
		Verificare il funzionamento dell'applicazione ed in caso di problemi 
		ripetere la procedura di configurazione.<br />
	<% Else %>
		Some errors happened during configuration.<br />
		Please check application status and path and repeat the configuration.<br />
	<% End If %>
	</td>
  </tr>
</table>
<% Else %>
<form action="setup.asp" method="post" name="frmSetupApp" onSubmit="return controlla(this)">
<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="100%" colspan="2"><div align="justify"><br />
	<% If infoAsgTypeLanguage = "italiano" Then %>
		Se visualizzi questo messaggio vuol dire che la connessione 
		al database è stata impostata correttamente
		e puoi ora procedere con la configurazione del database e 
		delle impostazioni.<br />
	<% Else %>
		If you see this message it means that database and file path have been correctly
		adapted and now you can carry on with the configuration.<br />
	<% End If %>
	<br /></div>
	</td>
  </tr>
 </table>
<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="50%"></td>
    <td width="50%"></td>
  </tr>
  <tr class="menutitle">
	<td colspan="2" align="center" height="30"><br /><%= UCase(strAsgTxtConfigSettings) %><br /><br /></td>
  </tr>
<% If infoAsgTypeLanguage = "italiano" Then %>
<!-- italiano -->
  <tr>
	<td align="right"><%= strAsgTxtSiteName %>:<br />
	<span class="smalltext">Nome che verrà visualizzato nelle pagine e nei titoli del programma di statistica</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="SiteName" value="<%= strAsgSiteName %>" size="50" maxlength="140" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtSiteURLlocal %>:<br />
	<span class="smalltext">Indirizzo del sito testato in locale [Facoltativo]</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="URLlocal" value="<% If "[]" & strAsgSiteURLlocal = "[]" Then Response.Write("http://") Else Response.Write(strSiteURLLocal) %>" size="50" maxlength="140" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtSiteURLremote %>: <br />
	<span class="smalltext">Indirizzo web del sito internet</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="URLremote" value="<% If "[]" & strAsgSiteURLremote = "[]" Then Response.Write("http://") Else Response.Write(strAsgSiteURLremote) %>" size="50" maxlength="140" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtSiteEmail %>: <br />
	<span class="smalltext">E-mail di riferimento per i report e le comunicazioni dal sistema</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="SiteEmail" value="<%= strAsgSiteEmail %>" size="50" maxlength="140" /></td>
  </tr>
  <tr class="menutitle">
	<td colspan="2" align="center" height="30"><br /><%= UCase(strAsgTxtCountSettings) %><br /><br /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtStartVisits %>: <br />
	<span class="smalltext">Numero di Accessi Unici da cui il sistema inizerà il conteggio</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="StartVisits" value="<%= strAsgStartVisits %>" size="10" maxlength="8" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtStartHits %>: <br />
	<span class="smalltext">Numero di Pagine Visitate da cui il sistema inizerà il conteggio</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="StartHits" value="<%= strAsgStartHits %>" size="10" maxlength="8" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtFilterIPaddr %>: <br />
	<span class="smalltext">Immettere dgli indirizzi IP da escludere nel conteggio separati da <strong>,</strong><br /> senza usare spazi</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="FilterIP" value="<%= strAsgFilterIP %>" size="50" maxlength="200" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtTimeZoneOffSet %>: </td>
	<td align="left">&nbsp;&nbsp;&nbsp;<select name="serverTimeZonePosition">
		<option value="+" <% If Left(aryAsgTimeZone(0), 1) = "+" Then Response.Write("selected") %>>+</option>
		<option value="-" <% If Left(aryAsgTimeZone(0), 1) = "-" Then Response.Write("selected") %>>-</option>
	</select>
	<select name="serverTimeZoneValue">
		<% For looptmp = 0 to 23 %>
		<option value="<%= looptmp %>" <% If Cint(Right(aryAsgTimeZone(0), Len(aryAsgTimeZone(0))-1)) = looptmp Then Response.Write("selected") %>><%= looptmp %></option>
		<% Next %>
	</select>
	<span class="smalltext"><%= strAsgTxtOffSetClientServer %><br />&nbsp;&nbsp;&nbsp;<%= strAsgTxtServerDateTimeAre & ":&nbsp;<span class=""notetext"">" & Now() & "</span>" %></span>
    </td>
  </tr>
  <tr class="menutitle">
	<td colspan="2" align="center" height="30"><br /><%= UCase(strAsgTxtMonitSettings) %><br /><br /></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtCountServerAsReferer %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="RefererServer" value="True" checked /> <%= strAsgTxtCountServerAsReferer %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtStripPathQS %>.<br />
	www.mysite.com/page.asp?id=3 --&gt; www.mysite.com/page.asp</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="strAsgIPPathQS" value="True" /> <%= strAsgTxtStripPathQS %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtReferer %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitReferer" value="True" checked /> <%= strAsgTxtReferer %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtDailyMonit %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitDaily" value="True" checked /> <%= strAsgTxtDailyMonit %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtHourlyMonit %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitHourly" value="True" checked /> <%= strAsgTxtHourlyMonit %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtIPAddress %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitIP" value="True" checked /> <%= strAsgTxtIPAddress %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtSystems %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitSystem" value="True" checked /> <%= strAsgTxtSystems %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtBrowserLanguages %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitLanguages" value="True" checked /> <%= strAsgTxtBrowserLanguages %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtHits %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitPages" value="True" checked /> <%= strAsgTxtHits %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtSearchEngine & " " & strAsgTxtAnd & " " & strAsgTxtSearchQuery %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitEngine" value="True" checked /> <%= strAsgTxtSearchEngine & " " & strAsgTxtAnd & " " & strAsgTxtSearchQuery %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtCountry %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitCountry" value="True" checked /> <%= strAsgTxtCountry %></td>
  </tr>
  <tr class="menutitle">
	<td colspan="2" align="center" height="30"><br /><%= UCase(strAsgTxtSecuritySettings) %><br /><br /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtEntryPassword %>:</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="Password" value="" size="20" maxlength="20" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtStatsProtectionLevel %>:</span></td>
	<td align="left">&nbsp;&nbsp;
	&nbsp;<input type="radio" name="Protezione" value="0" <% If intAsgSecurity = 0 Then Response.Write "checked" %> /><%= strAsgTxtNone %>
	&nbsp;<input type="radio" name="Protezione" value="1" <% If intAsgSecurity = 1 Then Response.Write "checked" %> /><%= strAsgTxtLimited %>
	&nbsp;<input type="radio" name="Protezione" value="2" <% If intAsgSecurity = 2 Then Response.Write "checked" %> /><%= strAsgTxtFull %>
	</td>
  </tr>
<!-- / italiano -->
<% Else %>
<!-- english -->
  <tr>
	<td align="right"><%= strAsgTxtSiteName %>:<br />
	<span class="smalltext">The name will be show into the pages title</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="SiteName" value="<%= strAsgSiteName %>" size="50" maxlength="140" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtSiteURLlocal %>:<br />
	<span class="smalltext">URL used in local for testing [Optional]</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="URLlocal" value="<% If "[]" & strAsgSiteURLlocal = "[]" Then Response.Write("http://") Else Response.Write(strSiteURLLocal) %>" size="50" maxlength="140" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtSiteURLremote %>: <br />
	<span class="smalltext">Ufficial site URL</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="URLremote" value="<% If "[]" & strAsgSiteURLremote = "[]" Then Response.Write("http://") Else Response.Write(strAsgSiteURLremote) %>" size="50" maxlength="140" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtSiteEmail %>: <br />
	<span class="smalltext">Reference Email of the site for comunications</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="SiteEmail" value="<%= strAsgSiteEmail %>" size="50" maxlength="140" /></td>
  </tr>
  <tr class="menutitle">
	<td colspan="2" align="center" height="30"><br /><%= UCase(strAsgTxtCountSettings) %><br /><br /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtStartVisits %>: <br />
	<span class="smalltext">Counting of the visitors will start from this value</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="StartVisits" value="<%= strAsgStartVisits %>" size="10" maxlength="8" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtStartHits %>: <br />
	<span class="smalltext">Counting of the pages will start from this value</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="StartHits" value="<%= strAsgStartHits %>" size="10" maxlength="8" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtFilterIPaddr %>: <br />
	<span class="smalltext">Insert IP addresses you would like to exclude from monitoring separed with <strong>,</strong><br /> without using spaces</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="FilterIP" value="<%= strAsgFilterIP %>" size="50" maxlength="200" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtTimeZoneOffSet %>: </td>
	<td align="left">&nbsp;&nbsp;&nbsp;<select name="serverTimeZonePosition">
		<option value="+" <% If Left(aryAsgTimeZone(0), 1) = "+" Then Response.Write("selected") %>>+</option>
		<option value="-" <% If Left(aryAsgTimeZone(0), 1) = "-" Then Response.Write("selected") %>>-</option>
	</select>
	<select name="serverTimeZoneValue">
		<% For looptmp = 0 to 23 %>
		<option value="<%= looptmp %>" <% If Cint(Right(aryAsgTimeZone(0), Len(aryAsgTimeZone(0))-1)) = looptmp Then Response.Write("selected") %>><%= looptmp %></option>
		<% Next %>
	</select>
	<span class="smalltext"><%= strAsgTxtOffSetClientServer %><br />&nbsp;&nbsp;&nbsp;<%= strAsgTxtServerDateTimeAre & ":&nbsp;<span class=""notetext"">" & Now() & "</span>" %></span>
    </td>
  </tr>
  <tr class="menutitle">
	<td colspan="2" align="center" height="30"><br /><%= UCase(strAsgTxtMonitSettings) %><br /><br /></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtCountServerAsReferer %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="RefererServer" value="True" checked /> <%= strAsgTxtCountServerAsReferer %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtStripPathQS %>.<br />
	www.mysite.com/page.asp?id=3 --&gt; www.mysite.com/page.asp</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="strAsgIPPathQS" value="True" /> <%= strAsgTxtStripPathQS %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtReferer %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitReferer" value="True" checked /> <%= strAsgTxtReferer %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtDailyMonit %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitDaily" value="True" checked /> <%= strAsgTxtDailyMonit %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtHourlyMonit %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitHourly" value="True" checked /> <%= strAsgTxtHourlyMonit %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtIPAddress %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitIP" value="True" checked /> <%= strAsgTxtIPAddress %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtSystems %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitSystem" value="True" checked /> <%= strAsgTxtSystems %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtBrowserLanguages %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitLanguages" value="True" checked /> <%= strAsgTxtBrowserLanguages %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtHits %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitPages" value="True" checked /> <%= strAsgTxtHits %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtSearchEngine & " " & strAsgTxtAnd & " " & strAsgTxtSearchQuery %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitEngine" value="True" checked /> <%= strAsgTxtSearchEngine & " " & strAsgTxtAnd & " " & strAsgTxtSearchQuery %></td>
  </tr>
  <tr>
	<td align="right"><span class="smalltext"><%= strAsgTxtEnableMonit & "&nbsp;" & strAsgTxtCountry %></span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="checkbox" name="MonitCountry" value="True" checked /> <%= strAsgTxtCountry %></td>
  </tr>
  <tr class="menutitle">
	<td colspan="2" align="center" height="30"><br /><%= UCase(strAsgTxtSecuritySettings) %><br /><br /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtEntryPassword %>:</span></td>
	<td align="left">&nbsp;&nbsp;&nbsp;<input type="text" name="Password" value="" size="20" maxlength="20" /></td>
  </tr>
  <tr>
	<td align="right"><%= strAsgTxtStatsProtectionLevel %>:</span></td>
	<td align="left">&nbsp;&nbsp;
	&nbsp;<input type="radio" name="Protezione" value="0" <% If intAsgSecurity = 0 Then Response.Write "checked" %> /><%= strAsgTxtNone %>
	&nbsp;<input type="radio" name="Protezione" value="1" <% If intAsgSecurity = 1 Then Response.Write "checked" %> /><%= strAsgTxtLimited %>
	&nbsp;<input type="radio" name="Protezione" value="2" <% If intAsgSecurity = 2 Then Response.Write "checked" %> /><%= strAsgTxtFull %>
	</td>
  </tr>
<!-- / english -->
<% End If %>
  <tr>
	<td align="center" colspan="2" height="40"><br /><br />
		<input type="hidden" name="step" value="1" />&nbsp;&nbsp;
		<input type="reset" name="reset" value="Reset" />&nbsp;&nbsp;
		<input type="submit" name="imposta" value="<%= strAsgTxtUpdate %>" />
	</td>
  </tr>
</table>
</form>
<% End If %>

</body>
</html>
