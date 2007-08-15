<%

'/**
' * ASP Stats Generator - Powerful and reliable ASP website counter
' *
' * This file is part of the ASP Stats Generator package.
' * (c) 2003-2007 Simone Carletti <weppos@weppos.net>, All Rights Reserved
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
' * @copyright       2003-2007 Simone Carletti, All Rights Reserved
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id$
' */
 
'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


%>
<body bgcolor="<%= strAsgSknPageBgColour %>" background="<%= strAsgSknPageBgImage %>">
<table width="<%= strAsgSknPageWidth %>" border="0" align="center" cellpadding="1" cellspacing="0"><tr>
<td width="150" align="left" valign="top">&nbsp;<a href="http://www.weppos.com/" title="ASP Stats Generator Home Page"><img src="images/logo.gif" border="0" alt="Powered by ASP Stats Generator" /></a></td>
<td width="750" align="right" valign="top">
<table width="750" border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
  <tr><td>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="<%= strAsgSknTableLayoutBgColour %>">
	  <tr><td>
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr align="center" class="bartitle" bgcolor="<%= strAsgSknTableBarBgColour %>" valign="middle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" width="16%" height="20"><%= UCase(strAsgTxtGeneral) %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" width="16%"><%= UCase(strAsgTxtSystems) %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" width="16%"><%= UCase(strAsgTxtStats) %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" width="16%"><%= UCase(strAsgTxtProvenance) %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" width="16%"><%= UCase(strAsgTxtExtra) %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>"><%= UCase(strAsgTxtOptions) %></td>
		  </tr>
		</table>
	  </td></tr>
	</table>
  </td></tr>
</table>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="">
  <tr><td>
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
	  <tr align="center" class="smalltext">
		<td bgcolor="<%= strAsgSknTableContBgColour %>" width="16%" height="16"><a href="statistiche.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtIndexReport %>" class="linksmalltext"><%= strAsgTxtIndexReport %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>" width="16%"><a href="os.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtOS %>" class="linksmalltext"><%= strAsgTxtOS %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>" width="16%"><a href="stats_hourly.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtVisitsPerHour %>" class="linksmalltext"><%= strAsgTxtVisitsPerHour %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>" width="16%"><a href="referer.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtReferer %>" class="linksmalltext"><%= strAsgTxtReferer %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>" width="16%"><a href="ip_address.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtIPAddress %>" class="linksmalltext"><%= strAsgTxtIPAddress %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="settings_common.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtGeneralSettings %>" class="linksmalltext"><%= strAsgTxtGeneralSettings %></a></td>
	  </tr>
	  <tr align="center" class="smalltext">
		<td bgcolor="<%= strAsgSknTableContBgColour %>" height="16"><a href="visitors.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtVisitorsDetails %>" class="linksmalltext"><%= strAsgTxtVisitorsDetails %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="browser.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtBrowser %>" class="linksmalltext"><%= strAsgTxtBrowser %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="stats_daily.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtVisitsPerDay %>" class="linksmalltext"><%= strAsgTxtVisitsPerDay %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="engine.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtSearchEngine %>" class="linksmalltext"><%= strAsgTxtSearchEngine %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="country.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtCountry %>" class="linksmalltext"><%= strAsgTxtCountry %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="settings_security.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtSecuritySettings %>" class="linksmalltext"><%= strAsgTxtSecuritySettings %></a></td>
	  </tr>
	  <tr align="center" class="smalltext">
		<td bgcolor="<%= strAsgSknTableContBgColour %>" height="16"><a href="pages.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtHits %>" class="linksmalltext"><%= strAsgTxtHits %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="browser_lang.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtBrowserLanguages %>" class="linksmalltext"><%= strAsgTxtBrowserLanguages %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="stats_monthly.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtVisitsPerMonth %>" class="linksmalltext"><%= strAsgTxtVisitsPerMonth %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="query.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtSearchQuery %>" class="linksmalltext"><%= strAsgTxtSearchQuery %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="serp.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtSERPreports %>" class="linksmalltext"><%= strAsgTxtSERPreports %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="settings_reset.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtResetSettings %>" class="linksmalltext"><%= strAsgTxtResetSettings %></a></td>
	  </tr>
	  <tr align="center" class="smalltext">
		<td height="16">&nbsp;</td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="color.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtSmReso & " & " & strAsgTxtSmColor %>" class="linksmalltext"><%= strAsgTxtSmReso & " & " & strAsgTxtSmColor %></a></td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="stats_calendar.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtMonthlyCalendar %>" class="linksmalltext"><%= strAsgTxtMonthlyCalendar %></a></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="settings_exitcount.asp" title="<%= strAsgTxtShow & "&nbsp;" %>" class="linksmalltext"><%= strAsgTxtExclusionSettings %></a></td>
	  </tr>
	  <tr align="center" class="smalltext">
		<td height="16">&nbsp;</td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="system.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtSystems %>" class="linksmalltext"><%= strAsgTxtSystems %></a></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="settings_skin.asp" title="<%= strAsgTxtSkinSettings %>" class="linksmalltext"><%= strAsgTxtSkinSettings %></a></td>
	  </tr>
	  <tr align="center" class="smalltext">
		<td height="16">&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="check_server.asp" title="<%= strAsgTxtShow & "&nbsp;" & strAsgTxtServerInformations %>" class="linksmalltext"><%= strAsgTxtServerInformations %></a></td>
	  </tr>
	  <tr align="center" class="smalltext">
		<td height="16">&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td bgcolor="<%= strAsgSknTableContBgColour %>"><a href="compact_database.asp" title="<%= strAsgCompactDatabase  %>" class="linksmalltext"><%= strAsgCompactDatabase  %></a></td>
	  </tr>
	</table>
  </td></tr>
</table>
<% If Session("AsgLogin") = "Logged" Then Response.Write "<br /><div align=""right""><a href=""login.asp?Logout=True"" title=""" & strAsgTxtLogout & """ class=""linksmalltext"">" & strAsgTxtLogout & " &raquo;</a>&nbsp;<br /></div>" %>

</tr></table><br />

<%

'-----------------------------------------------------------------------------------------
' Check version for update!
'-----------------------------------------------------------------------------------------
' Esecuzioni in base a controllo
Select Case intAsgLastUpdate
	
	Case 0
		'Non calcolato
	Case 1
		'Corrisponde
	Case 2
		'Display the alert
		Response.Write("<script language=""JavaScript""><!-- ")
		Response.Write(vbCrLf & vbCrLf & "//Show the popup")
		Response.Write(vbCrLf & "checkUpdate = confirm('Available " & strAsgLastVersion & " version released on " & Right(dtmAsgLastUpdate, 2) & "/" & Mid(dtmAsgLastUpdate, 5, 2) & "/" & Left(dtmAsgLastUpdate, 4) & "! \nDownload the update?')")
		Response.Write(vbCrLf & "if (checkUpdate == true) {")
		If Len(urlAsgLastUpdate) > 0 Then
		Response.Write(vbCrLf & "	window.location='" & urlAsgLastUpdate & "'")
		Else
		Response.Write(vbCrLf & "	window.location='http://www.weppos.com/asg/'")
		End If
		Response.Write(vbCrLf & "}")
		Response.Write(vbCrLf & "// --></script>")
	Case 3
		'Display the alert
		Response.Write("<script language=""JavaScript""><!-- ")
		Response.Write(vbCrLf & vbCrLf & "//Show the popup")
		Response.Write(vbCrLf & "checkUpdate = confirm('Available for your " & strAsgLastVersion & " version an update released on " & Right(dtmAsgLastUpdate, 2) & "/" & Mid(dtmAsgLastUpdate, 5, 2) & "/" & Left(dtmAsgLastUpdate, 4) & ". \nDownload the update?')")
		Response.Write(vbCrLf & "if (checkUpdate == true) {")
		If Len(urlAsgLastUpdate) > 0 Then
		Response.Write(vbCrLf & "	window.location='" & urlAsgLastUpdate & "'")
		Else
		Response.Write(vbCrLf & "	window.location='http://www.weppos.com/asg/'")
		End If
		Response.Write(vbCrLf & "}")
		Response.Write(vbCrLf & "// --></script>")

End Select



Response.Write(vbCrLf & "<table width=""" & strAsgSknPageWidth & """ border=""0"" align=""center"" cellpadding=""1"" cellspacing=""0"" bgcolor=""" & strAsgSknTableLayoutBorderColour & """>")
Response.Write(vbCrLf & "  <tr><td>")
Response.Write(vbCrLf & "	<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
Response.Write(vbCrLf & "	  <tr><td bgcolor=""" & strAsgSknTableLayoutBgColour & """ background=""" & strAsgSknPathImage & strAsgSknTableLayoutBgImage & """>")

%>