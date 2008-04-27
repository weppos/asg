<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="config.asp" -->
<%

' 
' = ASP Stats Generator - Powerful and reliable ASP website counter
' 
' Copyright (c) 2003-2008 Simone Carletti <weppos@weppos.net>
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
' 
' 
' @category        ASP Stats Generator
' @package         ASP Stats Generator
' @author          Simone Carletti <weppos@weppos.net>
' @copyright       2003-2008 Simone Carletti
' @license         http://www.opensource.org/licenses/mit-license.php
' @version         SVN: $Id$
' 


'// ATTENZIONE! Protezione statistiche.
'	Modificare solo se necessario e se sicuri.
'	Impostazioni errate possono compromettere la privacy.
Call AllowEntry("False", "False", "False", intAsgSecurity)


'Esegui aggiornamenti

'//	Escludi
If Request.QueryString("act") = "excludepc" Then
	Response.Cookies(strAsgCookiePrefix& "exitcount") = "excludepc"
	Response.Cookies(strAsgCookiePrefix& "exitcount").Expires = dateAdd("yyyy", 1, date)

'//	Includi
ElseIf Request.QueryString("act") = "includepc" Then
	Response.Cookies(strAsgCookiePrefix& "exitcount") = ""
	Response.Cookies(strAsgCookiePrefix& "exitcount").Expires = dateAdd("yyyy", 1, date)

'//	Aggiorna
ElseIf Request.Form("Settings") = strAsgTxtUpdate AND Request.QueryString("act") = "upd" Then

	strAsgFilterIP = Trim(Request.Form("FilterIP"))
		'Pulisci spazi
		strAsgFilterIP = Replace(strAsgFilterIP, " ", "")

	'Inizializza SQL
	strAsgSQL = "UPDATE "&strAsgTablePrefix&"Config SET Filter_IP = '" & strAsgFilterIP & "'"
	objAsgConn.Execute(strAsgSQL)
	
	'Se si utilizzano le variabili Application aggiornale
	If blnApplicationConfig Then
						
		'Aggiorna Variabili Application
		Application("strAsgFilterIP") = strAsgFilterIP

		'Forza il ricalcolo delle Application
		Application("blnConfig") = False
	
	End If
	
	'Reset Server Objects
	Set objAsgRs = Nothing
	objAsgConn.Close
	Set objAsgConn = Nothing
	
	'Reindirizza per rivalorizzare dati
	Response.Redirect("settings_exitcount.asp")

End If

'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | powered by ASP Stats Generator <%= ASG_VERSION %></title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="copyright" content="Copyright (C) 2003-2008 Carletti Simone, All Rights Reserved" />
<meta name="generator" content="ASP Stats Generator <%= ASG_VERSION %>" /> <!-- leave this for stats -->

<!--#include file="asg-includes/layout/head.asp" -->

<!--
  ASP Stats Generator (release <%= ASG_VERSION %>) is a free software package
  completely written in ASP programming language, for real time visitor tracking.
  Get your own copy for free at http://www.asp-stats-com/ !
-->

</head>

<!--#include file="asg-includes/layout/header.asp" -->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtExclusionSettings) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		<form action="settings_exitcount.asp?act=upd" name="frmSettings" method="post">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtExitByIP) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%" align="right"><%= strAsgTxtFilterIPaddr %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="70%" align="left">&nbsp;<input type="text" name="FilterIP" value="<%= strAsgFilterIP %>" size="60" maxlength="200" /><input type="submit" name="Settings" value="<%= strAsgTxtUpdate %>" /><br /><%= strAsgTxtInformationsToExitByIpRange %></td>
		  </tr>
		</form>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtExitByCookie) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" height="18" align="right">
			<%
				'Open the string
				Response.Write(strAsgTxtThisPCisActually)
				
				If Request.Cookies(strAsgCookiePrefix& "exitcount") = "excludepc" Then
					Response.Write("&nbsp;&nbsp;&nbsp;<br /><span class=""notetext"">" & strAsgTxtExcluded & "&nbsp;</span>")
				Else
					Response.Write("&nbsp;&nbsp;&nbsp;<br /><span class=""notetext"">" & strAsgTxtIncluded & "&nbsp;</span>")
				End If
				
				'Close the string
				Response.Write(strAsgTxtIntoMonitoringProcess)
			
			%>&nbsp;&nbsp;&nbsp;
			</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>">&nbsp;<%
			
				If Request.Cookies(strAsgCookiePrefix& "exitcount") = "excludepc" Then
					Response.Write("<a href=""settings_exitcount.asp?act=includepc"" title=""" & strAsgTxtIncludePC & """><img src=""" & strAsgSknPathImage & "include.gif"" alt=""" & strAsgTxtIncludePC & """ border=""0""></a>")
				Else
					Response.Write("<a href=""settings_exitcount.asp?act=excludepc"" title=""" & strAsgTxtExcludePC & """><img src=""" & strAsgSknPathImage & "exclude.gif"" alt=""" & strAsgTxtExcludePC & """ border=""0""></a>")
				End If
			
			%>
			</td>
		  </tr>
		<%
				
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(2)

		%>
		</table><br />
<%

' Footer
Response.Write(vbCrLf & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
'// Row - Footer Border Line
Call BuildFooterBorderLine()

' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write("<tr align=""center"" valign=""middle"">")
Response.Write("<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer""><a href=""http://www.asp-stats.com/"" class=""linkfooter"" title=""ASP Stats Generator Homepage"">ASP Stats Generator</a> " & ASG_VERSION & " - &copy; 2003-2008 <a href=""http://www.weppos.com/"" class=""linkfooter"" title=""Weppos.com Homepage"">weppos</a>")
if ASG_CONFIG_ELABTIME then Response.Write(" - " & asgElabtime())
Response.Write("</td>")
Response.Write("</tr>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write("</table>")

%>
<!--#include file="asg-includes/layout/footer.asp" -->

</body></html>
