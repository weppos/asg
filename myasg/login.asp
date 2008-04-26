<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--include virtual="/myasg/config.asp" -->
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


'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

'Verifica Password
If Request.Form("Login") = strAsgTxtLogin Then

	Dim strAsgPassword
	Dim blnAsgErrore
	
	blnAsgErrore = False
	
	strAsgPassword = Trim(Request.Form("Password"))
	strAsgPassword = CleanInput(strAsgPassword)

	'Verifica
	If LCase(strAsgPassword) = LCase(strAsgSitePsw) Then
	
		'1° Versione --> Uso variabili di sessione
		'prossima implementazione cookie
		
		Session("AsgLogin") = "Logged"
		
	Else

		blnAsgErrore = True
		Session.Contents.Remove("AsgLogin")
		
	End If

End If

'Logout
If Request.QueryString("Logout") = "True" Then Session.Contents.Remove("AsgLogin")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | powered by ASP Stats Generator <%= ASG_VERSION %></title>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta name="copyright" content="Copyright (C) 2003-2008 Carletti Simone, All Rights Reserved" />
<meta name="generator" content="ASP Stats Generator <%= ASG_VERSION %>" /> <!-- leave this for stats -->
<%	If Session("AsgLogin") = "Logged" AND Len(Request.QueryString("backto")) > 0 Then %>
<meta http-equiv="Refresh" content="3;url=<%= Request.QueryString("backto") %>">
<%	ElseIf Session("AsgLogin") = "Logged" AND NOT Len(Request.QueryString("backto")) > 0 Then %>
<meta http-equiv="Refresh" content="3;url=asg-default.asp">
<%	End If %>

<!--#include file="asg-includes/layout/head.asp" -->

<!--
  ASP Stats Generator (release <%= ASG_VERSION %>) is a free software package
  completely written in ASP programming language, for real time visitor tracking.
  Get your own copy for free at http://www.asp-stats-com/ !
-->

</head>

<!--#include file="asg-includes/layout/header.asp" -->
		<form action="login.asp?backto=<%= Server.URLEncode(Request.QueryString("backto")) %>" name="frmLogin" method="post">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableBarBgColour %>" valign="middle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" align="center" height="20" class="bartitle"><%= UCase(strAsgTxtLogin) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		<% If Session("AsgLogin") <> "Logged" Then %>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtEntryPassword) %></td>
		  </tr>
			  <% If blnAsgErrore Then %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" align="center" height="16"><br /><strong><%= strAsgTxtWrongPassword %></strong><br /><br /></td>		  
		  </tr>
			  <% End If %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="50%" align="right"><%= strAsgTxtTypePassword %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="50%" align="left">&nbsp;<input type="password" name="Password" value="" size="20" maxlength="20" /></td>
		  </tr><%
				
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(2)

		  %><tr class="normaltitle">
			<td colspan="2" align="center"><script>document.frmLogin.Password.focus()</script><br />
				<input type="hidden" name="Login" value="<%= strAsgTxtLogin %>" />
				<input type="submit" name="submit" value="<%= strAsgTxtLogin %>" />
			</td>
		  </tr>
		<% Else %>
		  <tr class="normaltitle" bgcolor="<%= strAsgSknTableTitleBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtEntryAllowed) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center" colspan="2"><br />
				<%= strAsgTxtLoginCompleted & "<br />" & strAsgTxtEntryAllowed %><br /><br />
				<%= strAsgTxtGoingToBeRedirected  %><br />
				<a href="<% If Len(Request.QueryString("backto")) > 0 Then Response.Write(Request.QueryString("backto")) Else Response.Write("statistiche.asp") %>" title="<%= strAsgTxtGoToPage %>" class="linksmalltext"><%= strAsgTxtClickToRedirect %></a><br /><br />
				<a href="login.asp?Logout=True" title="<%= strAsgTxtLogout %>" class="linksmalltext"><%= strAsgTxtClickToLogout %></a><br /><br />
			</td>
		  </tr><%
				
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(2)

		   End If %>
		</table>
		</form>
<%

' Footer
Response.Write(vbCrLf & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
'// Row - Footer Border Line
Call BuildFooterBorderLine()

' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write("<tr align=""center"" valign=""middle"">")
Response.Write("<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer""><a href=""http://www.asp-stats.com/"" class=""linkfooter"" title=""ASP Stats Generator Homepage"">ASP Stats Generator</a> [" & ASG_VERSION & "] - &copy; 2003-2008 <a href=""http://www.weppos.com/"" class=""linkfooter"" title=""Weppos.com Homepage"">weppos</a>")
if ASG_CONFIG_ELABTIME then Response.Write(asgElabtime())
Response.Write("</td>")
Response.Write("</tr>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write("</table>")

%>
<!--#include file="asg-includes/layout/footer.asp" -->

</body></html>
