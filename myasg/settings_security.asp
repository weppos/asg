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


Dim blnErrore

'Inserimento record
If Request.Form("Impostazioni") = strAsgTxtUpdate AND Request.QueryString("Exc") = "Upd" Then

	Dim strAsgPswNuova
	Dim strAsgPswConferma
	Dim blnAsgErrore
	Dim blnAsgPswIns
	
	blnErrore = False
	blnAsgPswIns = False
	
	strAsgPswNuova = Trim(Request.Form("PswNuova"))
	strAsgPswConferma = Trim(Request.Form("PswConferma"))
	
	If IsNumeric(Request.Form("Protezione")) Then intAsgSecurity = CInt(Request.Form("Protezione"))
	
	strAsgPswNuova = CleanInput(strAsgPswNuova)
	strAsgPswConferma = CleanInput(strAsgPswConferma)
	
	If "[]" & strAsgPswNuova <> "[]" Then
	
		If strAsgPswNuova = strAsgPswConferma Then
			blnAsgPswIns = true
		Else
			blnErrore = True
		End If
	
	End If
	
	'Nessun errore rilevato. Procedi con inserimento.
	If blnErrore = False Then
	
		If blnAsgPswIns = True Then
			strAsgSQL = "UPDATE "&strAsgTablePrefix&"Config SET Sito_PSW = '" & strAsgPswNuova & "', Stats_Protezione = " & intAsgSecurity & ""
		Else
			strAsgSQL = "UPDATE "&strAsgTablePrefix&"Config SET Stats_Protezione = " & intAsgSecurity & ""
		End If
		
		objAsgConn.Execute(strAsgSQL)
	
		'Se si utilizzano le variabili Application aggiornale
		If blnApplicationConfig Then
					
			'Aggiorna Variabili Application
			If blnAsgPswIns = True Then Application("strAsgSitePsw") = strAsgPswNuova
			Application("intAsgSecurity") = CInt(intAsgSecurity)
			'Forza il ricalcolo delle Application
			Application("blnConfig") = False
		
		End If
		
		'Reset Server Objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing
		
		'Reindirizza per rivalorizzare dati
		Response.Redirect("settings_security.asp?Msg=Upd")
	
	End If
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
		<form action="settings_security.asp?Exc=Upd" name="frmSicurezza" method="post">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= strAsgTxtSecuritySettings %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <% If Request.QueryString("Msg") = "Upd" Then %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" align="center" height="15"><br /><strong><%= strAsgTxtUpdateSuccessfullyCompleted %></strong><br /><br /></td>		  
		  </tr>
		  <% End If %>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtEntryPassword) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" align="center" height="15">
			<% 
			If blnErrore = True Then 
				Response.Write("<strong>" & strAsgTxtAttentionPasswordNotMatching & "</strong>") 
			Else 
			    Response.Write(strAsgTypeOnlyToChangePassword)
			End If %>
			</td>		  
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right" width="50%"><%= strAsgTxtNewPassword %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left" width="50%">&nbsp;<input type="password" name="PswNuova" value="" size="20" maxlength="20" /></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtConfirmPassword %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<input type="password" name="PswConferma" value="" size="20" maxlength="20" /></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtStatsProtection) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtStatsProtectionLevel %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">
				&nbsp;<input type="radio" name="Protezione" value="0" <% If intAsgSecurity = 0 Then Response.Write "checked" %> /><%= strAsgTxtNone %>
				&nbsp;<input type="radio" name="Protezione" value="1" <% If intAsgSecurity = 1 Then Response.Write "checked" %> /><%= strAsgTxtLimited %>
				&nbsp;<input type="radio" name="Protezione" value="2" <% If intAsgSecurity = 2 Then Response.Write "checked" %> /><%= strAsgTxtFull %>
			</td>
		  </tr>
		  <%
					
			'// Row - End table spacer			
			Call BuildTableContEndSpacer(2)
	
		  %>
		  <tr class="normaltitle">
			<td colspan="2" align="center"><br /><input type="submit" name="Impostazioni" value="<%= strAsgTxtUpdate %>" /></td>
		  </tr>
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
