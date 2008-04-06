<%@LANGUAGE="VBSCRIPT"%>
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
<title><%= strAsgSiteName %> | ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2008 Carletti Simone, All Rights Reserved" />
<link href="stile.css" rel="stylesheet" type="text/css" />

<!-- 	ASP Stats Generator <%= strAsgVersion %> è una applicazione gratuita 
		per il monitoraggio degli accessi e dei visitatori ai siti web 
		creata e sviluppata da Simone Carletti.
		
		Puoi scaricarne una copia gratuita sul sito ufficiale http://www.weppos.com/ -->

</head>
<!--#include file="includes/header.asp" -->
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
Response.Write("<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer""><a href=""http://www.asp-stats.com/"" class=""linkfooter"" title=""ASP Stats Generator Homepage"">ASP Stats Generator</a> [" & strAsgVersion & "] - &copy; 2003-2008 <a href=""http://www.weppos.com/"" class=""linkfooter"" title=""Weppos.com Homepage"">weppos</a>")
If blnAsgElabTime Then Response.Write(" - " & strAsgTxtThisPageWasGeneratedIn & "&nbsp;" & FormatNumber(Timer() - startAsgElab, 4) & "&nbsp;" & strAsgTxtSeconds)
Response.Write("</td>")
Response.Write("</tr>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write("</table>")
Response.Write("</td></tr>")
Response.Write("</table>")
Response.Write("</td></tr>")
Response.Write("</table>")

%>
<!-- footer -->
<!--#include file="includes/footer.asp" -->
</body></html>