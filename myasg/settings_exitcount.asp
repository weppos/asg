<%@LANGUAGE="VBSCRIPT"%>
<!--include virtual="/myasg/config.asp" -->
<!--#include file="config.asp" -->
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
<title><%= strAsgSiteName %> | ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2007 Carletti Simone, All Rights Reserved" />
<link href="stile.css" rel="stylesheet" type="text/css" />

<!-- 	ASP Stats Generator <%= strAsgVersion %> è una applicazione gratuita 
		per il monitoraggio degli accessi e dei visitatori ai siti web 
		creata e sviluppata da Simone Carletti.
		
		Puoi scaricarne una copia gratuita sul sito ufficiale http://www.weppos.com/ -->

</head>
<!--include virtual="/myasg/includes/header.asp" -->
<!--#include file="includes/header.asp" -->
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
Response.Write("<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer""><a href=""http://www.asp-stats.com/"" class=""linkfooter"" title=""ASP Stats Generator Homepage"">ASP Stats Generator</a> [" & strAsgVersion & "] - &copy; 2003-2007 <a href=""http://www.weppos.com/"" class=""linkfooter"" title=""Weppos.com Homepage"">weppos</a>")
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