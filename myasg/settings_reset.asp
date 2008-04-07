<%@LANGUAGE="VBSCRIPT"%>
<% Option Explicit %>
<!--#include file="config.asp" -->
<!--#include file="includes/inc_array_table.asp" -->
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


Dim aryAsgTableWarining(10,2)		'Holds the array containing warning informations


'Set warning record level, joining table id to the related included array id
aryAsgTableWarining(1,1) = 5000
aryAsgTableWarining(2,1) = 2000
aryAsgTableWarining(3,1) = 750
aryAsgTableWarining(4,1) = 2000
aryAsgTableWarining(5,1) = 500
aryAsgTableWarining(6,1) = 3500
aryAsgTableWarining(7,1) = 3500
aryAsgTableWarining(8,1) = 2000
aryAsgTableWarining(9,1) = 750
aryAsgTableWarining(10,1) = 3500


'Execute a loop to count the records of each table
For intAsgTableLoop = 1 to Ubound(aryAsgTable)
	
	'Initialise SQL string to count records
	strAsgSQL = "SELECT COUNT(*) FROM "&strAsgTablePrefix& aryAsgTable(intAsgTableLoop,1) & ""
	'Open Rs
	objAsgRs.Open strAsgSQL, objAsgConn
	'Set the number of total hits
	If Not objAsgRs.EOF Then 
		aryAsgTableWarining(intAsgTableLoop,2) = objAsgRs(0)
	Else
		aryAsgTableWarining(intAsgTableLoop,2) = 0
	End If
	'Close Rs
	objAsgRs.Close
	
	'Set warning alert
	If aryAsgTableWarining(intAsgTableLoop,2) > aryAsgTableWarining(intAsgTableLoop,1) Then
		aryAsgTableWarining(intAsgTableLoop,0) = True
	Else
		aryAsgTableWarining(intAsgTableLoop,0) = False
	End If

Next

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

<!--#include file="includes/html-head.asp" -->

<!--
  ASP Stats Generator (release <%= strAsgVersion %>) is a free software package
  completely written in ASP programming language, for real time visitor tracking.
  Get your own copy for free at http://www.asp-stats-com/ !
-->

</head>
<!--#include file="includes/header.asp" -->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtResetSettings) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="70%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <% If Request.QueryString("msg") = "error" Then %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" align="center" height="15"><br />
			<strong><%= strAsgTxtErrorOccured %><br />
			<%= strAsgTxtCheckTableMatching %></strong><br /><br />
			</td>		  
		  </tr>
		  <% End If %>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtTableReset) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableContBgColour %>" class="smalltext" align="center">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" height="30">
			<img src="<%= strAsgSknPathImage %>warning_icon.gif" border="0" alt="<%= strAsgTxtAdvice %>" align="absmiddle">&nbsp;<%= strAsgTxtTablesWithWarningIconNeedsReset %></td>
		  </tr>
		<%
		
		For intAsgTableLoop = 0 to Ubound(aryAsgTable)
		
		%>
		  <form action="settings_reset_execute.asp" method="get" name="frmReset">
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="100%" colspan="2" height="20">&nbsp;
				<input type="hidden" name="table" value="<%= intAsgTableLoop %>" />
				<%
					
					'Show an alert icon if the table need a cleaning
					If aryAsgTableWarining(intAsgTableLoop,0) Then
					Response.Write(vbCrLf & "<img src=""" & strAsgSknPathImage & "warning_icon.gif"" border=""0"" align=""absmiddle"">")
					End If
					
					'Write table title and description
					Response.Write(vbCrLf & "<span class=""notetext"">" & aryAsgTable(intAsgTableLoop, 1) & "</span>&nbsp;(" & aryAsgTableWarining(intAsgTableLoop,2) & "&nbsp;" & strAsgTxtRecords & ")&nbsp;-&nbsp;" & aryAsgTable(intAsgTableLoop, 2))
				
				%>
			</td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="60%" height="18" align="right">&nbsp;
				<%= strAsgTxtDataReset %>
				<select name="timerange" class="smallform">
					<option value="full"><%= strAsgTxtFull %></option>
				<% 
				'Nei casi riportati di sotto mostra una select
				'limitata a causa delle impostazioni limitate delle strtture
				If aryAsgTable(intAsgTableLoop, 1) <> "IP" Then %>
					<option value="0"><%= strAsgTxtOlderThan & " " & strAsgTxtCurrent & " " & strAsgTxtMonth %></option>
					<option value="1"><%= strAsgTxtOlderThan & " 1 " & strAsgTxtMonth %></option>
					<% For looptmp = 2 to 12 %>
					<option value="<%= looptmp %>"><%= strAsgTxtOlderThan & " " & looptmp & " " & strAsgTxtMonths %></option>
					<% Next 
				ElseIf aryAsgTable(intAsgTableLoop, 1) = "Detail" Then %>
					<option value="0"><%= strAsgTxtOlderThan & " " & strAsgTxtCurrent & " " & strAsgTxtWeek %></option>
					<option value="1"><%= strAsgTxtOlderThan & " 1 " & strAsgTxtWeek %></option>
					<% For looptmp = 2 to 12 %>
					<option value="<%= looptmp %>"><%= strAsgTxtOlderThan & " " & looptmp & " " & strAsgTxtWeeks %></option>
					<% Next 
				End If
				%>
				</select>&nbsp;&nbsp;
			</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="40%">&nbsp;
				<input type="image" src="images/delete.gif" name="delete" value="deletenormal" /></td>
		  </tr>
		  <% If aryAsgTable(intAsgTableLoop, 1) = "Detail" Then %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="60%" height="18" align="right">&nbsp;
				<%= strAsgTxtDataReset %>
				<select name="weekrange" class="smallform">
					<option value=""></option>
					<option value="0"><%= strAsgTxtOlderThan & " " & strAsgTxtCurrent & " " & strAsgTxtWeek %></option>
					<option value="1"><%= strAsgTxtOlderThan & " 1 " & strAsgTxtWeek %></option>
					<% For looptmp = 2 to 12 %>
					<option value="<%= looptmp %>"><%= strAsgTxtOlderThan & " " & looptmp & " " & strAsgTxtWeeks %></option>
					<% Next %>
				</select>&nbsp;&nbsp;
			</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="40%">&nbsp;
				<input type="image" src="images/delete.gif" name="delete" value="deleteweek" /></td>
		  </tr>
		  <% End If 'Condizione tabella details	%>
		  </form>
		<%
		
		Next  
		
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