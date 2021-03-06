<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="config.asp" -->
<!--#include file="includes/inc_array_table.asp" -->
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
				<input type="image" src="images/delete.gif" name="delete" value="deletenormal" onclick="return confirm('Are you sure you want to delete selected records?');" /></td>
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
				<input type="image" src="images/delete.gif" name="delete" value="deleteweek" onclick="return confirm('Are you sure you want to delete selected records?');" /></td>
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
