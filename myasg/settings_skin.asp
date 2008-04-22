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


'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

Dim strAsgInput



'on error resume next
pathFrom = Server.MapPath(strAsgPathFolderWr & "inc_skin_file.asp")
pathFrom2 = Server.MapPath(strAsgPathFolderWr & "inc_skin_file_temp.asp")

Set objFso = CreateObject("Scripting.FileSystemObject")

objFso.CreateTextFile pathFrom2


if request.querystring = "update" then


	Set ts2 = objFso.OpenTextFile(pathFrom2, 2)
	Set ts = objFso.OpenTextFile(pathFrom, 1)


	Do While ts.AtEndOfStream <> True
			riga = ts.ReadLine

	if not left(riga,1)="'" AND mid(riga,2,1)<>"%" AND mid(riga,1,1)<>"%" then
	if riga<>"" then
	temp1=right(riga,(len(riga)-instr(riga,"'")+1))
'response.write "temp1 "&temp1&"<br>"
	temp2=left(riga,instr(riga,"=")-1)
'response.write "temp2 "&temp2&"<br>"
	temp3=trim(replace(temp2,"Const ",""))
'response.write "temp3 "&temp3&"<br>"
	strAsgInput = Request.Form(temp3)
	strAsgInput = Replace(strAsgInput, ":", "")
	ts2.writeline(temp2&" = """ & strAsgInput & """ "&temp1 )

	else
	ts2.writeline(riga)

	end if

	else
	ts2.writeline(riga)

	end if

	Loop
	ts.Close
	Set ts = Nothing
	ts2.Close
	Set ts2 = Nothing

objFso.deleteFile pathFrom
objFso.CopyFile pathFrom2, pathFrom
objFso.deleteFile pathFrom2

end if


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
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtSkinSettings) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<form method="post" name="skin" action="settings_skin.asp?update">
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" align="center" height="16" colspan="2"><%= UCase(strAsgTxtSkinSettings) %></td>
		  </tr>
		<%
		
		Set ts = objFso.OpenTextFile(pathFrom, 1)

		Do While ts.AtEndOfStream <> True
			
			riga = ts.ReadLine
		
		if not left(riga,1)="'" AND mid(riga,2,1)<>"%" AND mid(riga,1,1)<>"%" then
		temp=right(riga,(len(riga)-instr(riga,"=")))
	
		if temp <> "" then
		
		%>	
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="25%" align="right"><input type="text" name="<%=replace(trim(left(riga,instr(riga,"=")-1)),"Const ","")%>" value="<%=trim(replace(left(temp,instr(temp,"'")-1),"""",""))%>" />&nbsp;</td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="75%">&nbsp;<%=replace(right(temp,(len(temp)-instr(temp,"'"))),"""","")%></td>
          </tr>
		<%
		
		end if
	
		end if
	
		Loop
		ts.Close
		Set ts = Nothing
	
		Set objFso = Nothing

		'// Row - End table spacer			
		Call BuildTableContEndSpacer(2)
	
		%>
		  <tr class="smalltext">
            <td colspan="2" align="center"><br /><input type="submit" value="<%= strAsgTxtUpdate %>" name="invia" /></td>
          </tr>
		</table><br />
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
if blnAsgElabTime then Response.Write(asgElabtime())
Response.Write("</td>")
Response.Write("</tr>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write("</table>")

%>
<!--#include file="asg-includes/layout/footer.asp" -->

</body></html>
