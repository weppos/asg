<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="config.asp" -->
<!--#include file="asg-lib/vbscript.asp" -->
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
Call AllowEntry("True", "True", "False", intAsgSecurity)


'Faccio prima a richiamare la stessa funzione in config
'e passargli i nuovi parametri!
Call FormatInTimeZone(dtmAsgNow, aryAsgTimeZone(1))


'Dichiara Variabili
Dim dtmStsDataOggi	
Dim dtmStsDataIeri
Dim dtmStsMeseCorrente
Dim dtmStsMeseScorso
'Valori
Dim intStsAccessiOggi
Dim intStsAccessiIeri
Dim intStsAccessiMeseCorrente
Dim intStsAccessiMeseScorso
Dim intStsPagineOggi
Dim intStsPagineIeri
Dim intStsPagineMeseCorrente
Dim intStsPagineMeseScorso
Dim intAsgUsersOnline				'Number of Users online on the site


dtmStsDataOggi = dtmAsgDate
dtmStsDataIeri = DateAdd("d", -1, dtmAsgDate) : dtmStsDataIeri = Year(dtmStsDataIeri) & "/" & Month(dtmStsDataIeri) & "/" & Day(dtmStsDataIeri)
dtmStsMeseCorrente = Month(dtmAsgDate)
dtmStsMeseScorso = Month(DateAdd("m", -1, dtmAsgDate))
	if Len(dtmStsMeseCorrente) < 2 Then dtmStsMeseCorrente = "0" & dtmStsMeseCorrente
	dtmStsMeseCorrente = dtmStsMeseCorrente & "-" & dtmAsgYear
	if Len(dtmStsMeseScorso) < 2 Then dtmStsMeseScorso = "0" & dtmStsMeseScorso
	' Last year
	if Cint(dtmStsMeseScorso) = 12 then 
		dtmStsMeseScorso = dtmStsMeseScorso & "-" & (dtmAsgYear - 1)
	else
		dtmStsMeseScorso = dtmStsMeseScorso & "-" & dtmAsgYear
	end if


'Oggi
strAsgSQL = "SELECT Hits, Visits FROM "&strAsgTablePrefix&"Daily WHERE Data = #" & dtmStsDataOggi & "#"
objAsgRs.Open strAsgSQL, objAsgConn
	If NOT objAsgRs.EOF Then
		intStsAccessiOggi = objAsgRs("Visits")
		intStsPagineOggi = objAsgRs("Hits")
	Else
		intStsAccessiOggi = 0
		intStsPagineOggi = 0
	End If
objAsgRs.Close

'Ieri
strAsgSQL = "SELECT Hits, Visits FROM "&strAsgTablePrefix&"Daily WHERE Data = #" & dtmStsDataIeri & "#"
objAsgRs.Open strAsgSQL, objAsgConn
	If NOT objAsgRs.EOF Then
		intStsAccessiIeri = objAsgRs("Visits")
		intStsPagineIeri = objAsgRs("Hits")
	Else
		intStsAccessiIeri = 0
		intStsPagineIeri = 0
	End If
objAsgRs.Close

'Mese Corrente
strAsgSQL = "SELECT SUM(Hits) AS SumHits, SUM(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Daily WHERE Mese = '" & dtmStsMeseCorrente & "' GROUP BY Mese "
objAsgRs.Open strAsgSQL, objAsgConn
	If NOT objAsgRs.EOF Then
		intStsAccessiMeseCorrente = objAsgRs("SumVisits")
		intStsPagineMeseCorrente = objAsgRs("SumHits")
	Else
		intStsAccessiMeseCorrente = 0
		intStsPagineMeseCorrente = 0
	End If
objAsgRs.Close

'Mese Scorso
strAsgSQL = "SELECT SUM(Hits) AS SumHits, SUM(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Daily WHERE Mese = '" & dtmStsMeseScorso & "' GROUP BY Mese "
'Response.Write(strAsgSQL)
objAsgRs.Open strAsgSQL, objAsgConn
	If NOT objAsgRs.EOF Then
		intStsAccessiMeseScorso = objAsgRs("SumVisits")
		intStsPagineMeseScorso = objAsgRs("SumHits")
	Else
		intStsAccessiMeseScorso = 0
		intStsPagineMeseScorso = 0
	End If
objAsgRs.Close

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
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtVisitsInformations) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />

		<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr valign="top"><td width="48%">

		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
          <tr bgcolor="<%= strAsgSknTableTitleBgColour %>"class="normaltitle">
            <td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtGeneralInformations) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="70%"><span class="notetext"><%= strAsgTxtHits & "&nbsp;" & strAsgTxtToday %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%"><%= intStsPagineOggi %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext"><%= strAsgTxtVisits & "&nbsp;" & strAsgTxtToday %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= intStsAccessiOggi %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtHits & "&nbsp;" & strAsgTxtYesterday %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= intStsPagineIeri %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtVisits & "&nbsp;" & strAsgTxtYesterday %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= intStsAccessiIeri %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext"><%= strAsgTxtHits & "&nbsp;" & strAsgTxtCurrent & "&nbsp;" & strAsgTxtMonth %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= intStsPagineMeseCorrente %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext"><%= strAsgTxtVisits & "&nbsp;" & strAsgTxtCurrent & "&nbsp;" & strAsgTxtMonth %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= intStsAccessiMeseCorrente %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtHits & "&nbsp;" & strAsgTxtLastMonth %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= intStsPagineMeseScorso %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtVisits & "&nbsp;" & strAsgTxtLastMonth %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= intStsAccessiMeseScorso %></td>
          </tr>
		  <%
			  
			'// Row - End table spacer			
			Call BuildTableContEndSpacer(2)
			  
		  %>
		</table><br />
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
          <tr bgcolor="<%= strAsgSknTableTitleBgColour %>"class="normaltitle">
            <td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtYearlyInformations) %></td>
          </tr>
<%

'Annuali
strAsgSQL = "SELECT * FROM "&strAsgTablePrefix&"Counter ORDER BY Anno DESC "
objAsgRs.Open strAsgSQL, objAsgConn

If objAsgRs.EOF Then

%>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="70%"><%= strAsgTxtHits & " <span class=""notetext"">" & dtmAsgYear & "</span>" %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%"><%= strAsgStartHits %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtVisits & " <span class=""notetext"">" & dtmAsgYear & "</span>" %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgStartVisits %></td>
          </tr>
<%

Else

	Do While Not objAsgRs.EOF
%>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="70%"><%= strAsgTxtHits & " <span class=""notetext"">" & objAsgRs("Anno") & "</span>" %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%"><% Response.Write(objAsgRs("Hits") + strAsgStartHits) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtVisits & " <span class=""notetext"">" & objAsgRs("Anno") & "</span>" %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><% Response.Write(objAsgRs("Visits") + strAsgStartVisits) %></td>
          </tr>
<%
	objAsgRs.MoveNext
	Loop

End If
	
objAsgRs.Close

					
			'// Row - End table spacer			
			Call BuildTableContEndSpacer(2)
	
%>		  
		</table><br />
		
		</td><td width="4%" >
		</td><td width="48%">
		
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
            <td colspan="2" align="center" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"><%= UCase(strAsgTxtGeneralAverageInformations) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="70%"><span class="notetext"><%= strAsgTxtHits & "&nbsp;" & strAsgTxtToday & "&nbsp;" & strAsgTxtPerHour %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%"><%= MediaGiorno(intStsPagineOggi, 0, 1) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext"><%= strAsgTxtVisits & "&nbsp;" & strAsgTxtToday  & "&nbsp;" & strAsgTxtPerHour %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaGiorno(intStsAccessiOggi, 0, 1) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtHits & "&nbsp;" & strAsgTxtYesterday & "&nbsp;" & strAsgTxtPerHour %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaGiorno(intStsPagineIeri, 0, 2) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtVisits & "&nbsp;" & strAsgTxtYesterday & "&nbsp;" & strAsgTxtPerHour %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaGiorno(intStsAccessiIeri, 0, 2) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext"><%= strAsgTxtHits & "&nbsp;" & strAsgTxtCurrent & "&nbsp;" & strAsgTxtMonth & "&nbsp;" & strAsgTxtPerHour %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaMese(intStsPagineMeseCorrente, 1, 1) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext"><%= strAsgTxtVisits & "&nbsp;" & strAsgTxtCurrent & "&nbsp;" & strAsgTxtMonth & "&nbsp;" & strAsgTxtPerHour %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaMese(intStsAccessiMeseCorrente, 1, 1) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtHits & "&nbsp;" & strAsgTxtLastMonth & "&nbsp;" & strAsgTxtPerHour %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaMese(intStsPagineMeseScorso, 1, 2) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtVisits & "&nbsp;" & strAsgTxtLastMonth & "&nbsp;" & strAsgTxtPerHour %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaMese(intStsAccessiMeseScorso, 1, 2) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext"><%= strAsgTxtHits & "&nbsp;" & strAsgTxtCurrent & "&nbsp;" & strAsgTxtMonth & "&nbsp;" & strAsgTxtPerDay %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaMese(intStsPagineMeseCorrente, 2, 1) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext"><%= strAsgTxtVisits & "&nbsp;" & strAsgTxtCurrent & "&nbsp;" & strAsgTxtMonth & "&nbsp;" & strAsgTxtPerDay %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaMese(intStsAccessiMeseCorrente, 2, 1) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtHits & "&nbsp;" & strAsgTxtLastMonth & "&nbsp;" & strAsgTxtPerDay %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaMese(intStsPagineMeseScorso, 2, 2) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= strAsgTxtVisits & "&nbsp;" & strAsgTxtLastMonth & "&nbsp;" & strAsgTxtPerDay %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= MediaMese(intStsAccessiMeseScorso, 2, 2) %></td>
          </tr>
		  <%
					
			'// Row - End table spacer			
			Call BuildTableContEndSpacer(2)
	
		  %>
		</table><br />
		
		</td></tr>
		<tr valign="top"><td width="100%" colspan="3"><br />

<%

Dim dtmOnlineTime
dtmOnlineTime = DateAdd("n", -10, dtmAsgNow) : dtmOnlineTime = Year(dtmOnlineTime) & "/" & Month(dtmOnlineTime) & "/" & Day(dtmOnlineTime) & " " & Hour(dtmOnlineTime) & "." & Minute(dtmOnlineTime) & "." & Second(dtmOnlineTime)


'Query di conteggio degli utenti online
'strAsgSQL = "SELECT COUNT(*) FROM "&strAsgTablePrefix&"Detail WHERE Data > #" & dtmOnlineTime & "# GROUP BY Visitor_ID"

'Prepara il Rs
'objAsgRs.CursorType = 1
'objAsgRs.LockType = 3

'Apri il Rs
'objAsgRs.Open strAsgSQL, objAsgConn
'If objAsgRs.EOF Then
'	intAsgUsersOnline = 0
'Else
'	intAsgUsersOnline = objAsgRs(0)
'End If
'objAsgRs.Close


'Query di richiamo degli utenti
strAsgSQL = "SELECT Visitor_ID, LAST(IP) AS LastIP, LAST(Data) AS LastData, LAST(Page) AS LastPage FROM "&strAsgTablePrefix&"Detail WHERE Data > #" & dtmOnlineTime & "# GROUP BY Visitor_ID"
		
'Prepara il Rs
objAsgRs.CursorType = 3
objAsgRs.LockType = 3
		
'Apri il Rs
objAsgRs.Open strAsgSQL, objAsgConn
		
	'Mostra solo se ci sono utenti online
	If Not objAsgRs.Eof Then

%>		
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
          <tr bgcolor="<%= strAsgSknTableTitleBgColour %>"class="normaltitle">
            <td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="4" align="center" height="16"><%= objAsgRs.RecordCount & "&nbsp;" & UCase(strAsgTxtOnlineUsers) %></td>
          </tr>
<%
		Do While Not objAsgRs.EOF
%>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="18%" align="left">
<% 

			'Show icon only if security level = 0
			If intAsgSecurity = 0 OR Session("AsgLogin") = "Logged" Then
						
				'Tracking IP
				'// Link PopUp
				Response.Write(vbCrLf & "            <a href=""JavaScript:openWin('popup_tracking_ip.asp?IP=" & objAsgRs("LastIP") & "','Tracking','toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=425')"" title=""" & strAsgTxtIPTracking & """>")
				'// Icona espansa se Corrisponde
				Response.Write(vbCrLf & "            <img src=""" & strAsgSknPathImage & "tracking_small.gif"" alt=""" &  strAsgTxtIPTracking & """ border=""0"" align=""absmiddle"" />")
				'// Chiudi Link PopUp
				Response.Write("</a>")
				
			End If

			'Mostra solo se Loggato
			If Session("AsgLogin") = "Logged" Then
							
				'Icona Filter IP
				Call ShowIconFilterIp(objAsgRs("LastIP"))
							
			End If
			
			'Stampa l'IP
			Response.Write("&nbsp;" & objAsgRs("LastIP"))

%>
			</td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="20%" align="center"><%= FormatOutTimeZone(objAsgRs("LastData"), "Date") & "&nbsp;" & strAsgTxtAt & "&nbsp;" & FormatOutTimeZone(objAsgRs("LastData"), "Time") %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="55%">
				<a class="linksmalltext" href="<%= objAsgRs("LastPage") %>" target="_blank" title="<%= objAsgRs("LastPage") %>">
				<%= StripValueTooLong(objAsgRs("LastPage"), 65, 30, 30) %></a>
			</td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="7%"></td>
          </tr>
<%
		'Cicla
		objAsgRs.MoveNext
		Loop 
		
%>
		</table><br />
<%
	
	'/Condizione utenti online
	End If

objAsgRs.Close

'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

%>
		
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
          <tr bgcolor="<%= strAsgSknTableTitleBgColour %>"class="normaltitle">
            <td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="4" align="center" height="16"><%= UCase(strAsgTxtServerInformations) %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="20%"><span class="notetext"><%= strAsgTxtIISversion %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%"><%= Request.ServerVariables("SERVER_SOFTWARE") %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="20%"><span class="notetext"><%= strAsgTxtServerName %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%"><%= Request.ServerVariables("SERVER_NAME") %></td>
          </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext"><%= strAsgTxtProtocolVersion %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= Request.ServerVariables("SERVER_PROTOCOL") %></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext">VBScript Engine</span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%= asgGetScriptEngineInfo() %></td>
          </tr>
		  <%

			'Link a completo se loggato
			If Session("AsgLogin") = "Logged" Then

		  %>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><span class="notetext"><%= strAsgTxtYourIpIs %></span></td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><%

				'Filter IP
				'// Link PopUp
				Response.Write(vbCrLf & "						<a href=""JavaScript:openWin('popup_filter_ip.asp?IP=" & Request.ServerVariables("REMOTE_ADDR") & "','Filter','toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=200')"" title=""" & strAsgTxtFilterIPaddr & """>" & Request.ServerVariables("REMOTE_ADDR"))
				'// Chiudi Link PopUp
				Response.Write("</a>") 

				'Icona Filter IP
				Call ShowIconFilterIp(Request.ServerVariables("REMOTE_ADDR"))

			%>
			</td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" colspan="2" align="right">
				<a href="check_server.asp" title="<%= strAsgTxtServerInformations %>" class="linksmalltext"><%= strAsgTxtFullVersion %> <img src="images/arrow_small_dx.gif" alt="<%= strAsgTxtServerInformations %>" align="absmiddle" border="0" /></a>
			</td>
          </tr>
		  <%

			End If

			'// Row - End table spacer			
			Call BuildTableContEndSpacer(4)
	
		  %>
		</table><br />

		</td></tr>
		</table>

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
