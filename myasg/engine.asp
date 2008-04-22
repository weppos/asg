<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
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
Call AllowEntry("True", "False", "False", intAsgSecurity)


'Dichiara Variabili
Dim mese			'Riferimento per output
Dim elenca			'Tutti | Mese
Dim dettagli		'Motore di cui mostrare le informazioni


'Grafico
Dim intAsgLarColMax		'Larghezza massima in px delle colonne dipendente dalla pag

Dim intAsgMax				'Valore massimo di pagine visitate
Dim intAsgParte

Dim intAsgTotMeseHits		'Valore totale per mese di pagine visitate
Dim intAsgTotMeseVisits		'Valore totale per mese di accessi unici


'Read setting variables from querystring
mese = Request.QueryString("mese")
elenca = Request.QueryString("elenca")
dettagli = Request.QueryString("dettagli")
strAsgSortBy = Request.QueryString("sort")
strAsgSortOrder = Request.QueryString("order")


'If period variable is empty then set it to the current month
If mese = "" Then mese = dtmAsgMonth & "-" & dtmAsgYear
'If the variable is empty set it to monthly report
If elenca = "" Then elenca = "mese"
' Set the sorting order depending on querystring
if strAsgSortOrder = "ASC" then 
	strAsgSortOrder = "ASC"
else
	strAsgSortOrder = "DESC"
end if
'If a time period has been chosen then build the variable to query the database
If Request.QueryString("showperiod") = strAsgTxtShow Then mese = Request.QueryString("periodmm") & "-" & Request.QueryString("periodyy")


'Set max total column width
intAsgLarColMax = 300				'Largezza massima colonne | Rapportata alla dimensione della pagina


'Richiama totale
If elenca = "mese" Then 
	strAsgSQL = "SELECT Sum(Hits) As SumHits, Sum(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Query WHERE Mese = '" & mese & "' "
ElseIf elenca = "tutti" Then 
	strAsgSQL = "SELECT Sum(Hits) As SumHits, Sum(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Query "
End If

objAsgRs.Open strAsgSQL, objAsgConn
If objAsgRs.EOF Then
	intAsgTotMeseHits = 0
	intAsgTotMeseVisits = 0
Else
	intAsgTotMeseHits = objAsgRs("SumHits")
	intAsgTotMeseVisits = objAsgRs("SumVisits")
End If
objAsgRs.Close
'Accertati che non siano nulli
If intAsgTotMeseHits = 0 OR "[]" & intAsgTotMeseHits = "[]" Then intAsgTotMeseHits = 0
If intAsgTotMeseVisits = 0 OR "[]" & intAsgTotMeseVisits = "[]" Then intAsgTotMeseVisits = 0


'Richiama valore Massimo
If elenca = "mese" Then 
	strAsgSQL = "SELECT SUM(Hits) AS SumHits FROM "&strAsgTablePrefix&"Query WHERE Mese = '" & mese & "' GROUP BY Engine ORDER BY SUM(Hits) DESC"
ElseIf elenca = "tutti" Then 
	strAsgSQL = "SELECT SUM(Hits) AS SumHits FROM "&strAsgTablePrefix&"Query GROUP BY Engine ORDER BY SUM(Hits) DESC"
End If
objAsgRs.Open strAsgSQL, objAsgConn, 2, 3
If objAsgRs.EOF Then
	intAsgMax = 0
Else
	objAsgRs.MoveFirst
	intAsgMax = objAsgRs("SumHits")
End If
objAsgRs.Close

'Calcola unità singola
If intAsgMax = 0 OR "[]" & intAsgMax = "[]" Then intAsgMax = 1
intAsgParte = intAsgLarColMax/intAsgMax


'Read sorting order from querystring
'// Filter QS values and associate them 
'// with their respective database fields
Select Case strAsgSortBy
	Case "hts" strAsgSortByFld = "SUM(Hits)"
	Case "visits" strAsgSortByFld = "SUM(Visits)"
	Case "engine" strAsgSortByFld = "Engine"
	Case Else strAsgSortByFld = "SUM(Visits)"
End Select

'Richiama le Dichiarazioni per la 
'paginazione avanzata
Call DimPaginazioneAvanzata()

'Richiama le Dichiarazioni per la 
'paginazione avanzata dei dettagli
Call DimPaginazioneAvanzataDettagli()

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
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtSearchEngine) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td width="5%"  background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"></td>
			<td width="28%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtEngine) %>
				<a href="engine.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&dettagli=" & dettagli & "&sort=engine&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtEngine & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="engine.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&dettagli=" & dettagli & "&sort=engine&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtEngine & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="12%"  background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmHits) %></td>
			<td width="50%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>">
				<a href="engine.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&dettagli=" & dettagli & "&sort=hits&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtHits & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="engine.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&dettagli=" & dettagli & "&sort=hits&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtHits & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a>
			&nbsp;&nbsp;<%= UCase(strAsgTxtGraph) %>&nbsp;&nbsp;
				<a href="engine.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&dettagli=" & dettagli & "&sort=visits&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtVisits & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="engine.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&dettagli=" & dettagli & "&sort=visits&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtVisits & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="5%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"></td>
		  </tr>
		<%

		If elenca = "mese" Then 
			'Initialise SQL string to select data
			strAsgSQL = "SELECT SUM(Hits) AS SumHits, SUM(Visits) As SumVisits, Engine FROM "&strAsgTablePrefix&"Query WHERE Mese = '" & mese & "' "
			'Call the function to search into the database if there are enought information to do that
			strAsgSQL = CheckSearchForData(strAsgSQL, false)
		ElseIf elenca = "tutti" Then 
			'Initialise SQL string to select data
			strAsgSQL = "SELECT SUM(Hits) AS SumHits, SUM(Visits) As SumVisits, Engine FROM "&strAsgTablePrefix&"Query "
			'Call the function to search into the database if there are enought information to do that
			strAsgSQL = CheckSearchForData(strAsgSQL, true)
		End If
		
		'Group information by following fields
		strAsgSQL = strAsgSQL & " GROUP BY Engine "
		'Order record by selected field 
		strAsgSQL = strAsgSQL & " ORDER BY " & strAsgSortByFld & " " & strAsgSortOrder & ""
		
		'Prepara il Rs
		objAsgRs.CursorType = 3
		objAsgRs.LockType = 3
		
		'Apri il Rs
		objAsgRs.Open strAsgSQL, objAsgConn
			
			'Il Rs è vuoto
			If objAsgRs.EOF Then
				
				'If it is a search query then show no results advise
				If Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 Then
	
					'// Row - No current record	for search terms		
					Call BuildTableContNoRecord(5, "search")
					
				'Else show general no record information
				Else
	
					'// Row - No current record			
					Call BuildTableContNoRecord(5, "standard")
					
				End If
				
			Else
			
				objAsgRs.PageSize = RecordsPerPage
				objAsgRs.AbsolutePage = page
				
				For PaginazioneLoop = 1 To RecordsPerPage
					
					If Not objAsgRs.EOF Then			
					
		%>		  
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><img src="images/engine.asp?icon=<%= objAsgRs("Engine") %>" alt="<%= objAsgRs("Engine") %>" class="def-icon def-icon-searchengine" /></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left"><%
											
				'Write an anchor
				Response.Write(vbCrLf & "<a name=""" & objAsgRs("Engine") & """></a>")
			
			%>
			<a href="engine.asp?<%= "dettagli=" & objAsgRs("Engine") & "&page=" & page & "&mese=" & mese & "&elenca=" & elenca & "&searchfor=" & asgSearchfor & "&searchin=" & asgSearchin & "&sort=" & strAsgSortBy & "&order=" & strAsgSortOrder & "#" & objAsgRs("Engine") & "" %>" title="<%= strAsgTxtSearchQuery%>">
			<% 
				'Icona espansa se Corrisponde
				If Trim(dettagli) <> "" AND objAsgRs("Engine") = Trim(dettagli) Then
					Response.Write("<img src=""" & strAsgSknPathImage & "expanded.gif"" alt=""" & strAsgTxtSearchQuery & """ border=""0"" align=""absmiddle"" />")
				'Icona espandi se Differente
				Else
					Response.Write("<img src=""" & strAsgSknPathImage & "expand.gif"" alt=""" & strAsgTxtSearchQuery & """ border=""0"" align=""absmiddle"" />")
				End If
			
			%></a>&nbsp;<%= HighlightSearchKey(objAsgRs("Engine"), "Engine") %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= objAsgRs("SumHits") & "<br />" & objAsgRs("SumVisits") %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>">
				<img src="<%= strAsgSknPathImage%>bar_graph_image_hits_h.gif" width="<%= FormatNumber(objAsgRs("SumHits")*intAsgParte, 2) %>" height="9" alt="<%= strAsgTxtHits %>" align="absmiddle" /> [<%= CalcolaPercentuale(intAsgTotMeseHits, objAsgRs("SumHits")) %>]<br />
				<img src="<%= strAsgSknPathImage%>bar_graph_image_visits_h.gif" width="<%= FormatNumber(objAsgRs("SumVisits")*intAsgParte, 2) %>" height="9" alt="<%= strAsgTxtVisits %>" align="absmiddle" /> [<%= CalcolaPercentuale(intAsgTotMeseVisits, objAsgRs("SumVisits")) %>]
			</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"></td>
		  </tr>
		<% 
			If Trim(dettagli) <> "" AND objAsgRs("Engine") = Trim(dettagli) Then
				
				Dim objAsgRs2
				
				'Mostra le query al motore
				Set objAsgRs2 = Server.CreateObject("ADODB.Recordset")
				
				If elenca = "mese" Then 
					strAsgSQL = "SELECT Query, Hits, Visits, SERP_Page FROM "&strAsgTablePrefix&"Query WHERE Engine = '" & dettagli & "' AND Mese = '" & mese & "' "
				ElseIf elenca = "tutti" Then 
					strAsgSQL = "SELECT Query, Hits, Visits, SERP_Page FROM "&strAsgTablePrefix&"Query WHERE Engine = '" & dettagli & "' "
				End If
				
				strAsgSQL = strAsgSQL & " ORDER BY Visits DESC, Hits DESC"
		
		
		%>
		  <tr class="smalltext">
			<td colspan="5"><br />
				<!-- Contenitore Dettagli -->
				<table width="100%" border="0" cellspacing="0" cellpadding="1" align="center">
				  <tr>
					<td width="7%" valign="top" align="center"><img src="<%= strAsgSknPathImage %>openarrow.gif" width="25" height="25" border="0" alt="<%= strAsgTxtDetails %>"></td>
					<td width="86%">
					<!-- Dettagli Query Motore -->
					<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
					  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
						<td width="76%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"><%= UCase(strAsgTxtEngine) %></td>
						<td width="12%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmHits) %></td>
						<td width="12%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmVisits) %></td>
					  </tr>
					  <% 
					
					'Prepara il Rs
					objAsgRs2.CursorType = 3
					objAsgRs2.LockType = 3
					
					'Apri il Rs
					objAsgRs2.Open strAsgSQL, objAsgConn
						
						'Il Rs è vuoto
						If objAsgRs2.EOF Then
							
							'// Row - No current record			
							Call BuildTableContNoRecord(3, "standard")
							
						Else

							objAsgRs2.PageSize = detRecordsPerPage
							objAsgRs2.AbsolutePage = detpage
							
							For detPaginazioneLoop = 1 To detRecordsPerPage
								If Not objAsgRs2.EOF Then			

					  %>
					  <tr bgcolor="<%= strAsgSknTableContBgColour %>" class="smalltext">
						<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left"><% If objAsgRs2("SERP_Page") > 0 Then Response.Write("&nbsp;<a href=""serp.asp?serp=" & objAsgRs2("SERP_Page") & "&mese=" & mese &"&elenca=" & elenca & """ title=""" & strAsgTxtQuery & "&nbsp;" & objAsgRs2("SERP_Page") & "&deg;&nbsp;" & strAsgTxtPage & """><span class=""notetext"">[" & objAsgRs2("SERP_Page") & "]</span></a>") %>&nbsp;<%= objAsgRs2("Query") %></td>
						<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= objAsgRs2("Hits") %></td>
						<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= objAsgRs2("Visits") %></td>
					  </tr>
					  <%
						
								objAsgRs2.MoveNext
								End If
							Next
							End If
								
						'// Row - End table spacer			
						Call BuildTableContEndSpacer(3)
				
						'// Row - Advanced details data sorting
						Response.Write(vbCrLf & "<tr class=""smalltext""><td colspan=""3"" align=""center""><br />")
						Call PaginazioneAvanzataDettagli("engine.asp", "")
						Response.Write(vbCrLf & "<br /><br /></td></tr>")
									  
						objAsgRs2.Close
						Set objAsgRs2 = Nothing
					  
					  %>
					</table><br />
					<!-- Fine Dettagli Query Motore -->
					</td>
					<td width="7%"></td>
				  </tr>
				</table>
				<!-- Fine Contenitore Dettagli -->
			</td>
		  </tr>
		<%
			'Fine condizione dettagli
			End If

				objAsgRs.MoveNext
				End If
			Next
			End If
				
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(5)

		'// Row - Advanced data sorting
		Response.Write(vbCrLf & "<tr class=""smalltext""><td colspan=""5"" align=""center""><br />")
		Call PaginazioneAvanzata("engine.asp", "")
		Response.Write(vbCrLf & "<br /><br /></td></tr>")

		objAsgRs.Close

		'Reset Server Objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing

		'// Row - Data output panels
		Response.Write(vbCrLf & "<tr class=""smalltext"" align=""center"" valign=""top""><td colspan=""5"" height=""25""><br />")
		Call GoToPeriod("engine.asp", "")
		Call GoToGrouping("engine.asp", "")
		Call SearchForData("engine.asp", "", "Engine")
		Response.Write(vbCrLf & "</td></tr>")
		
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
if blnAsgElabTime then Response.Write(asgElabtime())
Response.Write("</td>")
Response.Write("</tr>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write("</table>")

%>
<!--#include file="asg-includes/layout/footer.asp" -->

</body></html>
