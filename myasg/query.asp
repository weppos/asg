<%@LANGUAGE="VBSCRIPT"%>
<% Option Explicit %>
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
Dim gruppo			'Raggruppa per Motori | Non raggruppare per Motori 
Dim intAsgCount		'Conteggio record

Dim intNumColspan


'Grafico
Dim intAsgLarColMax		'Larghezza massima in px delle colonne dipendente dalla pag

Dim intAsgMax				'Valore massimo di pagine visitate
Dim intAsgParte

Dim intAsgTotMeseHits		'Valore totale per mese di pagine visitate
Dim intAsgTotMeseVisits		'Valore totale per mese di accessi unici


'Read setting variables from querystring
mese = Request.QueryString("mese")
elenca = Request.QueryString("elenca")
gruppo = Request.QueryString("gruppo")
strAsgSortBy = Request.QueryString("sort")
strAsgSortOrder = Request.QueryString("order")


'If period variable is empty then set it to the current month
If mese = "" Then mese = dtmAsgMonth & "-" & dtmAsgYear
'If the variable is empty set it to monthly report
If elenca = "" Then elenca = "mese"
'If the variable is empty set engine grouping
If gruppo = "" Then gruppo = "motori"
' Set the sorting order depending on querystring
if strAsgSortOrder = "ASC" then 
	strAsgSortOrder = "ASC"
else
	strAsgSortOrder = "DESC"
end if
'If a time period has been chosen then build the variable to query the database
If Request.QueryString("showperiod") = strAsgTxtShow Then mese = Request.QueryString("periodmm") & "-" & Request.QueryString("periodyy")


'Set max total column width
If gruppo = "motori" Then 
	intAsgLarColMax = 200				'Largezza massima colonne | Rapportata alla dimensione della pagina
ElseIf gruppo = "query" Then 
	intAsgLarColMax = 200				'Largezza massima colonne | Rapportata alla dimensione della pagina
End If


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
If elenca = "mese" AND gruppo = "motori" Then 
	strAsgSQL = "SELECT MAX(Hits) AS SumHits FROM "&strAsgTablePrefix&"Query WHERE Mese = '" & mese & "' ORDER BY MAX(Hits) DESC "
ElseIf elenca = "tutti" AND gruppo = "motori" Then 
	strAsgSQL = "SELECT MAX(Hits) AS SumHits FROM "&strAsgTablePrefix&"Query ORDER BY MAX(Hits) DESC"
ElseIf elenca = "mese" AND gruppo = "query" Then 
	strAsgSQL = "SELECT SUM(Hits) AS SumHits FROM "&strAsgTablePrefix&"Query WHERE Mese = '" & mese & "' GROUP BY Query ORDER BY SUM(Hits) DESC "
ElseIf elenca = "tutti" AND gruppo = "query" Then 
	strAsgSQL = "SELECT SUM(Hits) AS SumHits FROM "&strAsgTablePrefix&"Query GROUP BY Query ORDER BY SUM(Hits) DESC "
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
If gruppo = "motori" Then
	Select Case strAsgSortBy
		Case "Hits" strAsgSortByFld = "Hits"
		Case "Visits" strAsgSortByFld = "Visits"
		Case "Query" strAsgSortByFld = "Query"
		Case "Engine" strAsgSortByFld = "Engine"
		Case Else strAsgSortByFld = "Visits"
	End Select
	
	intNumColspan = 5
	
ElseIf gruppo = "query" Then
	Select Case strAsgSortBy
		Case "Hits" strAsgSortByFld = "SUM(Hits)"
		Case "Visits" strAsgSortByFld = "SUM(Visits)"
		Case "Query" strAsgSortByFld = "Query"
		Case Else strAsgSortByFld = "SUM(Visits)"
	End Select
	
	intNumColspan = 4
	
End If

'Richiama le Dichiarazioni per la 
'paginazione avanzata
Call DimPaginazioneAvanzata()

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2007 Carletti Simone, All Rights Reserved" />
<link href="stile.css" rel="stylesheet" type="text/css" />
<!--#include file="includes/inc_meta.asp" -->

<!-- 	ASP Stats Generator <%= strAsgVersion %> è una applicazione gratuita 
		per il monitoraggio degli accessi e dei visitatori ai siti web 
		creata e sviluppata da Simone Carletti.
		
		Puoi scaricarne una copia gratuita sul sito ufficiale http://www.weppos.com/ -->

</head>
<!--#include file="includes/header.asp" -->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= Ucase(strAsgTxtSearchQuery) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td width="5%"  background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"></td>
			<% If gruppo = "motori" Then %>
			<td width="30%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= Ucase(strAsgTxtQuery) %>
				<a href="query.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=Query&order=DESC" %>" title="<%= strAsgTxtOrderBy & strAsgTxtQuery & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="query.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=Query&order=ASC" %>" title="<%= strAsgTxtOrderBy & strAsgTxtQuery & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="20%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= Ucase(strAsgTxtEngine) %>
				<a href="query.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=Engine&order=DESC" %>" title="<%= strAsgTxtOrderBy & strAsgTxtEngine & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="query.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=Engine&order=ASC" %>" title="<%= strAsgTxtOrderBy & strAsgTxtEngine & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<% ElseIf gruppo = "query" Then %>
			<td width="50%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= Ucase(strAsgTxtQuery) %>
				<a href="query.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=Query&order=DESC" %>" title="<%= strAsgTxtOrderBy & strAsgTxtQuery & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="query.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=Query&order=ASC" %>" title="<%= strAsgTxtOrderBy & strAsgTxtQuery & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<% End If %>
			<td width="12%"  background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmVisits) %></td>
			<td width="33%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>">
				<a href="query.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=Hits&order=DESC" %>" title="<%= strAsgTxtOrderBy & strAsgTxtHits & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="query.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=Hits&order=ASC" %>" title="<%= strAsgTxtOrderBy & strAsgTxtHits & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a>
			&nbsp;&nbsp;<%= UCase(strAsgTxtGraph) %>&nbsp;&nbsp;
				<a href="query.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=Visits&order=DESC" %>" title="<%= strAsgTxtOrderBy & strAsgTxtVisits & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="query.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=Visits&order=ASC" %>" title="<%= strAsgTxtOrderBy & strAsgTxtVisits & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
		  </tr>
		<%

		If elenca = "mese" AND gruppo = "motori" Then 
			'Initialise SQL string to select data
			strAsgSQL = "SELECT Query, Engine, SERP_Page, Hits, Visits FROM "&strAsgTablePrefix&"Query WHERE Mese = '" & mese & "' "
			'Call the function to search into the database if there are enought information to do that
			strAsgSQL = CheckSearchForData(strAsgSQL, false)
		ElseIf elenca = "tutti" AND gruppo = "motori" Then 
			'Initialise SQL string to select data
			strAsgSQL = "SELECT Query, Engine, SERP_Page, Hits, Visits FROM "&strAsgTablePrefix&"Query "
			'Call the function to search into the database if there are enought information to do that
			strAsgSQL = CheckSearchForData(strAsgSQL, true)
		ElseIf elenca = "mese" AND gruppo = "query" Then 
			'Initialise SQL string to select data
			strAsgSQL = "SELECT Query, AVG(SERP_Page) AS AvgSERP, SUM(Hits) AS SumHits, SUM(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Query WHERE Mese = '" & mese & "' "
			'Call the function to search into the database if there are enought information to do that
			strAsgSQL = CheckSearchForData(strAsgSQL, false)
			'Group information by following fields
			strAsgSQL = strAsgSQL & " GROUP BY Query "
		ElseIf elenca = "tutti" AND gruppo = "query" Then 
			'Initialise SQL string to select data
			strAsgSQL = "SELECT Query, AVG(SERP_Page) AS AvgSERP, SUM(Hits) AS SumHits, SUM(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Query "
			'Call the function to search into the database if there are enought information to do that
			strAsgSQL = CheckSearchForData(strAsgSQL, true)
			'Group information by following fields
			strAsgSQL = strAsgSQL & " GROUP BY Query "
		End If
		
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
					Call BuildTableContNoRecord(intNumColspan, "search")
					
				'Else show general no record information
				Else
	
					'// Row - No current record			
					Call BuildTableContNoRecord(intNumColspan, "standard")
					
				End If
				
			Else
			
				objAsgRs.PageSize = RecordsPerPage
				objAsgRs.AbsolutePage = page
				intAsgCount = (RecordsPerPage * (page-1))
				
				For PaginazioneLoop = 1 To RecordsPerPage
					
					If Not objAsgRs.EOF Then			
					intAsgCount = intAsgCount + 1
					
		%>		  
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><%= intAsgCount %></td>
			<% If gruppo = "motori" Then %>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left"><% If objAsgRs("SERP_Page") > 0 Then Response.Write("&nbsp;<a href=""serp.asp?serp=" & objAsgRs("SERP_Page") & "&mese=" & mese &"&elenca=" & elenca & """ title=""" & strAsgTxtQuery & "&nbsp;" & objAsgRs("SERP_Page") & "&deg;&nbsp;" & strAsgTxtPage & """><span class=""notetext"">[" & objAsgRs("SERP_Page") & "]</span></a>") %>&nbsp;<%= ShareWords(HighlightSearchKey(objAsgRs("Query"), "Query"), 40) %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<img src="images/engine.asp?icon=<%= objAsgRs("Engine") %>" alt="<%= objAsgRs("Engine") %>" class="def-icon def-icon-searchengine" /> <%= HighlightSearchKey(objAsgRs("Engine"), "Engine") %></td>
			<% ElseIf gruppo = "query" Then %>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left"><% If objAsgRs("AvgSERP") > 0 Then Response.Write("&nbsp;<span class=""notetext"">[" & objAsgRs("AvgSERP") & "]</span>") %>&nbsp;<%= ShareWords(HighlightSearchKey(objAsgRs("Query"), "Query"), 40) %></td>
			<% End If %>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right">&nbsp;<%
				If gruppo = "motori" Then
					Response.Write objAsgRs("Hits") & "<br />" & objAsgRs("Visits")
				ElseIf gruppo = "query" Then
					Response.Write objAsgRs("SumHits") & "<br />" & objAsgRs("SumVisits")
				End If %>
			</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"><% If gruppo = "motori" Then %>
				<img src="<%= strAsgSknPathImage%>bar_graph_image_hits_h.gif" width="<%= FormatNumber(objAsgRs("Hits")*intAsgParte, 2) %>" height="9" alt="<%= strAsgTxtHits %>" align="absmiddle" /> [<%= CalcolaPercentuale(intAsgTotMeseHits, objAsgRs("Hits")) %>]<br />
				<img src="<%= strAsgSknPathImage%>bar_graph_image_visits_h.gif" width="<%= FormatNumber(objAsgRs("Visits")*intAsgParte, 2) %>" height="9" alt="<%= strAsgTxtVisits %>" align="absmiddle" /> [<%= CalcolaPercentuale(intAsgTotMeseVisits, objAsgRs("Visits")) %>]
				<% ElseIf gruppo = "query" Then %>
				<img src="<%= strAsgSknPathImage%>bar_graph_image_hits_h.gif" width="<%= FormatNumber(objAsgRs("SumHits")*intAsgParte, 2) %>" height="9" alt="<%= strAsgTxtHits %>" align="absmiddle" /> [<%= CalcolaPercentuale(intAsgTotMeseHits, objAsgRs("SumHits")) %>]<br />
				<img src="<%= strAsgSknPathImage%>bar_graph_image_visits_h.gif" width="<%= FormatNumber(objAsgRs("SumVisits")*intAsgParte, 2) %>" height="9" alt="<%= strAsgTxtVisits %>" align="absmiddle" /> [<%= CalcolaPercentuale(intAsgTotMeseVisits, objAsgRs("SumVisits")) %>]
				<% End If %>
			</td>
		  </tr>
		<%

				objAsgRs.MoveNext
				End If
			Next
			End If
		
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(intNumColspan)

		'// Row - Advanced data sorting
		Response.Write(vbCrLf & "<tr class=""smalltext""><td colspan=""" & intNumColspan & """ align=""center""><br />")
		Call PaginazioneAvanzata("query.asp", "")
		Response.Write(vbCrLf & "<br /><br /></td></tr>")
		
		objAsgRs.Close

		'Reset Server Objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing

		'// Row - Data output panels
		Response.Write(vbCrLf & "<tr class=""smalltext"" align=""center"" valign=""top""><td colspan=""" & intNumColspan & """ height=""25""><br />")
		Call GoToPeriod("query.asp", "")
		Call GoToGrouping("query.asp", "")

		If gruppo = "motori" Then
		Call SearchForData("query.asp", "", "Query|Engine")
		ElseIf gruppo = "query" Then
		Call SearchForData("query.asp", "", "Query")
		End If

		%>
			<!-- grouping panel -->
			<table width="300" border="0" cellspacing="0" cellpadding="0" height="30">
			  <tr class="smalltext" valign="middle" align="center">
				<td width="100%">
				  <% If gruppo = "query" Then %>
					<input type="button" onClick="location.href='query.asp?<%= "gruppo=motori&mese="  & mese & "&elenca=" & elenca & "&page=" & page & "&sort=" & strAsgSortBy & "&order" & strAsgSortOrder %>'" name="<%= strAsgTxtGroupByEngine %>" value="<%= strAsgTxtGroupByEngine %>" />
				  <% ElseIf gruppo = "motori" Then %>
					<input type="button" onClick="location.href='query.asp?<%= "gruppo=query&mese=" & mese & "&elenca=" & elenca & "&page=" & page & "&sort=" & strAsgSortBy & "&order" & strAsgSortOrder %>'" name="<%= strAsgTxtGroupByQuery %>" value="<%= strAsgTxtGroupByQuery %>" />
				  <% End If %>
				</td>
			  </tr>
			</table>
			<!-- / grouping panel -->
		<%
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