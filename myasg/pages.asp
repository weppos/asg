<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
Call AllowEntry("True", "False", "False", intAsgSecurity)


'Dichiara Variabili
Dim mese			'Riferimento per output
Dim elenca			'Tutti | Mese
Dim gruppo			'Raggruppa per Path | Non raggruppare per Path
Dim dettagli		'Path di cui mostrare le informazioni
Dim asgOutputPage	'Pagina di output
Dim intAsgCount		'Conteggio record


intAsgCount	= 0


'Read setting variables from querystring
mese = Request.QueryString("mese")
elenca = Request.QueryString("elenca")
gruppo = Request.QueryString("gruppo")
dettagli = Request.QueryString("dettagli")
strAsgSortBy = Request.QueryString("sort")
strAsgSortOrder = Request.QueryString("order")


'If period variable is empty then set it to the current month
If mese = "" Then mese = dtmAsgMonth & "-" & dtmAsgYear
'If the variable is empty set it to monthly report
If elenca = "" Then elenca = "mese"
'If the variable is empty set no record grouping
If gruppo = "" Then gruppo = "nessuno"
' Set the sorting order depending on querystring
if strAsgSortOrder = "ASC" then 
	strAsgSortOrder = "ASC"
else
	strAsgSortOrder = "DESC"
end if
'If a time period has been chosen then build the variable to query the database
If Request.QueryString("showperiod") = strAsgTxtShow Then mese = Request.QueryString("periodmm") & "-" & Request.QueryString("periodyy")


'Read sorting order from querystring
'// Filter QS values and associate them 
'// with their respective database fields
Select Case strAsgSortBy
	Case "hits" strAsgSortByFld = "SUM(Hits)"
	Case "visits" strAsgSortByFld = "SUM(Visits)"
	Case "page" 
		If gruppo = "path" Then
			strAsgSortByFld = "Max(Page)"
		ElseIF gruppo = "nessuno" Then
			strAsgSortByFld = "Page"
		End If 
	Case "path" strAsgSortByFld = "Page_Stripped"
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
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtHits) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td width="5%"  background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"><%= UCase(strAsgTxtNum) %></td>
			<% If gruppo = "nessuno" Then %>
			<td width="66%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtHits) %>
				<a href="pages.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=page&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtURL & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="pages.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=page&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtURL & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<% ElseIf gruppo = "path" Then %>
			<td width="66%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtPath) %>
				<a href="pages.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=path&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtURL & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="pages.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=path&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtURL & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<% End If %>			
			<td width="12%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmHits) %>
				<a href="pages.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=hits&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtHits & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="pages.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=hits&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtHits & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="12%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmVisits) %>
				<a href="pages.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=visits&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtVisits & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="pages.asp?<%= "mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&sort=visits&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtVisits & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="5%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"></td>
		  </tr>
		<%

		If elenca = "mese" AND gruppo = "nessuno" Then
			'Initialise SQL string to select data
			strAsgSQL = "SELECT Page, SUM(Hits) AS SumHits, SUM(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Page WHERE Mese = '" & mese & "' "
			'Call the function to search into the database if there are enought information to do that
			strAsgSQL = CheckSearchForData(strAsgSQL, false)
		ElseIf elenca = "tutti" AND gruppo = "nessuno" Then
			'Initialise SQL string to select data
			strAsgSQL = "SELECT Page, SUM(Hits) AS SumHits, SUM(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Page "
			'Call the function to search into the database if there are enought information to do that
			strAsgSQL = CheckSearchForData(strAsgSQL, true)
		ElseIf elenca = "mese" AND gruppo = "path" Then
			'Initialise SQL string to select data
			strAsgSQL = "SELECT Page_Stripped, SUM(Hits) AS SumHits, SUM(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Page WHERE Mese = '" & mese & "' "
			'Call the function to search into the database if there are enought information to do that
			strAsgSQL = CheckSearchForData(strAsgSQL, false)
		ElseIf elenca = "tutti" AND gruppo = "path" Then
			'Initialise SQL string to select data
			strAsgSQL = "SELECT Page_Stripped, SUM(Hits) AS SumHits, SUM(Visits) AS SumVisits FROM "&strAsgTablePrefix&"Page "
			'Call the function to search into the database if there are enought information to do that
			strAsgSQL = CheckSearchForData(strAsgSQL, true)
		End If

		If gruppo = "nessuno" Then
			'Group information by following fields
			strAsgSQL = strAsgSQL & " GROUP BY Page "
		ElseIf gruppo = "path" Then
			'Group information by following fields
			strAsgSQL = strAsgSQL & " GROUP BY Page_Stripped "
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
					Call BuildTableContNoRecord(5, "search")
					
				'Else show general no record information
				Else
	
					'// Row - No current record			
					Call BuildTableContNoRecord(5, "standard")
					
				End If
				
			Else
			
				objAsgRs.PageSize = RecordsPerPage
				objAsgRs.AbsolutePage = page
				intAsgCount = (RecordsPerPage * (page-1))
				
				For PaginazioneLoop = 1 To RecordsPerPage
					
					If Not objAsgRs.EOF Then			
					intAsgCount = intAsgCount + 1
				
		Response.Write vbCrLf & "<tr class=""smalltext"" bgcolor=""" & strAsgSknTableContBgColour & """>"

			'Numero
			Response.Write vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"">" & intAsgCount & "</td>"
			
			'-----------------------------------------------------------|
			
			'No raggruppamenti - Mostra PAGINA
			If gruppo = "nessuno" Then
				

				'PAGINA
				Response.Write vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""left""><a class=""linksmalltext"" href=""" & objAsgRs("Page") & """ target=""_blank"" title=""" & objAsgRs("Page") & """>"

					'Verifica la pagina e mostra o meno
					'una icona standard di corrispondenza dominio.
					Response.Write(ChooseDomainIcon(objAsgRs("Page"), "classic"))
						
					'TAGLIA STRINGHE LUNGHE
					'Se la stringa supera i 55 caratteri inserisci ... in mezzo e accorcia
					'Caso "Nessun Raggruppamento" - Max 55 Caratteri
					Response.Write(HighlightSearchKey(StripValueTooLong(asgOutputPage, 55, 25, 25), "Page"))
	
				Response.Write vbCrLf & "</a></td>"
					
			'-----------------------------------------------------------|
					
			'Raggruppamento Path - Mostra cella PATH
			ElseIf gruppo = "path" Then
			
				asgOutputPage = objAsgRs("Page_Stripped")
				'Taglia tutto il prefisso sito + http:// se non è una pagina sconosciuta
				If Mid(asgOutputPage, 1, Len(strAsgSiteURLremote)) = strAsgSiteURLremote Then asgOutputPage = Mid(asgOutputPage, Len(strAsgSiteURLremote) + 1) 

				'PATH
				Response.Write(vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""left"">&nbsp;")
					
					'Write an anchor
					Response.Write(vbCrLf & "<a name=""" & objAsgRs("Page_Stripped") & """></a>")
						
					'Se è un raggruppamento per PATH mostra
					'l'icona per l'espansione dei dettagli		
					If gruppo = "path" Then
		
						Response.Write(vbCrLf & "<a href=""pages.asp?dettagli=" & objAsgRs("Page_Stripped") & "&mese=" & mese & "&elenca=" & elenca & "&gruppo=" & gruppo & "&searchfor=" & asgSearchfor & "&searchin=" & asgSearchin & "&sort=" & strAsgSortBy & "&order=" & strAsgSortOrder & "&page=" & page & "#" & objAsgRs("Page_Stripped") & """ title=""" & strAsgTxtHits & """>")
		
						'Icona espansa se Corrisponde
						If Trim(dettagli) <> "" AND objAsgRs("Page_Stripped") = Trim(dettagli) Then
							Response.Write("<img src=""" & strAsgSknPathImage & "expanded.gif"" alt=""" & strAsgTxtHits & """ border=""0"" align=""absmiddle"" />")
						'Icona espandi se Differente
						Else
							Response.Write("<img src=""" & strAsgSknPathImage & "expand.gif"" alt=""" & strAsgTxtHits & """ border=""0"" align=""absmiddle"" />")
						End If
		
						Response.Write("</a>")
		
					End If

					'Verifica la pagina e mostra o meno
					'una icona standard di corrispondenza dominio.
					Response.Write(ChooseDomainIcon(objAsgRs("Page_Stripped"), "classic"))
						
					'TAGLIA STRINGHE LUNGHE
					'Se la stringa supera i 25 caratteri inserisci ... in mezzo e accorcia
					'Max 55 Caratteri
					Response.Write(HighlightSearchKey(StripValueTooLong(asgOutputPage, 75, 35, 35), "Page_Stripped"))
	
				Response.Write vbCrLf & "</td>"
					
			'-----------------------------------------------------------|
					
			End If
			
			'-----------------------------------------------------------|
			
			'Visite
			Response.Write(vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""right"">")
		  	Response.Write(objAsgRs("SumHits")) 
		 	Response.Write(vbCrLf & "</td>")
			
			'-----------------------------------------------------------|
			
			'Accessi
			Response.Write(vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""right"">")
			Response.Write(objAsgRs("SumVisits") )
		 	Response.Write(vbCrLf & "</td>")
			
			'-----------------------------------------------------------|
			
			'Ultima cella
			Response.Write(vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"">")
			Response.Write("</td>")
			
			'-----------------------------------------------------------|

		Response.Write vbCrLf & "</tr>"
			
		If gruppo = "path" Then

			If Trim(dettagli) <> "" AND objAsgRs("Page_Stripped") = Trim(dettagli) Then
				
				Dim objAsgRs2
				
				'Mostra le query al motore
				Set objAsgRs2 = Server.CreateObject("ADODB.Recordset")
				
				If elenca = "mese" Then 
					strAsgSQL = "SELECT Page, Hits, Visits FROM "&strAsgTablePrefix&"Page WHERE Page_Stripped = '" & dettagli & "' AND Mese = '" & mese & "' "
				ElseIf elenca = "tutti" Then 
					strAsgSQL = "SELECT Page, Hits, Visits FROM "&strAsgTablePrefix&"Page WHERE Page_Stripped = '" & dettagli & "' "
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
					<!-- Dettagli Pagine -->
					<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
					  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
						<td width="80%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"><%= UCase(strAsgTxtHits) %></td>
						<td width="10%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmHits) %></td>
						<td width="10%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmVisits) %></td>
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
						<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left"><a href="<%= objAsgRs2("Page") %>" title="<%= objAsgRs2("Page") %>" target="_blank" class="linksmalltext">
						<%= StripValueTooLong(objAsgRs2("Page"), 75, 35, 35) %></a>
						</td>
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
						Response.Write(vbCrLf & "<tr class=""smalltext""><td colspan=""4"" align=""center""><br />")
						Call PaginazioneAvanzataDettagli("pages.asp", "")
						Response.Write(vbCrLf & "<br /><br /></td></tr>")
									  
						objAsgRs2.Close
						Set objAsgRs2 = Nothing
					  
					  %>
					</table><br />
					<!-- Fine Dettagli Pagine -->
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

		'Fine condizione Caso Path
		End If
		
				objAsgRs.MoveNext
				End If
			Next
			End If
				
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(5)

		'// Row - Advanced data sorting
		Response.Write(vbCrLf & "<tr class=""smalltext""><td colspan=""5"" align=""center""><br />")
		Call PaginazioneAvanzata("pages.asp", "")
		Response.Write(vbCrLf & "<br /><br /></td></tr>")

		objAsgRs.Close

		'Reset Server Objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing

		'// Row - Data output panels
		Response.Write(vbCrLf & "<tr class=""smalltext"" align=""center"" valign=""top""><td colspan=""5"" height=""25""><br />")
		Call GoToPeriod("pages.asp", "")
		Call GoToGrouping("pages.asp", "")
		
		If gruppo = "nessuno" Then
		Call SearchForData("pages.asp", "", "Page")
		ElseIf gruppo = "path" Then
		Call SearchForData("pages.asp", "", "Page_Stripped")
		End If
		
		%>
			<!-- grouping panel -->
			<table width="300" border="0" cellspacing="0" cellpadding="0" height="30">
			  <tr class="smalltext" valign="middle" align="center">
				<td width="100%">
				  <% If gruppo = "nessuno" Then %>
					<input type="button" onClick="location.href='pages.asp?<%= "gruppo=path&mese="    & mese & "&elenca=" & elenca & "&sort=" & strAsgSortBy & "&order=" & strAsgSortOrder %>'" name="<%= strAsgTxtGroupByPath %>" value="<%= strAsgTxtGroupByPath %>" />
				  <% ElseIf gruppo = "path" Then %>
					<input type="button" onClick="location.href='pages.asp?<%= "gruppo=nessuno&mese=" & mese & "&elenca=" & elenca & "&sort=" & strAsgSortBy & "&order=" & strAsgSortOrder %>'" name="<%= strAsgTxtGroupByPage %>" value="<%= strAsgTxtGroupByPage %>" />
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