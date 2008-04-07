<%@LANGUAGE="VBSCRIPT"%>
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
Dim gruppo			'Nessuno | Dominio
Dim tipo			'Esterni | Interni | Tutti
Dim dettagli		'Referer di cui mostrare le informazioni
Dim intAsgCount		'Conteggio record


'Read setting variables from querystring
mese = Request.QueryString("mese")
gruppo = Request.QueryString("gruppo")
tipo = Request.QueryString("tipo")
dettagli = Request.QueryString("dettagli")
strAsgSortBy = Request.QueryString("sort")
strAsgSortOrder = Request.QueryString("order")
intAsgCount	= 0


'If period variable is empty then set it to the current month
If mese = "" Then mese = dtmAsgMonth & "-" & dtmAsgYear
'If the variable is empty set it to external referers
If tipo = "" Then tipo = "esterni"
'If the variable is empty set no grouping mode
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
	Case "visits" 
		strAsgSortByFld = "SUM(Visits)"
	Case "page" 
		If gruppo = "nessuno" Then 
			strAsgSortByFld = "Referer"
		ElseIf gruppo = "dominio" Then
			strAsgSortByFld = "Dominio"
		End If
	Case "data" 
		strAsgSortByFld = "MAX(Last_Access)"
	Case Else 
		strAsgSortByFld = "SUM(Visits)"
End Select

'Richiama le Dichiarazioni per la 
'paginazione avanzata
Call DimPaginazioneAvanzata()

'Richiama le Dichiarazioni per la 
'paginazione avanzata [dei dettagli]
Call DimPaginazioneAvanzataDettagli()

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | powered by ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2008 Carletti Simone, All Rights Reserved" />
<meta name="generator" content="ASP Stats Generator <%= strAsgVersion %>" /> <!-- leave this for stats -->

<!--#include file="includes/html-head.asp" -->

<!--
  ASP Stats Generator (release <%= strAsgVersion %>) is a free software package
  completely written in ASP programming language, for real time visitor tracking.
  Get your own copy for free at http://www.asp-stats-com/ !
-->

</head>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle">
			<% 
			Select Case tipo
				Case "esterni" Response.Write UCase(strAsgTxtRefererOut)
				Case "interni" Response.Write UCase(strAsgTxtRefererIn)
				Case "tutti" Response.Write UCase(strAsgTxtRefererAll)
			End Select
			%>
			</td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td width="5%"  background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"></td>
			<td width="55%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><% 
			If gruppo = "nessuno" Then 
				Response.Write(UCase(strAsgTxtReferer)) %>
				<a href="referer.asp?<%= "mese=" & mese & "&tipo=" & tipo & "&gruppo=" & gruppo & "&sort=page&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtReferer & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="referer.asp?<%= "mese=" & mese & "&tipo=" & tipo & "&gruppo=" & gruppo & "&sort=page&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtReferer & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a>
			<%	
			ElseIf gruppo = "dominio" Then
				Response.Write(UCase(strAsgTxtDomain)) %>
				<a href="referer.asp?<%= "mese=" & mese & "&tipo=" & tipo & "&gruppo=" & gruppo & "&sort=page&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtDomain & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="referer.asp?<%= "mese=" & mese & "&tipo=" & tipo & "&gruppo=" & gruppo & "&sort=page&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtDomain & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a>
			    <%	
			End If
			%></td>
			<td width="23%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtLastAccess) %>
				<a href="referer.asp?<%= "mese=" & mese & "&tipo=" & tipo & "&gruppo=" & gruppo & "&sort=data&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtLastAccess & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="referer.asp?<%= "mese=" & mese & "&tipo=" & tipo & "&gruppo=" & gruppo & "&sort=data&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtLastAccess & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="12%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmVisits) %>
				<a href="referer.asp?<%= "mese=" & mese & "&tipo=" & tipo & "&gruppo=" & gruppo & "&sort=visits&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtVisits & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="referer.asp?<%= "mese=" & mese & "&tipo=" & tipo & "&gruppo=" & gruppo & "&sort=visits&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtVisits & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="5%"  background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"></td>
		  </tr>
		<%
		
		'Adatta temporaneamente la Path
		strAsgSiteURLremote = Right(strAsgSiteURLremote, Len(strAsgSiteURLremote)-7)

		'Componi la query di richiamo dei record
		'// Query base...
		strAsgSQL = "SELECT SUM(Visits) AS SumVisits, "

		'// Stabilisci condizione di Raggruppamento - I Parte
		If gruppo = "nessuno" Then 
			'Richiama i Referer del Mese raggruppando per Referer 
			'( Mantieni compatibilità sintassi Max() )
			strAsgSQL = strAsgSQL & " Referer, Max(Last_Access) AS MaxData FROM "&strAsgTablePrefix&"Referer WHERE Mese = '" & mese & "' "
		
		ElseIf gruppo = "dominio" Then
			'Richiama i Referer del Mese raggruppando per Dominio
			strAsgSQL = strAsgSQL & " Dominio, Max(Last_Access) AS MaxData FROM "&strAsgTablePrefix&"Referer WHERE Mese = '" & mese & "' "
		
		End If
		
		'Call the function to search into the database if there are enought information to do that
		strAsgSQL = CheckSearchForData(strAsgSQL, false)
		
		'// Stabilisci condizione di Filtro tipologia
		If tipo = "interni" AND gruppo <> "dominio"  Then
			'Richiama filtrando solo i domini interni usando il dominio del sito
			'inserito nella configurazione.
			'Se il raggruppamento è per dominio è inutile mostrare la condizione
			'perchè tanto sarebbe un solo record e fatica sprecata!
			strAsgSQL = strAsgSQL & " AND Dominio = '" & strAsgSiteURLremote & "' "
	
		ElseIf tipo = "tutti" Then
			'Non filtri una mazza! Predi finchè ne hai... ;oP
			
		Else
			'
			strAsgSQL = strAsgSQL & " AND Dominio <> '" & strAsgSiteURLremote & "'"
			
		End If

		'// Stabilisci condizione di Raggruppamento - II Parte
		If gruppo = "nessuno" Then 
			'Richiama i Referer del Mese raggruppando per Referer 
			'( Mantieni compatibilità sintassi Max() )
			strAsgSQL = strAsgSQL & " GROUP BY Referer "
		
		ElseIf gruppo = "dominio" Then
			'Richiama i Referer del Mese raggruppando per Dominio
			strAsgSQL = strAsgSQL & " GROUP BY Dominio "
		
		End If

		'// Stabilisci condizione di Ordinamento
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

		%>		  
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center" height="16"><%= intAsgCount %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">
				<%
					
				If gruppo = "nessuno" Then 
					'Genera il link alla pagina
					Response.Write("<a class=""linksmalltext"" href=""" & objAsgRs("Referer") & """ target=""_blank"">")				
					'TAGLIA STRINGHE LUNGHE
					Response.Write(HighlightSearchKey(StripValueTooLong(objAsgRs("Referer"), 65, 30, 30), "Referer"))
					Response.Write("</a>")				

				ElseIf gruppo = "dominio" Then
										
					'Write an anchor
					Response.Write(vbCrLf & "<a name=""" & objAsgRs("Dominio") & """></a>")
		
					'Espandi Dettagli
					'// Link
					Response.Write(vbCrLf & "				<a href=""referer.asp?dettagli=" & objAsgRs("Dominio") & "&mese=&page=" & page & "&gruppo=" & gruppo & "&searchfor=" & asgSearchfor & "&searchin=" & asgSearchin & "&sort=" & strAsgSortBy & "&order=" & strAsgSortOrder & "#" & objAsgRs("Dominio") & """ title=""" & objAsgRs("Dominio") & """>")
					'// Icona espansa se Corrisponde
					If Trim(dettagli) <> "" AND objAsgRs("Dominio") = Trim(dettagli) Then
						Response.Write(vbCrLf & "				<img src=""" & strAsgSknPathImage & "expanded.gif"" alt=""" & objAsgRs("Dominio") & """ border=""0"" align=""absmiddle"" />")
					'// Icona espandi se Differente
					Else
						Response.Write(vbCrLf & "				<img src=""" & strAsgSknPathImage & "expand.gif"" alt=""" & objAsgRs("Dominio") & """ border=""0"" align=""absmiddle"" />")
					End If
					'// Chiudi Link
					Response.Write("</a>&nbsp;&nbsp;")
		
					'Genera il link al dominio
					Response.Write("<a class=""linksmalltext"" href=""http://" & objAsgRs("Dominio") & """ target=""_blank"">")				
					'TAGLIA STRINGHE LUNGHE
					Response.Write(HighlightSearchKey(StripValueTooLong(objAsgRs("Dominio"), 65, 30, 30), "Dominio"))
					Response.Write("</a>")				
				
				End If
				
				%>
			</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><%= FormatOutTimeZone(objAsgRs("MaxData"), "Date") & "&nbsp;" & strAsgTxtAt & "&nbsp;" & FormatOutTimeZone(objAsgRs("MaxData"), "Time") %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right" ><%= objAsgRs("SumVisits") %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>"				 ></td>
		  </tr>
		<% 
			'Solo nel caso ci sia un raggruppamento per dominio 
			'verifica la condizione dei dettagli
			If gruppo = "dominio" Then
			
			If Trim(dettagli) <> "" AND objAsgRs("Dominio") = Trim(dettagli) Then
				
				Dim objAsgRs2
				
				'Mostra le query al motore
				Set objAsgRs2 = Server.CreateObject("ADODB.Recordset")
				
				'Disabilitato al momento la classificazione
				'divisa per mese.
				'L'applicazione richiama solamente i referer del mese in corso.
				'If elenca = "mese" Then 
					strAsgSQL = "SELECT Referer, Visits, Last_Access FROM "&strAsgTablePrefix&"Referer WHERE Dominio = '" & dettagli & "' AND Mese = '" & mese & "' "
				'ElseIf elenca = "tutti" Then 
				'	strAsgSQL = "SELECT Referer, Visits, Data FROM "&strAsgTablePrefix&"Referer WHERE Dominio = '" & dettagli & "' "
				'End If
				
				strAsgSQL = strAsgSQL & " ORDER BY Visits DESC "
		
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
						<td width="68%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"><%= UCase(strAsgTxtReferer) %></td>
						<td width="20%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtLastAccess) %></td>
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
						<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left" height="16">&nbsp;
							<a href="<%= objAsgRs2("Referer") %>" title="<%= objAsgRs2("Referer") %>" target="_blank" class="linksmalltext">
							<%= StripValueTooLong(objAsgRs2("Referer"), 70, 32, 32) %></a></td>
						<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><%= FormatOutTimeZone(objAsgRs2("Last_Access"), "Date") %></td>
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
						Call PaginazioneAvanzataDettagli("referer.asp", "")
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
			'Fine condizione
			'corrispondenza dei dettagli
			End If

			'Fine condizione
			'verifica dei dettagli
			End If
			
				
				objAsgRs.MoveNext
				End If
			Next
			End If
				
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(5)

		'// Row - Advanced data sorting
		Response.Write(vbCrLf & "<tr class=""smalltext""><td colspan=""5"" align=""center""><br />")
		Call PaginazioneAvanzata("referer.asp", "")
		Response.Write(vbCrLf & "<br /><br /></td></tr>")

		objAsgRs.Close
		
		'Reset Server Objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing

		'// Row - Data output panels
		Response.Write(vbCrLf & "<tr class=""smalltext"" align=""center"" valign=""top""><td colspan=""5"" height=""25""><br />")
		Call GoToPeriod("referer.asp", "") 
		
		If blnRefererServer AND gruppo <> "dominio" Then %>
			<!-- Visualizza in base a tipo -->
			<table width="300" border="0" cellspacing="0" cellpadding="0" height="30">
			<form action="referer.asp?<%= "mese=" & mese & "&gruppo=" & gruppo & "&sort=" & strAsgSortBy & "&order=" & strAsgSortOrder & "" %>" method="post">
			  <tr class="smalltext" valign="middle" align="left">
				<td width="25%"><%= strAsgTxtTypology %></td>
				<td width="65%">
				<select name="tipo" class="smallform">
					<option value="esterni" <% If tipo = "esterni" Then Response.Write "selected" End If %>><%= strAsgTxtRefererOut %></option>
					<option value="interni" <% If tipo = "interni" Then Response.Write "selected" End If %>><%= strAsgTxtRefererIn %></option>
					<option value="tutti" <% If tipo = "tutti" Then Response.Write "selected" End If %>><%= strAsgTxtRefererAll %></option>
				</select>
				</td>
				<td width="10%"><input type="Submit" name="Mostra_tipo" value="Mostra" /></td>
			  </tr>
			</form>
			
			</table>
			<!-- Fine Visualizza in base a tipo -->
		<% 
		
		End If 
		
		If gruppo = "nessuno" Then
		Call SearchForData("referer.asp", "", "Referer")
		ElseIf gruppo = "dominio" Then
		Call SearchForData("referer.asp", "", "Dominio")
		End If
		
		%>
			<!-- grouping panel -->
			<table width="300" border="0" cellspacing="0" cellpadding="0" height="30">
			  <tr class="smalltext" valign="middle" align="center">
				<td width="100%">
				  <% If gruppo = "nessuno" Then %>
					<input type="button" onClick="location.href='referer.asp?<%= "gruppo=dominio&mese=" & mese & "&tipo=" & tipo & "&sort=" & strAsgSortBy & "&order=" & strAsgSortOrder %>'" name="<%= strAsgTxtGroupByDomain %>" value="<%= strAsgTxtGroupByDomain %>" />
				  <% ElseIf gruppo = "dominio" Then %>
					<input type="button" onClick="location.href='referer.asp?<%= "gruppo=nessuno&mese=" & mese & "&tipo=" & tipo & "&sort=" & strAsgSortBy & "&order=" & strAsgSortOrder %>'" name="<%= strAsgTxtGroupByReferer %>" value="<%= strAsgTxtGroupByReferer %>" />
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