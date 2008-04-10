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
Dim elenca			'Tutti | Mese
Dim intAsgCount		'Conteggio record
Dim strSerpValue	'Valori di Serp
Dim arySerpValue	'Valori di Serp splittati in array
Dim loopSerp		'Loop di output dei dati ricavati in serp
Dim loopData		'Loop di output dei dati ricavati in serp
Dim blnNoData
Dim blnSerpValue	'Imposta a true se selezionata una serp


'Read setting variables from querystring
mese = Request.QueryString("mese")
elenca = Request.QueryString("elenca")
strSerpValue = Request.QueryString("serp")
strAsgSortByFld = "Visits"
strAsgSortOrder = "DESC"


'If period variable is empty then set it to the current month
If mese = "" Then mese = dtmAsgMonth & "-" & dtmAsgYear
'If the variable is empty set it to monthly report
If elenca = "" Then elenca = "mese"
'If a time period has been chosen then build the variable to query the database
If Request.QueryString("showperiod") = strAsgTxtShow Then mese = Request.QueryString("periodmm") & "-" & Request.QueryString("periodyy")


'Read SERP value from Querystring
If IsNumeric(strSerpValue) AND Len(strSerpValue) > 0 Then
	strSerpValue = CInt(strSerpValue)
Else
	strSerpValue = ""
End If

'Set to false elaboration variables
blnNoData = False
blnSerpValue = False

	
'Richiama le Dichiarazioni per la 
'paginazione avanzata
Call DimPaginazioneAvanzata()
	

'Procedi con la SERP scelta
If strSerpValue <> "" Then
	
	strSerpValue = CInt(strSerpValue)
	arySerpValue = Split(strSerpValue, "|")
	blnSerpValue = True

Else

	'Verifica i valori delle pagine SERP
	'presenti nel database
	If elenca = "mese" Then 
		strAsgSQL = "SELECT DISTINCT Serp_Page FROM "&strAsgTablePrefix&"Query WHERE Mese = '" & mese & "' "
	ElseIf elenca = "tutti" Then 
		strAsgSQL = "SELECT DISTINCT Serp_Page FROM "&strAsgTablePrefix&"Query "
	End If
	
	'Richiama le informazioni ed inseriscile in una stringa
	objAsgRs.Open strAsgSQL, objAsgConn
	If Not objAsgRs.EOF Then
		'Cicla i record trovati
		Do While Not objAsgRs.EOF
			'Inserisci il valore in stringa solo se sono presenti dei dati
			If Trim(Len(objAsgRs("SERP_Page"))) > 0 Then strSerpValue = strSerpValue & objAsgRs("SERP_Page") & "|"
			objAsgRs.MoveNext
		Loop
		'Controllo record estratti
		If Not Trim(Len(strSerpValue)) > 0 Then blnNoData = True
	'Verifica la mancanza di dati per procedere
	Else
		blnNoData = True
	End If
	objAsgRs.Close

	'Procedi solo se presenti dati
	If Not blnNoData Then
		'Purifica la stringa per evitare split non corretti
		strSerpValue = Left(strSerpValue, Len(strSerpValue)-1)
		'Splitta i valori ricavando un array
		arySerpValue = Split(strSerpValue, "|")
		'Imposta il nuovo valore massimo di record
		RecordsPerPage = 10
	End If

'/Scelta singola SERP
End If

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
<!--#include file="includes/header.asp" -->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= Ucase(strAsgTxtSERPreports) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
<%

'Procedi solo verificata la presenza di dati
If Not blnNoData Then

	'Esegui un ciclo che mostri il TOP 5 di ogni SERP
	For loopSerp = 0 to Ubound(arySerpValue)
		If Len(Trim(arySerpValue(loopSerp))) > 0 Then

%>		
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td width="5%"  background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"></td>
			<td width="58%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><% If Not blnSerpValue Then Response.Write(Ucase(strAsgTxtTop) & "&nbsp;") End If : Response.Write(Ucase(strAsgTxtQuery) & "&nbsp;" & Ucase(strAsgTxtOn)  & "&nbsp;" & arySerpValue(loopSerp) & "&deg;&nbsp;" & Ucase(strAsgTxtPage)) %></td>
			<td width="25%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= Ucase(strAsgTxtEngine) %></td>
			<td width="12%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmVisits) %></td>
		  </tr>
		<%

		If elenca = "mese" Then 
			'Initialise SQL string to select data
			strAsgSQL = "SELECT Query, Engine, Visits, SERP_Page FROM "&strAsgTablePrefix&"Query WHERE Mese = '" & mese & "' AND SERP_Page = " & arySerpValue(loopSerp) & ""
		ElseIf elenca = "tutti" Then 
			'Initialise SQL string to select data
			strAsgSQL = "SELECT Query, Engine, Visits, SERP_Page FROM "&strAsgTablePrefix&"Query WHERE SERP_Page = " & arySerpValue(loopSerp) & ""
		End If
		
		'Call the function to search into the database if there are enought information to do that
		strAsgSQL = CheckSearchForData(strAsgSQL, false)
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
					Call BuildTableContNoRecord(4, "search")
					
				'Else show general no record information
				Else
	
					'// Row - No current record			
					Call BuildTableContNoRecord(4, "standard")
					
				End If
				
			Else

				'Imposta paginazione record
				If blnSerpValue Then
					objAsgRs.PageSize = RecordsPerPage
					objAsgRs.AbsolutePage = page
					intAsgCount = (RecordsPerPage * (page-1))
				Else
					intAsgCount = 0
				End If
			
				For loopData = 1 to RecordsPerPage
					
					If Not objAsgRs.EOF Then			
					intAsgCount = intAsgCount + 1
					
		%>		  
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><%= intAsgCount %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left"><% If Len(objAsgRs("SERP_Page")) > 0 Then Response.Write("&nbsp;<span class=""notetext"">[" & objAsgRs("SERP_Page") & "]</span>") %>&nbsp;<%= ShareWords(HighlightSearchKey(objAsgRs("Query"), "Query"), 40) %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<img src="images/engine.asp?icon=<%= objAsgRs("Engine") %>" alt="<%= objAsgRs("Engine") %>" align="absmiddle" /> <%= HighlightSearchKey(objAsgRs("Engine"), "Engine") %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right">&nbsp;<%= objAsgRs("Visits") %></td>
		  </tr>
		<%

				objAsgRs.MoveNext
				End If
			Next
			End If
		
		%>
		  <tr bgcolor="<%= strAsgSknTableContBgColour %>" align="center" class="smalltext">
			<td width="100%" background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" height="16" colspan="4" align="right">
			<%
			
			'Costruisci base link...
			Response.Write(vbCrlf & "<a href=""serp.asp?searchfor=" & asgSearchfor & "&searchin=" & asgSearchin & "")
			'...aggiungi querystring se necessario...
			If Not blnSerpValue Then Response.Write("&serp=" & arySerpValue(loopSerp) & "&mese=" & mese & "&elenca=" & elenca & "")
			'...concludi link...
			Response.Write(""" title=""" & strAsgTxtServerInformations & """ class=""linksmalltext"">")
			'...stampa a video tipo dati...
			If blnSerpValue Then 
			Response.Write(strAsgTxtTop & "&nbsp;" & strAsgTxtQuery)
			Else
			Response.Write(strAsgTxtFullVersion)
			End If
			'...concludi link
			Response.Write(vbCrlf & "<img src=""" & strAsgSknPathImage & "arrow_small_dx.gif"" alt=""" & strAsgTxtFullVersion & """ align=""middle"" border=""0"" /></a>")

			%>
			</td>
		  </tr>
		<%
		
			'// Row - End table spacer			
			Call BuildTableContEndSpacer(4)

		'Imposta paginazione record
		If blnSerpValue Then
	
			'// Row - Advanced data sorting
			Response.Write(vbCrLf & "<tr class=""smalltext""><td colspan=""4"" align=""center""><br />")
			Call PaginazioneAvanzata("serp.asp", "")
			Response.Write(vbCrLf & "<br /><br /></td></tr>")
			
		End If

		objAsgRs.Close

		%>
		</table><%
		
		'Per staccare il layout stampa lo spazio
		If Not blnSerpValue Then Response.Write("<br />")

		'/Condizione valore SERP
		End If
	'/Ciclo di SERP
	Next


'Mancanza di dati
Else

		'//	Costruisci layout ed informazioni
		Response.Write(vbCrLf & "<table width=""90%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"">")
		Response.Write(vbCrLf & "<tr bgcolor=""" & strAsgSknTableTitleBgColour & """ align=""center"" class=""normaltitle"">")
		Response.Write(vbCrLf & "<td width=""100%"" background=""" & strAsgSknPathImage & strAsgSknTableTitleBgImage & """ height=""16""></td>")
		Response.Write(vbCrLf & "</tr>")
				
		
		'If it is a search query then show no results advise
		If Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 Then
	
			'// Row - No current record	for search terms		
			Call BuildTableContNoRecord(4, "search")
					
		'Else show general no record information
		Else
	
			'// Row - No current record			
			Call BuildTableContNoRecord(4, "standard")
					
		End If
				
		
		'//	Chiudi layout ed informazioni
		Response.Write(vbCrLf & "</table>")

'/Condizione presenza dati
End If


		'Reset Server Objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing

		%>		  
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		<%

		'// Row - Data output panels
		Response.Write(vbCrLf & "<tr class=""smalltext"" align=""center"" valign=""top""><td colspan=""4"" height=""25""><br />")
		Call GoToPeriod("serp.asp", "")
		Call GoToGrouping("serp.asp", "")
		Call SearchForData("serp.asp", "", "Query|Engine")
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
Response.Write("</td></tr>")
Response.Write("</table>")
Response.Write("</td></tr>")
Response.Write("</table>")

%>
<!-- footer -->
<!--#include file="includes/footer.asp" -->
</body></html>