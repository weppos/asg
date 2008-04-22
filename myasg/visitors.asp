<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="config.asp" -->
<!--#include file="includes/functions_count.asp" -->
<!--#include file="asg-lib/file.asp" -->
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
Call AllowEntry("True", "False", "False", intAsgSecurity)


'Dichiara Variabili
Dim mese				'Riferimento per output
Dim dettagli				'Hold details info to expand
Dim asgOutputPage		'Pagina di output
Dim blnDatabaseIsEmpty
Dim intAsgRecordCountHits
Dim strAsgSearchInTmp
Dim blnAsgShowDetails		'Set true if you must show visitors details
Dim strAsgActiveRange		'Hold the active time range	ot the selected visitor		


'Check if there's a field to search in and write the name in a tmp variable
If Len(Trim(Request.QueryString("searchin"))) > 0 Then strAsgSearchInTmp = Trim(Request.QueryString("searchin"))


'Read setting variables from querystring
dettagli = Request.QueryString("dettagli")


'Richiama le Dichiarazioni per la 
'paginazione avanzata
Call DimPaginazioneAvanzata()


If Request.QueryString("showall") = "true" then
	RecordsPerPage = 5
	blnAsgShowDetails = True
Else
	RecordsPerPage = 15
	blnAsgShowDetails = False
End If

blnDatabaseIsEmpty = True

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
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtVisitorsDetails) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
	<%

	'Adatta temporaneamente la Path
	strAsgSiteURLremote = Right(strAsgSiteURLremote, Len(strAsgSiteURLremote)-7)
	
	
	'Read the value of the field to search in and adapt the name to a SQL query friendly use
	Select Case strAsgSearchInTmp
		Case "Browser"	asgSearchin = "MAX(Browser)"
		Case "OS"		asgSearchin = "MAX(OS)"
		Case "Reso"		asgSearchin = "MAX(Reso)"
		Case "Color"	asgSearchin = "MAX(Color)"
		Case "IP"		asgSearchin = "MAX(IP)"
		Case "Country"	asgSearchin = "MAX(Country)"
	End Select

	'Initialise SQL string to select data
	strAsgSQL = "SELECT Visitor_ID FROM "&strAsgTablePrefix&"Detail "
	'Call the function to search into the database if there are enought information to do that
	strAsgSQL = CheckSearchForData(strAsgSQL, true)
	'Group information by following fields and order by the most recent date/time
	strAsgSQL = strAsgSQL & " GROUP BY Visitor_ID ORDER BY MAX(Data) DESC"
	
	'Return the original value
	asgSearchin = strAsgSearchInTmp
		
	'Prepara il Rs
	objAsgRs.CursorType = 3
	objAsgRs.LockType = 3
		
	'Apri il Rs
	objAsgRs.Open strAsgSQL, objAsgConn
			
		'Il Rs  vuoto
		If objAsgRs.EOF Then
				
			Response.Write vbCrLf & "<!-- no data -->"
			Response.Write vbCrLf & "<table width=""90%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">"
			Response.Write vbCrLf & "<tr bgcolor=""" & strAsgSknTableTitleBgColour & """ align=""center"">"
			Response.Write vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ height=""16"" class=""smalltext"">" & strAsgTxtNoRecordInDatabase & "</td>"
			Response.Write vbCrLf & "</tr>"
			Response.Write vbCrLf & "</table><br />"
			Response.Write vbCrLf & "<!-- / no data -->"

		Else
			
			blnDatabaseIsEmpty = False
			objAsgRs.PageSize = RecordsPerPage
			objAsgRs.AbsolutePage = page

			Dim objAsgRs2
			Set objAsgRs2 = Server.CreateObject("ADODB.Recordset")

			For PaginazioneLoop = 1 To RecordsPerPage
				
				If Not objAsgRs.EOF Then			
				
				'Informazioni
				strAsgSQL = "SELECT TOP 1 * FROM "&strAsgTablePrefix&"Detail WHERE Visitor_ID = '" & objAsgRs("Visitor_ID") & "' "
				objAsgRs2.Open strAsgSQL, objAsgConn
					
											
				'Write an anchor
				Response.Write(vbCrLf & "<a name=""" & objAsgRs("Visitor_ID") & """></a>")
		
		%>		  
		<!-- Visitatore -->
				<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
				  <tr><td><table border="0" cellpadding="0" cellspacing="1" width="100%">
				  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
					<td width="25%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"><%= UCase(strAsgTxtBrowser) %></td>
					<td width="25%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtOS) %></td>
					<td width="25%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtReso) %></td>
					<td width="25%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtIPAddress) %></td>
				  </tr>
				  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
					<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><%= HighlightSearchKey(objAsgRs2("Browser"), "Browser") %></td>
					<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><%= HighlightSearchKey(objAsgRs2("OS"), "OS") %></td>
					<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><%= HighlightSearchKey(objAsgRs2("Reso"), "Reso") & " - " & HighlightSearchKey(objAsgRs2("Color"), "Color") & " bit" %></td>
					<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center">
						<a href="JavaScript:openWin('popup_tracking_ip.asp?IP=<%= objAsgRs2("IP") %>','Tracking','toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=425')" class="linksmalltext" title="<%= strAsgTxtIPTracking %>"><%= HighlightSearchKey(objAsgRs2("IP"), "IP") %></a>&nbsp;
						<% 

							
						'Tracking IP
						'// Link PopUp
						Response.Write(vbCrLf & "						<a href=""JavaScript:openWin('popup_tracking_ip.asp?IP=" & objAsgRs2("IP") & "','Tracking','toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=425')"" title=""" & strAsgTxtIPTracking & """>")
						'// Icona espansa se Corrisponde
						Response.Write(vbCrLf & "						<img src=""" & strAsgSknPathImage & "tracking_small.gif"" alt=""" &  strAsgTxtIPTracking & """ border=""0"" align=""absmiddle"" />")
						'// Chiudi Link PopUp
						Response.Write("</a>&nbsp;")
							
									
						'Mostra solo se Loggato
						If Session("AsgLogin") = "Logged" Then
							
							'Icona Filter IP
							Call ShowIconFilterIp(objAsgRs2("IP"))

						End If
									
						%>
						</td>
				  </tr></table></td></tr>
				  <tr><td><table border="0" cellpadding="0" cellspacing="1" width="100%">
				  <tr>
					<td width="25%" align="center" bgcolor="<%= strAsgSknTableTitleBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16" class="normaltitle"><%= UCase(strAsgTxtCountry) %></td>
					<td width="75%" align="left" bgcolor="<%= strAsgSknTableContBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" class="smalltext"><img src="<%= asgFlagIcon("asg-includes/images/icons/flags/", objAsgRs2("Country2")) %>" border="0" align="absmiddle" />&nbsp;<%= HighlightSearchKey(objAsgRs2("Country"), "Country") %></td>
				  </tr>
				<% If objAsgRs2("Referer") <> "(unknown)" AND objAsgRs2("Engine") <> "" Then %>
				  <tr>
					<td align="center" bgcolor="<%= strAsgSknTableTitleBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16" class="normaltitle"><%= UCase(strAsgTxtSearchEngine) %></td>
					<td align="left" bgcolor="<%= strAsgSknTableContBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" class="smalltext"><img src="images/engine.asp?icon=<%= objAsgRs2("Engine") %>" alt="<%= objAsgRs2("Engine") %>" align="absmiddle" height="14" width="14" /> <%= objAsgRs2("Engine") %></td>
				  </tr>
				  <tr>
					<td align="center" bgcolor="<%= strAsgSknTableTitleBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16" class="normaltitle"><%= UCase(strAsgTxtSearchQuery) %></td>
					<td align="left" bgcolor="<%= strAsgSknTableContBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" class="smalltext"><% = objAsgRs2("Query") %></td>
				  </tr>
				  <tr>
					<td align="center" bgcolor="<%= strAsgSknTableTitleBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16" class="normaltitle"><%= UCase(strAsgTxtReferer) %></td>
					<td align="left" bgcolor="<%= strAsgSknTableContBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" class="smalltext"><a href="<% = objAsgRs2("Referer") %>" title="<%= strAsgTxtGoToPage %>" class="linksmalltext" target="_blank"><%= StripValueTooLong(objAsgRs2("Referer"), 85, 40, 40) %></a></td>
				  </tr>
				<% ElseIf objAsgRs2("Referer") <> "(unknown)" Then %>
				  <tr>
					<td align="center" bgcolor="<%= strAsgSknTableTitleBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16" class="normaltitle"><%= UCase(strAsgTxtReferer) %></td>
					<td align="left" bgcolor="<%= strAsgSknTableContBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" class="smalltext"><a href="<% = objAsgRs2("Referer") %>" title="<%= strAsgTxtGoToPage %>" class="linksmalltext" target="_blank"><%= StripValueTooLong(objAsgRs2("Referer"), 85, 40, 40) %></a></td>
				  </tr>
				<% End If 
				
				
				objAsgRs2.Close
					
					
				'Initialise SQL string to count records
				strAsgSQL = "SELECT COUNT(Details_ID) FROM "&strAsgTablePrefix&"Detail WHERE Visitor_ID = '" & objAsgRs("Visitor_ID") & "'"
				'Open Rs
				objAsgRs2.Open strAsgSQL, objAsgConn
				'Set the number of total hits
				If Not objAsgRs2.EOF Then 
					intAsgRecordCountHits = objAsgRs2(0)
				Else
					intAsgRecordCountHits = 0
				End If
				'Close Rs
				objAsgRs2.Close


				If dettagli = objAsgRs("Visitor_ID") OR blnAsgShowDetails Then
					

				'Initialise SQL string to update values
				strAsgSQL = "SELECT * FROM "&strAsgTablePrefix&"Detail WHERE Visitor_ID = '" & objAsgRs("Visitor_ID") & "' ORDER BY Data DESC "
					
				objAsgRs2.CursorType = 1
				objAsgRs2.Open strAsgSQL, objAsgConn

				'Build Layout
				Response.Write vbCrLf & "<tr bgcolor=""" & strAsgSknTableTitleBgColour & """ align=""center"" class=""normaltitle"">"
				Response.Write vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableTitleBgImage & """ height=""15"">" & UCase(strAsgTxtDate) & "</td>"
				Response.Write vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableTitleBgImage & """>" & intAsgRecordCountHits & "&nbsp;" & UCase(strAsgTxtHits) & "</td>"
				Response.Write vbCrLf & "</tr>"

					Do While NOT objAsgRs2.EOF
						
				  %>
				  <tr bgcolor="<%= strAsgSknTableContBgColour %>" class="smalltext">
					<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><%= FormatOutTimeZone(objAsgRs2("Data"), "Date") & "&nbsp;" & strAsgTxtAt & "&nbsp;" & FormatOutTimeZone(objAsgRs2("Data"), "Time") %></td>
					<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">
					<%
					  
					Response.Write vbCrLf & "<a class=""linksmalltext"" href=""" & objAsgRs2("Page") & """ target=""_blank"" title=""" & objAsgRs2("Page") & """>"

					'Verifica la pagina e mostra o meno
					'una icona standard di corrispondenza dominio.
					Response.Write(ChooseDomainIcon(objAsgRs2("Page"), "visitors" ))
						
					'TAGLIA STRINGHE LUNGHE
					'Se la stringa supera i 55 caratteri inserisci ... in mezzo e accorcia
					'Caso "Nessun Raggruppamento" - Max 55 Caratteri
					Response.Write(StripValueTooLong(asgOutputPage, 55, 25, 25))
	
					Response.Write vbCrLf & "</a></td>"
					
				%>
					</td>
				  </tr>
				<%
					
					objAsgRs2.MoveNext
					Loop
				objAsgRs2.Close

				Else
				
				'Reset the variable
				strAsgActiveRange = Null
				
				'Initialise SQL string to get first e last record
				strAsgSQL = "SELECT Data FROM "&strAsgTablePrefix&"Detail WHERE Visitor_ID = '" & objAsgRs("Visitor_ID") & "' ORDER BY Data "
				'Set Rs properties
				objAsgRs2.CursorType = 1
				'Open Rs
				objAsgRs2.Open strAsgSQL, objAsgConn
				'Set the number of total hits
				If Not objAsgRs2.EOF Then 
					
					'Get the first visited page
					strAsgActiveRange = strAsgTxtFrom & "&nbsp;" & FormatOutTimeZone(objAsgRs2("Data"), "Date") & "&nbsp;" & strAsgTxtAt & "&nbsp;" & FormatOutTimeZone(objAsgRs2("Data"), "Time")
					
					'Move to the last visited page
					objAsgRs2.MoveLast
					
					'Get the first visited page
					strAsgActiveRange = strAsgActiveRange & "&nbsp;" & strAsgTxtTo & "&nbsp;" & FormatOutTimeZone(objAsgRs2("Data"), "Date") & "&nbsp;" & strAsgTxtAt & "&nbsp;" & FormatOutTimeZone(objAsgRs2("Data"), "Time")
				
				End If
				'Close Rs
				objAsgRs2.Close

				'Build Layout
				Response.Write(vbCrLf & "<tr>")
				Response.Write(vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableTitleBgImage & """ bgcolor=""" & strAsgSknTableTitleBgColour & """ align=""center"" height=""15"" class=""normaltitle"">" & intAsgRecordCountHits & "&nbsp;" & UCase(strAsgTxtHits))
				Response.Write(vbCrLf & "<td background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ bgcolor=""" & strAsgSknTableContBgColour & """ align=""left"" class=""smalltext"">")
				If Len(strAsgActiveRange) Then Response.Write(strAsgActiveRange)
				Response.Write(vbCrLf & "</td>")
				Response.Write(vbCrLf & "</tr>")
				
				
				%>
				  <tr>
					<td align="right" bgcolor="<%= strAsgSknTableContBgColour %>" background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" class="smalltext" colspan="2"><a href="visitors.asp?<%= "dettagli=" & objAsgRs("Visitor_ID") & "&page=" & page & "#" & objAsgRs("Visitor_ID") & "" %>" title="<%= strAsgTxtDetails %>" class="linksmalltext"><%= strAsgTxtFullVersion %>
					<img src="images/arrow_small_dx.gif" alt="<%= strAsgTxtFullVersion %>" align="middle" border="0"></a></td>
				  </tr>
				<%
				
				' / Show user details if thevariable is true
				End If
					

				%>
				  </table></td></tr>
				<%
					  
					'// Row - End table spacer			
					Call BuildTableContEndSpacer(2)
					  
				%>
				</table><br />
		<!-- Fine Visitatore -->
		<%
				objAsgRs.MoveNext
				End If
			Next
			Set objAsgRs2 = Nothing
		End If
	
	
	'If database contains records to show call advanced data sorting
	If blnDatabaseIsEmpty = False Then
		
		'// Row - Advanced data sorting
		Response.Write(vbCrLf & "<div class=""smalltext"" align=""center"">")
		Call PaginazioneAvanzata("visitors.asp", "")
		Response.Write(vbCrLf & "<br /><br />")
		Call SearchForData("visitors.asp", "", "Browser|OS|Reso|Color|IP|Country")
		Response.Write(vbCrLf & "</div><br />")

	End If

	objAsgRs.Close
		
	'Reset Server Objects
	Set objAsgRs = Nothing
	objAsgConn.Close
	Set objAsgConn = Nothing

	%>		  
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
