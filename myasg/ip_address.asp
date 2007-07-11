<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Option Explicit %>
<!--#include file="config.asp" -->
<!--#include file="includes/functions_images.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
' Copyright 2003-2006 - Carletti Simone										'
'-------------------------------------------------------------------------------'
'																				'
'	Autore:																		'
'	--------------------------													'
'	Simone Carletti (weppos)													'
'																				'
'	Collaboratori 																'
'	[che ringrazio vivamente per l'impegno ed il tempo dedicato]				'
'	--------------------------													'
'	@ imente 			- www.imente.it | www.imente.org						'
'	@ ToroSeduto		- www.velaforfun.com									'
'																				'
'	Hanno contribuito															'
'	[anche a loro un grazie speciale per le idee apportate]						'
'	--------------------------													'
'	@ Gli utenti del forum con consigli e segnalazioni							'
'	@ subxus (suggerimento generazione grafica dei report)						'
'																				'
'	Verifica le proposte degli utenti, implementate o da implementare al link	'
'	http://www.weppos.com/forum/forum_posts.asp?TID=140&PN=1					'
'																				'
'-------------------------------------------------------------------------------'
'																				'
'	Informazioni sulla Licenza													'
'	--------------------------													'
'	Questo è un programma gratuito; potete modificare ed adattare il codice		'
'	(a vostro rischio) in qualsiasi sua parte nei termini delle condizioni		'
'	della licenza che lo accompagna.											'
'																				'
'	Non è consentito utilizzare l'applicazione per conseguire ricavi 			'
'	personali, distribuirla, venderla o diffonderla come una propria 			'
'	creazione anche se modificata nel codice, senza un esplicito e scritto 		'
'	consenso dell'autore.														'
'																				'
'	Potete modificare il codice sorgente (a vostro rischio) per adattarlo 		'
'	alle vostre esigenze o integrarlo nel sito; nel caso le funzioni possano	'
'	essere di utilità pubblica vi invitiamo a comunicarlo per poterle 			'
'	implementare in una futura versione e per contribuire allo sviluppo 		'
'	del programma.																'
'																				'
'	In nessun caso l'autore sarà responsabile di danni causati da una 			'
'	modifica, da un uso non corretto o da un uso qualsiasi 						'
'	dell'applicazione.															'
'																				'
'	Nell'utilizzo devono rimanere intatte tutte le informazioni sul 			'
'	copyright; è possibile modificare o rimuovere unicamente le indicazioni 	'
'	espressamente specificate.													'
'																				'
'	Numerose ore sono state impiegate nello sviluppo del progetto e, anche 		'
'	se non vincolante ai fini dell'uso, sarebbe gratificante l'inserimento		'
'	di un link all'applicazione sul vostro sito.								'
'																				'
'	NESSUNA GARANZIA															'
'	------------------------- 													'
'	Questo programma è distribuito nella speranza che possa essere utile ma 	'
'	senza GARANZIA DI ALCUN GENERE.												'
'	L'utente si assume tutte le responsabilità nell'uso.						'
'																				'
'-------------------------------------------------------------------------------'

'********************************************************************************'
'*																				*'	
'*	VIOLAZIONE DELLA LICENZA													*'
'*	 																			*'
'*	L'utilizzo dell'applicazione violando le condizioni di licenza comporta la 	*'
'*	perdita immediata della possibilità d'uso ed è PERSEGUIBILE LEGALMENTE!		*'
'*																				*'
'********************************************************************************'


'// ATTENZIONE! Protezione statistiche.
'	Modificare solo se necessario e se sicuri.
'	Impostazioni errate possono compromettere la privacy.
Call AllowEntry("True", "False", "False", intAsgSecurity)


'Dichiara Variabili
Dim mese			'Riferimento per output
Dim gruppo			'
Dim dettagli		'IP di cui mostrare le informazioni
Dim asgOutput
Dim intAsgCount		'Conteggio record


'Read setting variables from querystring
mese = Request.QueryString("mese")
dettagli = Request.QueryString("dettagli")
strAsgSortBy = Request.QueryString("sort")
strAsgSortOrder = Request.QueryString("order")


'If period variable is empty then set it to the current month
If mese = "" Then mese = dtmAsgMonth & "-" & dtmAsgYear
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
	Case "ip" strAsgSortByFld = "IP"
	Case "data" strAsgSortByFld = "MAX(Last_Access)"
	Case Else strAsgSortByFld = "SUM(Visits)"
End Select

'Richiama le Dichiarazioni per la 
'paginazione avanzata
Call DimPaginazioneAvanzata()

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2004 Carletti Simone" />
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
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtIPAddress) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td width="5%"  background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" height="16"></td>
			<td width="35%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtIP) %>
				<a href="ip_address.asp?<%= "mese=" & mese & "&dettagli=" & dettagli & "&sort=ip&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtIPAddress & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="ip_address.asp?<%= "mese=" & mese & "&dettagli=" & dettagli & "&sort=ip&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtIPAddress & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="31%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtLastAccess) %>
				<a href="ip_address.asp?<%= "mese=" & mese & "&dettagli=" & dettagli & "&sort=data&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtLastAccess & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="ip_address.asp?<%= "mese=" & mese & "&dettagli=" & dettagli & "&sort=data&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtLastAccess & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="12%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmHits) %>
				<a href="ip_address.asp?<%= "mese=" & mese & "&dettagli=" & dettagli & "&sort=hits&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtHits & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="ip_address.asp?<%= "mese=" & mese & "&dettagli=" & dettagli & "&sort=hits&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtHits & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="12%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"><%= UCase(strAsgTxtSmVisits) %>
				<a href="ip_address.asp?<%= "mese=" & mese & "&dettagli=" & dettagli & "&sort=visits&order=DESC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtVisits & "&nbsp;" & strAsgTxtDesc %>">
				<img src="<%= strAsgSknPathImage%>arrow_down.gif" border="0" align="absmiddle" /></a>
				<a href="ip_address.asp?<%= "mese=" & mese & "&dettagli=" & dettagli & "&sort=visits&order=ASC" %>" title="<%= strAsgTxtOrderBy & "&nbsp;" & strAsgTxtVisits & "&nbsp;" & strAsgTxtAsc %>">
				<img src="<%= strAsgSknPathImage%>arrow_up.gif" border="0" align="absmiddle" /></a></td>
			<td width="5%" background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>"></td>
		  </tr>
		<%
		
		'Initialise SQL string to select data
		strAsgSQL = "SELECT IP, Max(Last_Access) AS MaxData, Sum(Visits) AS SumVisits, Sum(Hits) AS SumHits FROM "&strAsgTablePrefix&"IP "
		'Call the function to search into the database if there are enought information to do that
		strAsgSQL = CheckSearchForData(strAsgSQL, true)
		'Group information by following fields
		strAsgSQL = strAsgSQL & " GROUP BY IP "
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
					Call BuildTableContNoRecord(6, "search")
					
				'Else show general no record information
				Else
	
					'// Row - No current record			
					Call BuildTableContNoRecord(6, "standard")
					
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
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<%
									
				'Write an anchor
				Response.Write(vbCrLf & "<a name=""" & objAsgRs("IP") & """></a>")
	
				'Espandi Dettagli
				'// Link
				Response.Write(vbCrLf & "				<a href=""ip_address.asp?dettagli=" & objAsgRs("IP") & "&mese=&page=" & page & "&searchfor=" & asgSearchfor & "&searchin=" & asgSearchin & "&sort=" & strAsgSortBy & "&order=" & strAsgSortOrder & "#" & objAsgRs("IP") & """ title=""" & strAsgTxtShowIpInformation & """>")
				'// Icona espansa se Corrisponde
				If Trim(dettagli) <> "" AND objAsgRs("IP") = Trim(dettagli) Then
					Response.Write(vbCrLf & "				<img src=""" & strAsgSknPathImage & "expanded.gif"" alt=""" & strAsgTxtShowIpInformation & """ border=""0"" align=""absmiddle"" />")
				'// Icona espandi se Differente
				Else
					Response.Write(vbCrLf & "				<img src=""" & strAsgSknPathImage & "expand.gif"" alt=""" & strAsgTxtShowIpInformation & """ border=""0"" align=""absmiddle"" />")
				End If
				'// Chiudi Link
				Response.Write("</a>&nbsp;")
				
				'Mostra solo se Loggato
				If Session("AsgLogin") = "Logged" Then
	
					'Icona Filter IP
					Call ShowIconFilterIp(objAsgRs("IP"))
						
				End If
				
				%>
				<a href="JavaScript:openWin('popup_tracking_ip.asp?IP=<%= objAsgRs("IP") %>','profile','toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=425')" class="linksmalltext" title="<%= strAsgTxtIPTracking %>"><%= HighlightSearchKey(objAsgRs("IP"), "IP") %></a>
			</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><%= FormatOutTimeZone(objAsgRs("MaxData"), "Date") %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= objAsgRs("SumHits") %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= objAsgRs("SumVisits") %></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center">
			<% 
				
			'Tracking IP
			'// Link PopUp
			Response.Write(vbCrLf & "				<a href=""JavaScript:openWin('popup_tracking_ip.asp?IP=" & objAsgRs("IP") & "','Tracking','toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=425')"" title=""" & strAsgTxtIPTracking & """>")
			'// Icona espansa se Corrisponde
			Response.Write(vbCrLf & "				<img src=""" & strAsgSknPathImage & "tracking.gif"" alt=""" &  strAsgTxtIPTracking & """ border=""0"" />")
			'// Chiudi Link PopUp
			Response.Write("</a>")

			%>
			</td>
		  </tr>
		<% 
			If Trim(dettagli) <> "" AND objAsgRs("IP") = Trim(dettagli) Then
				
				Dim objAsgRs2
				
				'Mostra le query al motore
				Set objAsgRs2 = Server.CreateObject("ADODB.Recordset")
				strAsgSQL = "SELECT * FROM "&strAsgTablePrefix&"IP WHERE IP = '" & dettagli & "' "
				strAsgSQL = strAsgSQL & " ORDER BY Visits DESC, Hits DESC"
		
		
		%>
		  <tr class="smalltext">
			<td colspan="7"><br />
				<!-- Contenitore Dettagli -->
				<table width="100%" border="0" cellspacing="0" cellpadding="1" align="center">
				  <tr>
					<td width="7%" valign="top" align="center"><img src="<%= strAsgSknPathImage %>openarrow.gif" width="25" height="25" border="0" alt="<%= strAsgTxtDetails %>"></td>
					<td width="86%">
					<!-- Dettagli IP -->
					<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
					  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
						<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" width="40%" height="16"><%= UCase(strAsgTxtIP) %></td>
						<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" width="30%"><%= UCase(strAsgTxtLastAccess) %></td>
						<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" width="15%"><%= UCase(strAsgTxtSmHits) %></td>
						<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" width="15%"><%= UCase(strAsgTxtSmVisits) %></td>
					  </tr>
					  <% 
					  objAsgRs2.Open strAsgSQL, objAsgConn
						If objAsgRs2.EOF Then
							
							'// Row - No current record			
							Call BuildTableContNoRecord(4, "standard")
							
						Else
							Do While NOT objAsgRs2.EOF

					  %>
					  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
						<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left" height="16">&nbsp;<%= objAsgRs2("IP") %></td>
						<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center"><%= FormatOutTimeZone(objAsgRs2("Last_Access"), "Date") %></td>
						<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= objAsgRs2("Hits") %></td>
						<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= objAsgRs2("Visits") %></td>
					  </tr>
					  <%
							objAsgRs2.MoveNext
							Loop
						End If
								
						'// Row - End table spacer			
						Call BuildTableContEndSpacer(4)
				
					  objAsgRs2.Close
					  Set objAsgRs2 = Nothing
					  %>
					</table><br />
					<!-- Fine Dettagli IP -->
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
		Call BuildTableContEndSpacer(7)

		'// Row - Advanced data sorting
		Response.Write(vbCrLf & "<tr class=""smalltext""><td colspan=""7"" align=""center""><br />")
		Call PaginazioneAvanzata("ip_address.asp", "")
		Response.Write(vbCrLf & "<br /><br /></td></tr>")
	
		objAsgRs.Close
		
		'Reset Server Objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing

		'// Row - Data output panels
		Response.Write(vbCrLf & "<tr class=""smalltext"" align=""center"" valign=""top""><td colspan=""7"" height=""25""><br />")
		Call SearchForData("ip_address.asp", "", "IP")
		Response.Write(vbCrLf & "</td></tr>")
		%>		  
		</table><br />
<%

'Footer
Response.Write(vbCrLf & "		<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
'// Row - Footer Border Line
Call BuildFooterBorderLine()

'***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
'***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write(vbCrLf & "		  <tr align=""center"" valign=""middle"">")
Response.Write(vbCrLf & "			<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer"">ASP Stats Generator [" & strAsgVersion & "] - &copy; 2003-2006 <a href=""http://www.weppos.com/"" class=""linkfooter"">weppos</a>")
If blnAsgElabTime Then Response.Write(" - " & strAsgTxtThisPageWasGeneratedIn & "&nbsp;" & FormatNumber(Timer() - startAsgElab, 4) & "&nbsp;" & strAsgTxtSeconds)
Response.Write(						"</td>")
Response.Write(vbCrLf & "		  </tr>")
'***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
'***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write(vbCrLf & "		</table>")
Response.Write(vbCrLf & "	  </td></tr>")
Response.Write(vbCrLf & "	</table>")
Response.Write(vbCrLf & "  </td></tr>")
Response.Write(vbCrLf & "</table>")

%>
<!--#include file="includes/footer.asp" -->
</body></html>