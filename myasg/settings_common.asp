<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Option Explicit %>
<!--#include file="config.asp" -->
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
Call AllowEntry("False", "False", "False", intAsgSecurity)


'Inserimento record
If Request.Form("Settings") = strAsgTxtUpdate AND Request.QueryString("act") = "upd" Then

	'Pulisci e controlla URL
	strAsgSiteURLremote = Trim(Request.Form("URLremote"))
		If Left(strAsgSiteURLremote, 7) <> "http://" Then strAsgSiteURLremote = "http://" & strAsgSiteURLremote
		If Right(strAsgSiteURLremote, 1) <> "/" Then strAsgSiteURLremote = strAsgSiteURLremote & "/"
	strAsgSiteURLlocal = Trim(Request.Form("URLlocal"))
		If Left(strAsgSiteURLlocal, 7) <> "http://" Then strAsgSiteURLlocal = "http://" & strAsgSiteURLlocal
		If Right(strAsgSiteURLlocal, 1) <> "/" Then strAsgSiteURLlocal = strAsgSiteURLlocal & "/"
	'Richiama dati da Form
	strAsgSiteName = FilterSQLInput(Trim(Server.HTMLEncode(Request.Form("SiteName"))))
	strAsgSiteEmail = Trim(Request.Form("SiteEmail"))
	
	strAsgStartHits = Clng(Trim(Request.Form("StartHits")))
	strAsgStartVisits = Clng(Trim(Request.Form("StartVisits")))
	'Momentaneamente disabilitato per incongruenza periodi accavallati
	'strAsgTimeZone = Request.Form("serverTimeZonePosition") & Request.Form("serverTimeZoneValue") & "|" & Request.Form("gmtTimeZonePosition") & Request.Form("gmtTimeZoneValue")
	strAsgTimeZone = Request.Form("serverTimeZonePosition") & Request.Form("serverTimeZoneValue") & "|+0"
	
	'Filter data
	'Change from Bolean to Int() type to use data with MySQL
	blnRefererServer = CInt(CBool(Request.Form("RefererServer")))
	blnStripPathQS = CInt(CBool(Request.Form("strAsgIPPathQS")))
	blnMonitReferer = CInt(CBool(Request.Form("MonitReferer")))
	blnMonitDaily = CInt(CBool(Request.Form("MonitDaily")))
	blnMonitIP = CInt(CBool(Request.Form("MonitIP")))
	blnMonitHourly = CInt(CBool(Request.Form("MonitHourly")))
	blnMonitSystem = CInt(CBool(Request.Form("MonitSystem")))
	blnMonitLanguages = CInt(CBool(Request.Form("MonitLanguages")))
	blnMonitPages = CInt(CBool(Request.Form("MonitPages")))
	blnMonitEngine = CInt(CBool(Request.Form("MonitEngine")))
	blnMonitCountry = CInt(CBool(Request.Form("MonitCountry")))

	blnAsgCheckIcon = CInt(CBool(Request.Form("CheckIcon")))
	

	'Initialise SQL string to update values
	strAsgSQL = "UPDATE "&strAsgTablePrefix&"config SET " &_
	"Sito_Nome = '" & strAsgSiteName & "', " &_
	"Sito_URL_Remoto = '" & strAsgSiteURLremote & "', " &_
	"Sito_URL_Locale = '" & strAsgSiteURLlocal & "', " &_
	"Sito_Email = '" & strAsgSiteEmail & "', " &_
	"Start_Hits = " & strAsgStartHits & ", " &_
	"Start_Visits = " & strAsgStartVisits & ", " &_
	"Time_Zone = '" & strAsgTimeZone & "', " &_
	"Opt_Referer_Server = " & blnRefererServer & ", " &_
	"Opt_Strip_Path_QS = " & blnStripPathQS & ", " &_
	"Opt_Monit_Referer = " & blnMonitReferer & ", " &_
	"Opt_Monit_Daily = " & blnMonitDaily & ", " &_
	"Opt_Monit_IP = " & blnMonitIP & ", " &_
	"Opt_Monit_Hourly = " & blnMonitHourly & ", " &_
	"Opt_Monit_System = " & blnMonitSystem & ", " &_
	"Opt_Monit_Languages = " & blnMonitLanguages & ", " &_
	"Opt_Monit_Pages = " & blnMonitPages & ", " &_
	"Opt_Monit_Engine = " & blnMonitEngine & ", " &_
	"Opt_Monit_Country = " & blnMonitCountry & ", " &_
	"Opt_Check_Icon = " & blnAsgCheckIcon & " "

	'Execute the update
	objAsgConn.Execute(strAsgSQL)
	'Response.Write(strAsgSQL) : Response.End()
	
	'Se si utilizzano le variabili Application aggiornale
	If blnApplicationConfig Then
						
		'Aggiorna Variabili Application
		Application("strAsgSiteName") = strAsgSiteName
		Application("strAsgSiteURLremote") = strAsgSiteURLremote
		Application("strAsgSiteURLlocal") = strAsgSiteURLlocal
		Application("strAsgSiteEmail") = strAsgSiteEmail
	
		Application("strAsgStartHits") = CLng(strAsgStartHits)
		Application("strAsgStartVisits") = CLng(strAsgStartVisits)
		Application("strAsgTimeZone") = strAsgTimeZone
	
		Application("blnRefererServer") = CBool(blnRefererServer)
		Application("blnStripPathQS") = CBool(blnStripPathQS)
		Application("blnMonitReferer") = CBool(blnMonitReferer)
		Application("blnMonitDaily") = CBool(blnMonitDaily)
		Application("blnMonitIP") = CBool(blnMonitIP)
		Application("blnMonitHourly") = CBool(blnMonitHourly)
		Application("blnMonitSystem") = CBool(blnMonitSystem)
		Application("blnMonitLanguages") = CBool(blnMonitLanguages)
		Application("blnMonitPages") = CBool(blnMonitPages)
		Application("blnMonitEngine") = CBool(blnMonitEngine)
		Application("blnMonitCountry") = CBool(blnMonitCountry)

		Application("blnAsgCheckIcon") = CBool(blnAsgCheckIcon)

		'Forza il ricalcolo delle Application
		Application("blnConfig") = False
	
	End If
	
	'Reset Server Objects
	Set objAsgRs = Nothing
	objAsgConn.Close
	Set objAsgConn = Nothing
	
	'Reindirizza per rivalorizzare dati
	Response.Redirect("settings_common.asp")

End If

'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2004 Carletti Simone" />
<link href="stile.css" rel="stylesheet" type="text/css" />

<!-- 	ASP Stats Generator <%= strAsgVersion %> è una applicazione gratuita 
		per il monitoraggio degli accessi e dei visitatori ai siti web 
		creata e sviluppata da Simone Carletti.
		
		Puoi scaricarne una copia gratuita sul sito ufficiale http://www.weppos.com/ -->

</head>
<!--#include file="includes/header.asp" -->
		<form action="settings_common.asp?act=upd" name="frmImpostazioni" method="post">
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtGeneralSettings) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtConfigSettings) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right" width="30%"><%= strAsgTxtSiteName %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left" width="70%">&nbsp;<input type="text" name="SiteName" value="<%= strAsgSiteName %>" size="60" maxlength="140" /></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtSiteURLlocal %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<input type="text" name="URLlocal" value="<% If "[]" & strAsgSiteURLlocal = "[]" Then Response.Write("http://") Else Response.Write(strAsgSiteURLLocal) %>" size="60" maxlength="140" /></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtSiteURLremote %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<input type="text" name="URLremote" value="<% If "[]" & strAsgSiteURLremote = "[]" Then Response.Write("http://") Else Response.Write(strAsgSiteURLremote) %>" size="60" maxlength="140" /></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtSiteEmail %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<input type="text" name="SiteEmail" value="<%= strAsgSiteEmail %>" size="60" maxlength="140" /></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtCountSettings) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%" align="right"><%= strAsgTxtStartVisits %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="70%" align="left">&nbsp;<input type="text" name="StartVisits" value="<%= strAsgStartVisits %>" size="10" maxlength="8" /></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtStartHits %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<input type="text" name="StartHits" value="<%= strAsgStartHits %>" size="10" maxlength="8" /></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtDateSettings) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="30%" align="right"><%= strAsgTxtTimeZoneOffSet %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="70%" align="left">&nbsp;<select name="serverTimeZonePosition" class="normalform">
					<option value="+" <% If Left(aryAsgTimeZone(0), 1) = "+" Then Response.Write("selected") %>>+</option>
					<option value="-" <% If Left(aryAsgTimeZone(0), 1) = "-" Then Response.Write("selected") %>>-</option>
				</select>
				<select name="serverTimeZoneValue" class="normalform">
					<% For looptmp = 0 to 23 %>
					<option value="<%= looptmp %>" <% If Cint(Right(aryAsgTimeZone(0), Len(aryAsgTimeZone(0))-1)) = looptmp Then Response.Write("selected") %>><%= looptmp %></option>
					<% Next %>
				</select>
				<%= strAsgTxtOffSetClientServer %><br />&nbsp;&nbsp;<%= strAsgTxtServerDateTimeAre & ":&nbsp;<span class=""notetext"">" & Now() & "</span>" %>
			</td>
		  </tr>
		  <!-- Momentaneamente disabilitato per incongruenza periodi accavallati!
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtTimeZoneOffSet %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">&nbsp;<select name="gmtTimeZonePosition" class="normalform">
					<option value="+" <% If Left(aryAsgTimeZone(1), 1) = "+" Then Response.Write("selected") %>>+</option>
					<option value="-" <% If Left(aryAsgTimeZone(1), 1) = "-" Then Response.Write("selected") %>>-</option>
				</select>
				<select name="gmtTimeZoneValue" class="normalform">
					<% For looptmp = 0 to 23 %>
					<option value="<%= looptmp %>" <% If Cint(Right(aryAsgTimeZone(1), Len(aryAsgTimeZone(1))-1)) = looptmp Then Response.Write("selected") %>><%= looptmp %></option>
					<% Next %>
				</select>
				<%= strAsgTxtOffSetGMTtoUser %><br />&nbsp;&nbsp;<%= strAsgTxtReportDateTimeAre & ":&nbsp;<span class=""notetext"">" & FormatOutTimeZone(dtmAsgNow, "Date") & "&nbsp;" & FormatOutTimeZone(dtmAsgNow, "Time") & "</span>" %>
			</td>
		  </tr>
		  -->
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtMonitSettings) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtMonitOptions %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">
				&nbsp;<input type="checkbox" name="RefererServer" value="True" <% If blnRefererServer Then Response.Write "checked" %> /> <%= strAsgTxtCountServerAsReferer %><br />
				&nbsp;<input type="checkbox" name="strAsgIPPathQS" value="True" <% If blnStripPathQS Then Response.Write "checked" %> /> <%= strAsgTxtstrAsgIPPathQS %>
			</td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"><%= strAsgTxtEnableMonit %>: &nbsp;&nbsp;</td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">
				&nbsp;<input type="checkbox" name="MonitReferer" value="True" <% If blnMonitReferer Then Response.Write "checked" %> /> <%= strAsgTxtReferer %><br />
				&nbsp;<input type="checkbox" name="MonitDaily" value="True" <% If blnMonitDaily Then Response.Write "checked" %> /> <%= strAsgTxtDailyMonit %><br />
				&nbsp;<input type="checkbox" name="MonitHourly" value="True" <% If blnMonitHourly Then Response.Write "checked" %> /> <%= strAsgTxtHourlyMonit %><br />
				&nbsp;<input type="checkbox" name="MonitIP" value="True" <% If blnMonitIP Then Response.Write "checked" %> /> <%= strAsgTxtIPAddress %> <br />
				&nbsp;<input type="checkbox" name="MonitSystem" value="True" <% If blnMonitSystem Then Response.Write "checked" %> /> <%= strAsgTxtSystems & ": " & strAsgTxtBrowser & ", " & strAsgTxtOS & ", " & strAsgTxtColor & ", " & strAsgTxtReso %><br />
				&nbsp;<input type="checkbox" name="MonitLanguages" value="True" <% If blnMonitLanguages Then Response.Write "checked" %> /> <%= strAsgTxtBrowserLanguages %><br />
				&nbsp;<input type="checkbox" name="MonitPages" value="True" <% If blnMonitPages Then Response.Write "checked" %> /> <%= strAsgTxtHits %><br />
				&nbsp;<input type="checkbox" name="MonitEngine" value="True" <% If blnMonitEngine Then Response.Write "checked" %> /> <%= strAsgTxtSearchEngine & " " & strAsgTxtAnd & " " & strAsgTxtSearchQuery %><br />
				&nbsp;<input type="checkbox" name="MonitCountry" value="True" <% If blnMonitCountry Then Response.Write "checked" %> /> <%= strAsgTxtCountry %>
			</td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtProgramTools) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="right"></td>
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="left">
				&nbsp;<input type="checkbox" name="CheckIcon" value="True" <% If blnAsgCheckIcon Then Response.Write "checked" %> /> <%= strAsgTxtReportUnknownIcons %>
			</td>
		  </tr>
		  <%
					
			'// Row - End table spacer			
			Call BuildTableContEndSpacer(2)
	
		  %>
		  <tr class="normaltitle">
			<td colspan="2" align="center"><br /><input type="submit" name="Settings" value="<%= strAsgTxtUpdate %>" /></td>
		  </tr>
		</table><br />
		</form>

		<!-- write an anchor. It could be usefult in future -->
		<a name="monitoringstring"></a>
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" align="center" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" colspan="2" align="center" height="16"><%= UCase(strAsgTxtUsingApplication) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" align="center" colspan="2">
				<%= strAsgTxtPageMonitoringString %><br />
				<textarea name="monitoringstring" cols="80" rows="3">&lt;script type="text/javascript" language="JavaScript" src="http://<%= Request.ServerVariables("HTTP_HOST") & Left(Request.ServerVariables("URL"), InStrRev(Request.ServerVariables("URL"), "/")-1) %>/stats_js.asp"&gt; &lt;/script&gt;
				</textarea>
			</td>
		  </tr>
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