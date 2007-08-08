<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright 2003-2004 - Carletti Simone										'
'-------------------------------------------------------------------------------'
'																				'
'	Author:																		'
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
'	QuefileTo è un programma gratuito; potete modificare ed adattare il codice		'
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
'	QuefileTo programma è distribuito nella speranza che possa essere utile ma 	'
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


'// WARNING! Program protection.
'	Changing default values may allow users to access the page.
Call checkPermission("False", "False", "False", appAsgSecurity)

' Dimension variables
Dim collItem			' Holds the object of the Server.Variables collection
Dim blnAsgServerInfo	' Set to true to show server info
Dim blnAsgServerVars	' Set to true to show server variables collection
Dim strLayout
Dim strAsgServerOs
Dim blnAsgServerBad	' Set to true if the server os is not good to run the application

blnAsgServerBad = false

' Get settings from querystring
if Request.QueryString("servinfo") = 1 then
	blnAsgServerInfo = true
else
	blnAsgServerInfo = false
end if

' Get settings from querystring
if Request.QueryString("servars") = 1 then
	blnAsgServerVars = true
else
	blnAsgServerVars = false
end if

' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

%>
<%= STR_ASG_PAGE_DOCTYPE %>
<html>
<head>
<title><%= appAsgSiteName %> | powered by ASP Stats Generator v<%= ASG_VERSION %></title>
<%= STR_ASG_PAGE_CHARSET %>
<meta name="copyright" content="Copyright (C) 2003-2005 Carletti Simone" />
<!--#include file="includes/meta.inc.asp" -->

<!-- ASP Stats Generator v. <%= ASG_VERSION %> is created and developed by Simone Carletti.
To download your Free copy visit the official site http://www.weppos.com/asg/ -->

</head>

<body>
<!--#include file="includes/header.asp" -->

<div align="center">
	<div id="layout">
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Tools & " &raquo;</span> " & MENUSECTION_ServerInfo %></div>
		<div id="layout_content">

		<!-- content table -->
		<table class="tlayout_border" cellpadding="5" cellspacing="1" border="0" width="100%" align="center">
		<tr>
			<td class="tlayout_cat"><%= MENUSECTION_ServerInfo %></td>
		</tr>
		<tr>
			<td style="padding:0px">

<%

if blnAsgServerInfo then

	' VbsEngine layer
	strAsgTmpLayer = "<p>" & getScriptEngineInfo() & "</p>"
	' Create the layer
	Response.Write(buildLayer("layerVbs", TXT_VbsEngine, "", strAsgTmpLayer))
	
	' Server Os layer
	strAsgTmpLayer = "<p>"
	' List of possibile server OS
	if Instr(Request.ServerVariables("SERVER_SOFTWARE"), "IIS/4") > 0 then
		strAsgServerOs = "Microsoft Windows NT"
		blnAsgServerBad = true
	elseif Instr(Request.ServerVariables("SERVER_SOFTWARE"), "IIS/5.0") > 0 then
		strAsgServerOs = "Microsoft Windows 2000"
	elseif Instr(Request.ServerVariables("SERVER_SOFTWARE"), "IIS/5.1") > 0 then
		strAsgServerOs = "Microsoft Windows XP"
	elseif Instr(Request.ServerVariables("SERVER_SOFTWARE"), "IIS/6") > 0 then
		strAsgServerOs = "Microsoft Windows 2003"
	end if
	
	strAsgTmpLayer = strAsgTmpLayer & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "def/os.asp?icon=" & strAsgServerOs & """ align=""middle"" alt=""" & strAsgServerOs & """ />&nbsp;" & strAsgServerOs & "</p>"
	if blnAsgServerBad then strAsgTmpLayer = strAsgTmpLayer & "<p class=""errortext"">" & TXT_Server_bados_warning & "</p>"
	' Create the layer
	Response.Write(buildLayer("layerServerOs", TXT_OS, "", strAsgTmpLayer))

	
	strLayout = ""
	strLayout = strLayout & vbCrLf & "<div id=""layerServerInfo"" style=""display: block;"">"
	strLayout = strLayout & vbCrLf & "<fieldset class=""fldlayer""><legend class=""fldlegendtext""><span class=""fldlegendtitle"">" & MENUSECTION_ServerInfo & " </span></legend>"
	Response.Write(strLayout)

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <% 
	for each collItem in Request.ServerVariables() 
		' Only server info
		if Instr(collItem, "SERVER_") > 0 then
  %>
  <tr>
	<td class="tfieldset_col" width="30%"><strong><%= collItem %></strong></td>
	<td class="tfieldset_col"><%= ShareWords(Request.ServerVariables(collItem), 80) %>&nbsp;</td>
  </tr>
  <% 
  		end if
	next %>
</table>
<%
	
	strLayout = ""
	strLayout = strLayout & vbCrLf & "</fieldset>"
	strLayout = strLayout & "</div>"
	Response.Write(strLayout)

end if

' Server variables
if blnAsgServerVars then

	strLayout = ""
	strLayout = strLayout & vbCrLf & "<div id=""layerServerVars"" style=""display: block;"">"
	strLayout = strLayout & vbCrLf & "<fieldset class=""fldlayer""><legend class=""fldlegendtext""><span class=""fldlegendtitle"">" & MENUSECTION_ServerVariables & " </span></legend>"
	Response.Write(strLayout)

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
  <% for each collItem in Request.ServerVariables() %>
  <tr>
	<td class="tfieldset_col" width="30%"><strong><%= collItem %></strong></td>
	<td class="tfieldset_col"><%= ShareWords(Request.ServerVariables(collItem), 80) %>&nbsp;</td>
  </tr>
  <% next %>
</table>
<%
	strLayout = ""
	strLayout = strLayout & vbCrLf & "</fieldset>"
	strLayout = strLayout & "</div>"
	Response.Write(strLayout)

end if

%>

			
			</td>
		</tr>
		</table>
		<!-- / content table -->

		</div>
	</div>
</div>

<br /></div>
<!-- / body -->
<%

' Footer
Response.Write(vbCrLf & "<div id=""footer"">")
' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write(vbCrLf & "<br /><div style=""text-align: center;"">Powered by <a href=""http://www.weppos.com/asg/"" title=""ASP Stats Generator"">ASP Stats Generator</a> v" & ASG_VERSION & " ") 
if ASG_BUILDINFO then Response.Write("build " & ASG_VERSION_BUILD)
Response.Write(vbCrLf & "<br />Copyright &copy; 2003-2005 <a href=""http://www.weppos.com/"">weppos</a></div>")
if ASG_ELABORATION_TIME then Response.Write("<div class=""elabtime"">" & Replace(TXT_elabtime, "$time$", FormatNumber(Timer() - startAsgElab, 4)) & "</div>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******
Response.Write(vbCrLf & "</div>")

%>
<!--#include file="includes/footer.asp" -->
</body></html>