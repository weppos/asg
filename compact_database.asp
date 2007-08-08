<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<!--#include file="includes/inc_array_table.asp" -->
<!--#include file="lib/functions_filesystem.asp" -->
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


'Inizializza variabili
Dim strAsgMsg				'Holds the reset message 
Dim blnAsgDatabaseOptimized	'Set to true if the database has been compacted/optimized
Dim i


if Request.QueryString("compact") = 1 then
	
	' Optimize MySQL database
	If ASG_USE_MYSQL then
		
		' Optimize mySQL database
		for i = 1 to Ubound(aryAsgTable)
			blnAsgDatabaseOptimized = databaseMySqlOptimize(aryAsgTable(i, 1), objAsgConn)
		next
			
		' Reset objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing
	
	' Compact Access database
	else
		
		' Reset objects
		Set objAsgRs = Nothing
		objAsgConn.Close
		Set objAsgConn = Nothing
		
		strAsgMsg = strAsgMsg & vbCrLf & "<p>" & TXT_Db_weight & "&nbsp;" & TXT_BeforeCompating & "&nbsp;: " & getFileSize(strAsgMapPath) & "<br />"
		
		' Compact database
		blnAsgDatabaseOptimized = CompactAccessDatabase()
		
		' Replace the old database with the compacted one
		Call RinominaFile(strAsgMapPathTo, strAsgMapPath)
		' Call RipristinaFile(strAsgMapPathTo, strAsgMapPath)
	
		strAsgMsg = strAsgMsg & vbCrLf & "" & TXT_Db_weight & "&nbsp;" & TXT_AfterCompating & "&nbsp;: " & getFileSize(strAsgMapPath) & "</p>"
		
	end if
	
	' In case of errors show them
	if err <> 0 then 
		 strAsgMsg = strAsgMsg & "<p><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/no.png"" alt=""" & TXT_Error_Occured & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Error_Occured & ": <br />" & err.description & "</p>"
	' Show execution information
	elseif blnAsgDatabaseOptimized then
		' Notify a message depending on the database
		if ASG_USE_MYSQL then
			strAsgMsg = strAsgMsg & "<p><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Db_mysql_optimized & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Db_mysql_optimized & "</p>"
		else
			strAsgMsg = strAsgMsg & "<p><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Db_access_compacted & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Db_access_compacted & "<br /><span class=""notetext"">" & strAsgMapPathTo & "</span></p>"
			strAsgMsg = strAsgMsg & "<p><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Db_access_renamed & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Db_access_renamed & "<br /><span class=""notetext"">" & strAsgMapPath & "</span></p>"
		end If
	end if 

end if

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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Administration & " &raquo; " & MENUSECTION_Maintenance & " &raquo;</span> " & MENUSECTION_CompactDatabase %></div>
		<div id="layout_content">

<form action="compact_database.asp?compact=1" name="frmCompact" method="post">
<%

' :: Open tlayout :: MENUSECTION_CompactDatabase
Response.Write(builTableTlayout("", "open", MENUSECTION_CompactDatabase))
	
	
if Request.QueryString("compact") <> 1 then
	
	' Information about the compact batch
	strAsgTmpLayer = "<p style=""text-align: justify;"">" & TXT_Db_compact_info & "</p>" 
	strAsgTmpLayer = strAsgTmpLayer & "<p>" & TXT_Db_compact_info2 & "</p>" 

		' :: Create the layer ::
		Response.Write(buildLayer("layerCompact", TXT_Info, "", strAsgTmpLayer))

else
	
	' Information about the compact batch
	strAsgTmpLayer = strAsgMsg 

		' :: Create the layer ::
		Response.Write(buildLayer("layerCompact", LABEL_Exec_Report, "", strAsgTmpLayer))

end if

if Request.QueryString("compact") <> 1 then
	' Submit form area
	Response.Write("<div class=""submitarea""><input type=""submit"" name=""settings"" value=""" & TXT_CompactAndOptimize & """ /></div>")
end if

' :: Open tlayout :: MENUSECTION_CompactDatabase
Response.Write(builTableTlayout("", "close", ""))

%>
</form>

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