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


' Inizializza variabili
Dim strAsgTable					' Holds the value of the table to reset
Dim strAsgTimerange
Dim strAsgWeekrange
Dim strAsgMsg						' Holds the reset message 
Dim blnAsgDisallow
Dim dtmAsgResetDateTime			' Holds a DateTime variable used to reset data
Dim blnAsgDatabaseOptimized	' Set to true if the database has been compacted/optimized
Dim i


'-----------------------------------------------------------------------------------------
' Reset tabella
'-----------------------------------------------------------------------------------------
' Function:	Reset della tabella in base a parametri passati
' Date: 	03.04.2004
' Comment:		
'-----------------------------------------------------------------------------------------
public function deleteTableData(databasetable)
	

	' Create the SQL string to delete information depending on database
	If ASG_USE_MYSQL then
		strAsgSQL = "DELETE FROM " & ASG_TABLE_PREFIX & databasetable & " "
	else
		strAsgSQL = "DELETE * FROM " & ASG_TABLE_PREFIX & databasetable & " "
	end if
	
	'-----------------------------------------------------------------------------------------
	' Weekly or particular data reset
	'-----------------------------------------------------------------------------------------
	if databasetable = "Detail" AND IsNumeric(strAsgWeekrange) then
			
			' Allow reset
			blnAsgDisallow = false
			' Cast the value to int
			strAsgWeekrange = CInt(strAsgWeekrange)
			' Calculate reset date
			dtmAsgResetDateTime = DateAdd("ww", -strAsgWeekrange, dtmAsgDate)
			dtmAsgResetDateTime = Year(dtmAsgResetDateTime) & "/" & Month(dtmAsgResetDateTime) & "/" & Day(dtmAsgResetDateTime)
			' Build condition line depending on the database
			if ASG_USE_MYSQL then
				strAsgSQL = strAsgSQL & "WHERE Detail_date < '" & dtmAsgResetDateTime & "' "
			else
				strAsgSQL = strAsgSQL & "WHERE Detail_date < #" & dtmAsgResetDateTime & "# "
			end if

	end If 


	'-----------------------------------------------------------------------------------------
	' Monthly or normal data reset
	'-----------------------------------------------------------------------------------------
	' Reset all tables
	if strAsgTimerange = "full" then

		' Allow reset
		blnAsgDisallow = false
	
	' Reset data olther than current month
	elseif strAsgTimerange = "0" then
		
		' Controllo coerenza reset mensile
		if databasetable <> "Detail" AND databasetable <> "IP" then 
			strAsgSQL = strAsgSQL & "WHERE Mese <> '" & dtmAsgMonth & "-" & dtmAsgYear & "' "
			' Allow reset
			blnAsgDisallow = false
		
		elseif databasetable = "Detail" then 
			' Calculate reset date
			dtmAsgResetDateTime = Date() 'DateAdd("m", -1, dtmAsgDate)
			dtmAsgResetDateTime = Year(dtmAsgResetDateTime) & "/" & Month(dtmAsgResetDateTime) & "/" & 1
			' Build condition line depending on the database
			If ASG_USE_MYSQL then
				strAsgSQL = strAsgSQL & "WHERE Detail_date < '" & dtmAsgResetDateTime & "' "
			else
				strAsgSQL = strAsgSQL & "WHERE Detail_date < #" & dtmAsgResetDateTime & "# "
			end if
			' Allow reset
			blnAsgDisallow = false
		Else
			' Disallow reset
			blnAsgDisallow = true
		End If
	
	' Reset data older than selected periodi
	elseif IsNumeric(strAsgTimerange) AND CInt(strAsgTimerange) > 0 then
		
		Dim dateloop
		
		'Controllo coerenza reset mensile
		if databasetable <> "Detail" AND databasetable <> "IP" then 
			
			' Allow reset
			blnAsgDisallow = false
			' Build condition line depending on the database
			strAsgSQL = strAsgSQL & "WHERE Mese <> '" & Right("0" & Month(dtmAsgDate), 2) & "-" & Year(dtmAsgDate) & "' "

			' Cast the value to int
			strAsgTimerange = Cint(strAsgTimerange)
			
			For dateloop = 1 to strAsgTimerange
				'Calcola la differenza di tempo
				dtmAsgDate = DateAdd("m", -1, dtmAsgDate)
				'Prepara la condizione
				strAsgSQL = strAsgSQL & "AND Mese <> '" & Right("0" & Month(dtmAsgDate), 2) & "-" & Year(dtmAsgDate) & "' "
			Next

		elseIf databasetable = "Detail" then 
			
			'Allow reset execution
			blnAsgDisallow = False
			'Trasforma il valore in numerico
			strAsgTimerange = Cint(strAsgTimerange)
			' Calculate reset date
			dtmAsgResetDateTime = DateAdd("m", -strAsgTimerange, dtmAsgDate)
			dtmAsgResetDateTime = Year(dtmAsgResetDateTime) & "/" & Month(dtmAsgResetDateTime) & "/" & Day(dtmAsgResetDateTime)
			' Build condition line depending on the database
			If ASG_USE_MYSQL then
				strAsgSQL = strAsgSQL & "WHERE Detail_date < '" & dtmAsgResetDateTime & "' "
			else
				strAsgSQL = strAsgSQL & "WHERE Detail_date < #" & dtmAsgResetDateTime & "# "
			end if

		else
			' Disallow reset
			blnAsgDisallow = true
		end If
		
	else
			' Disallow reset
			blnAsgDisallow = true
	end If

	' Response.Write(strAsgSQL) : Response.End()
	' If the reset is allowed execute it
	if not blnAsgDisallow then
		objAsgConn.Execute(strAsgSQL)
		strAsgMsg = strAsgMsg & "<tr>" &_
			"<td align=""right"">" &_
				"<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Deldata_Completed & """ /> " &_
				TXT_Table & "&nbsp;<span class=""notetext"">" & databasetable & "</span>:</td>" &_
			"<td align=""left"">" & TXT_Deldata_completed & "</td>" &_
		"</tr>"
	end If

end function

' Get reset options from querystring
strAsgTable = Trim(Request.QueryString("table"))
strAsgTimerange = Trim(Request.QueryString("timerange"))
strAsgWeekrange = Trim(Request.QueryString("weekrange"))

' On Error Resume Next


' Reset all tables
If strAsgTable = 0 Then
	
	' Loop all tables to reset data	
	for i = 1 to Ubound(aryAsgTable)
		Call deleteTableData(aryAsgTable(i, 1))
	next

	' Complete the layout
	strAsgMsg = "<table align=""center"" border=""0"" cellpadding=""3"" cellspacin=""1"">" & strAsgMsg & "</table>"
	
' Single reset
elseif strAsgTable <> 0 AND Len(strAsgTable) > 0 then
	
	' Reset just the selected table
	Call deleteTableData(aryAsgTable(strAsgTable, 1))

	' Complete the layout
	strAsgMsg = "<table align=""center"" border=""0"" cellpadding=""3"" cellspacin=""1"">" & strAsgMsg & "</table>"
	
' No action	
else
	
	' Reset objects
	Set objAsgRs = Nothing
	objAsgConn.Close
	Set objAsgConn = Nothing
	Response.Redirect("batch_delete_old_data.asp?msg=error")

end if


' Optimize MySQL database
if ASG_USE_MYSQL then
	
	' Optimize all tables
	if strAsgTable = 0 then
			
		for i = 1 to Ubound(aryAsgTable)
			blnAsgDatabaseOptimized = databaseMySqlOptimize(aryAsgTable(i, 1), objAsgConn)
		next
		
	' Optimize selected table
	elseif strAsgTable <> 0 AND Len(strAsgTable) > 0 then
			blnAsgDatabaseOptimized = databaseMySqlOptimize(aryAsgTable(strAsgTable, 1), objAsgConn)
	end If
		
	Set objAsgRs = Nothing
	objAsgConn.Close
	Set objAsgConn = Nothing

' Compact Access database
else
	
	Set objAsgRs = Nothing
	objAsgConn.Close
	Set objAsgConn = Nothing
	
	' Compact database
	blnAsgDatabaseOptimized = CompactAccessDatabase()
	
	' After compacting replace the old database
	' with the compacted one
	Call RinominaFile(strAsgMapPathTo, strAsgMapPath)
	' Call RipristinaFile(strAsgMapPathTo, strAsgMapPath)
	
end if


' In case of errors show them
if err <> 0 then 
    strAsgMsg = strAsgMsg & "<p><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/no.png"" alt=""" & TXT_Error_Occured & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Error_Occured & ": <br />" & err.description & "</p>"
' else show execution information
elseif blnAsgDatabaseOptimized then
	
	' Notify a message depending on the database
	if ASG_USE_MYSQL then
		strAsgMsg = strAsgMsg & "<p><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Db_mysql_optimized & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Db_mysql_optimized & "</p>"
	else
		strAsgMsg = strAsgMsg & "<p><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Db_access_compacted & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Db_access_compacted & "<br /><span class=""notetext"">" & strAsgMapPathTo & "</span></p>"
		strAsgMsg = strAsgMsg & "<p><img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/ok.png"" alt=""" & TXT_Db_access_renamed & """ border=""0"" align=""middle"" />&nbsp;" & TXT_Db_access_renamed & "<br /><span class=""notetext"">" & strAsgMapPath & "</span></p>"
	end If
	
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
		<div id="layout_menutitle"><%= "<span class=""menusubtitle"">" & MENUGROUP_Administration & " &raquo; " & MENUSECTION_Maintenance & " &raquo;</span> " & MENUSECTION_BatchDeleteOldData %></div>
		<div id="layout_content">

<%

' :: Open tlayout :: MENUSECTION_BatchDeleteOldData
Response.Write(builTableTlayout("", "open", MENUSECTION_BatchDeleteOldData))
	
		' :: Create the layer ::
		Response.Write(buildLayer("layerDelete", LABEL_Exec_Report, "", strAsgMsg))

' :: Open tlayout :: MENUSECTION_BatchDeleteOldData
Response.Write(builTableTlayout("", "close", ""))

%>

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