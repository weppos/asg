<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'-------------------------------------------------------------------------------'
'																				'
'	Questo file è stato creato per consentire l'aumento delle prestazioni in	'
'	elaborazione ed allo stesso tempo ovviare al problema dell'uso dei percorsi	'
'	relativi a path superiori su Win2003 Server.								'
'																				'
'	Qui saranno/sono inclusi tutti i file che necessiteranno di essere eseguiti	'
'	dall'applicazione generale ma non al file di conteggio mentre il file 		'
'	config_common.asp è stato dedicato all'uso principale dal file di conteggio.'
'	In questo modo il file di conteggio non sarà obbligato a caricare in		'
'	memoria inutili variabili di elaborazione utili solo al processo di report	' 
'	rallenterebbero l'applicazione.												'
'																				'
'-------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------'
' !! All files included in this include need a folder with write permission !!
'-------------------------------------------------------------------------------'
%><!--#include file="w2k3_write_permission.asp" --><%

'-------------------------------------------------------------------------------'
' Include config settings
'-------------------------------------------------------------------------------'
%><!--#include file="w2k3_config_common.asp" --><%

'-------------------------------------------------------------------------------'
' Include stats functions
'-------------------------------------------------------------------------------'
%><!--#include file="lib/functions_stats.asp" --><%

'-------------------------------------------------------------------------------'
' Include layout functions
'-------------------------------------------------------------------------------'
%><!--#include file="lib/utils.layout.asp" --><%

'-------------------------------------------------------------------------------'
' Include file functions
'-------------------------------------------------------------------------------'
%><!--#include file="lib/functions_filesystem.asp" --><%

'-------------------------------------------------------------------------------'
' 
'-------------------------------------------------------------------------------'
%><!--#include file="lib/utils.security.asp" --><%

'-------------------------------------------------------------------------------'
' 
'-------------------------------------------------------------------------------'
%><!--#include file="lib/utils.search.asp" --><%
	
'-------------------------------------------------------------------------------'
' Include language translation
'-------------------------------------------------------------------------------'
%><!--#include file="lang/default/common.asp" --><%


'---------------------------------------------------
'	Common variables
'---------------------------------------------------

' Show toolbar
Dim blnAsgShowToolbar	' Set to true to show the toolbar
blnAsgShowToolbar = true

	
	'---------------------------------------------------
	'	Dimension variables : show icons
	'---------------------------------------------------
	Dim strAsgIconaTemp
	Dim Index
	
	'---------------------------------------------------
	'	Dimension variables : sorting records
	'---------------------------------------------------
	Dim strAsgSortBy
	Dim strAsgSortByFld
	Dim strAsgSortOrder
	
	'---------------------------------------------------
	'	Dimension variables : report data search engine
	'---------------------------------------------------
	Dim strAsgSQLsearchstring			'Holds the search string to query database
	Dim asgSearchfor					'Holds the keywords to search for
	Dim asgSearchIn						'Holds the name of the table to search in
	
	'---------------------------------------------------
	'	Dimension variables : other elaborations
	'---------------------------------------------------
	Dim strAsgUnknownIcon				'Holds unknown icons information
	
	'---------------------------------------------------
	'	Controllo aggiornamenti
	'---------------------------------------------------
	Dim strAsgLastVersion		'
	Dim dtmAsgLastUpdate		'
	Dim intAsgLastUpdate		'Casiste possibili:
		intAsgLastUpdate =	0 	'  - variabile non in calcolo
								'1 - tutto combacia
								'2 - differente versione e data
								'3 - uguale versione ma differente data
	Dim urlAsgLastUpdate		'


' Check setuplock file
Dim intAsgSetupLock
' Ignore checking
if not ASG_SETUPLOCK then
	intAsgSetupLock = 2
else
	' The file doesn't exist
	if not file_exists(STR_ASG_PATH_FOLDER_WR & ASG_COOKIE_PREFIX & ASG_SETUPLOCK_FILE) then
		intAsgSetupLock = 0
	' The file exists
	else
		intAsgSetupLock = 1
	end if
end if


'---------------------------------------------------
'	Apri connessione al Database
'---------------------------------------------------
If blnAsgConnIsOpen = False Then
	'Se si usano variabili application apri la connessione
	'dato che non è stata aperta per gestire le risorse
	'del file di conteggio.
	objAsgConn.Open strAsgConn
'---------------------------------------------------
End If
'---------------------------------------------------


'-----------------------------------------------------------------------------------------
' Check version for update!
'-----------------------------------------------------------------------------------------
' Controllo differenza data ed esecuzione solo se amministratore
If Clng(Clng(Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)) - appAsgCheckUpdate) > 7 AND Session("asgLogin") = "Logged" Then

	'Esegui controllo versione
'	Call CheckUpdate(ASG_VERSION, ASG_VERSION_BUILD)
	
	'Aggiornamento informazioni sul controllo
	strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "config SET conf_check_update = " & Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & ""
	objAsgConn.Execute(strAsgSQL)
	
	'Se si utilizzano le variabili Application aggiornale
	If blnApplicationConfig Then
						
		'Aggiorna Variabili Application
		Application(ASG_APPLICATION_PREFIX & "CheckUpdate") = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)

	End If
	
End If

%>
