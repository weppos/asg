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


' Include HTTP functions
%><!--#include file="asg-lib/http.asp" --><%

' Include Update functions
%><!--#include file="asg-lib/update.asp" --><%

' Include Datetime functions
%><!--#include file="asg-lib/datetime.asp" --><%



' *** Update checker variables ***

Dim strAsgLatestVersion
Dim dtmAsgLatestUpdate
Dim intAsgLatestUpdate
Dim urlAsgLatestUpdate

intAsgLatestUpdate = 0 ' by default disable any alert



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
' Includi il file di skin!
'-------------------------------------------------------------------------------'
'
' 									>>	 DA ADATTARE!	<<
'
' Enter the path of the folder where the skin file 
' has been uploaded.

%><!--#include file="mdb/inc_skin_file.asp" --><%
'
'
'-------------------------------------------------------------------------------'


'-------------------------------------------------------------------------------'
' Includi le informazioni generiche di configurazione
'-------------------------------------------------------------------------------'
%><!--#include file="config_common.asp" --><%

'-------------------------------------------------------------------------------'
' Includi le informazioni di gestione dei report statistici
'-------------------------------------------------------------------------------'
%><!--#include file="includes/functions_stats.asp" --><%

'-------------------------------------------------------------------------------'
' Includi le informazioni sullo sviluppo del layout tramite funzioni
'-------------------------------------------------------------------------------'
%><!--#include file="includes/functions_layout.asp" --><%

	
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
	

'-------------------------------------------------------------------------------'
' Includi le informazioni sulla traduzione in uso
'-------------------------------------------------------------------------------'

'//	Italiano %>
<!--include file="languages/italiano.asp" --><%

'//	English %>
<!--#include file="languages/english.asp" --><%

'//	Espanol %>
<!--include file="languages/espanol.asp" --><%


'---------------------------------------------------
'	Apri connessione al Database
'---------------------------------------------------
If blnConnectionIsOpen = False Then
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
if Clng(asgDatestamp(Now()) - blnAsgCheckUpdate) > 7 and Session("AsgLogin") = "Logged" then

  Dim aryAsgLatestVersion
  aryAsgLatestVersion = asgVersionCheck(strAsgVersion)
  
  if Ubound(aryAsgLatestVersion) > 0 then
    strAsgLatestVersion = aryAsgLatestVersion(0)
    dtmAsgLatestUpdate  = aryAsgLatestVersion(1)
    urlAsgLatestUpdate  = aryAsgLatestVersion(2)
    
    ' compare versions and display alert in case of greather release
    if strAsgLatestVersion > strAsgVersion then
      intAsgLatestUpdate = 1
    end if
     
  end if

  ' update database
  strAsgSQL = "UPDATE "&strAsgTablePrefix&"Config SET Opt_Check_Update = " & asgDatestamp(Now())
  objAsgConn.Execute(strAsgSQL)

  ' update application config if enabled
  if blnApplicationConfig then
    Application("blnAsgCheckUpdate") = asgDatestamp(Now())
  end if
	
end if

%>
