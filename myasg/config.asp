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


' Include HTTP functions
%><!--#include file="asg-lib/http.asp" --><%

' Include Update functions
%><!--#include file="asg-lib/update.asp" --><%

' Include Datetime functions
%><!--#include file="asg-lib/datetime.asp" --><%

' Include Binary functions
' TODO: consider to include the file only when necessary
%><!--#include file="asg-lib/binary.asp" --><%

' Include Layout functions
%><!--#include file="asg-lib/layout.asp" --><%

' This file is required to keep compatibility with releases < 2.2
' TODO: remove as soon as skin support is completely dropped
%><!--#include file="asg-config/oldskin.asp" --><%


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
