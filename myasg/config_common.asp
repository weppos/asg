<%

'/**
' * ASP Stats Generator - Powerful and reliable ASP website counter
' *
' * This file is part of the ASP Stats Generator package.
' * (c) 2003-2007 Simone Carletti <weppos@weppos.net>, All Rights Reserved
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
' * @copyright       2003-2007 Simone Carletti, All Rights Reserved
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id$
' */
 
'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


Server.ScriptTimeout = 90
Session.Timeout = 20
Response.Buffer = True

'Deve essere attivata l'opzione sottostante, nel caso che ci sia il CDate Error 
'1040 è la session per l'Italia
'Session.LCID = 1040

'Non salvare in cache
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


'---------------------------------------------------
'	Dichiara le Variabili Globali
'---------------------------------------------------

Dim strAsgSQL			'Stringa SQL di esecuzione
Dim objAsgConn			'Oggetto Connessione
Dim strAsgConn			'Stringa di connessione al db
Dim objAsgRs			'Oggetto RecordSet
Dim startAsgElab		'Inizio elaborazione 
Dim blnAsgIsVisit		'Imposta a vero se è una visita
Dim intAsgVisitValue	'Numero di aumento
Dim dtmAsgYear			'Anno
Dim dtmAsgMonth			'Mese
Dim dtmAsgDay			'Giorno
Dim dtmAsgNow			'Ora e giorno
Dim dtmAsgDate			'Data attuale


'---------------------------------------------------
'	Dichiara le Variabili Dettagli
'---------------------------------------------------

'// Dal Js
Dim strAsgResolution		'Risoluzione video	
Dim strAsgColor				'Profondità di Colore (Bit)
Dim strAsgPage				'Pagina attuale richiamata
Dim strAsgPageStripped		'Pagina attuale richiamata senza querystring
Dim strAsgReferer			'Referer
Dim strAsgFontSmoothing		'					-- NON IMPLEMENTATO ATTUALMENTE --
Dim strAsgJavaEnabled		'Java Abilitati		-- NON IMPLEMENTATO ATTUALMENTE --

'// Da asp
Dim strAsgIP				'IP User
Dim strAsgUA				'User Agent User
Dim strAsgSessionID

'// Per IP
Dim strCountry				' Nazione
Dim strCountry2				' Nazione con 2 caratteri

'// Da Classe Imente
Dim objClassI				'Oggetto classe
Dim strAsgBrowser			'Browser User
Dim strAsgBrowserLang		'Lingua del Browser
Dim	strAsgOS				'Sistema Operativo User
Dim strAsgBrowserActCookie	'Browser con accettazione dei cookie


'---------------------------------------------------
'	Dichiara le Variabili Elaborazione
'---------------------------------------------------
Dim blnExitCount			' Imposta a True se deve uscire dal monitoraggio
Dim blnExitEngine			' By Engine - Vero se deve uscire dal ciclo
Dim strAsgSingleIP			' By IP - IP splittati da filtrare
Dim strToStrip				' Var di funz					
Dim strBuffer				' Var di funz
Dim strAsgRefererDom		' By Referer - Dominio
Dim intLoop					' Variabile di Ciclo Elaborazione
Dim strAsgEngineName		' Nome del Motore
Dim strAsgEngineQS			' QS di ricerca
Dim strAsgSelectValue		' Valore della select in elaborazione
	
	
'---------------------------------------------------
'	Dichiara le Variabili Config
'---------------------------------------------------
Dim strAsgSiteName			'Nome del Sito
Dim strAsgSiteURLremote		'URL Web del Sito - In remoto
Dim strAsgSiteURLlocal		'URL Web del Sito - In locale
Dim strAsgSiteEmail			'Indirizzo Email del sito
Dim strAsgSitePsw			'Password di amministrazione
Dim dtmAsgStartStats		'Inizio monitoraggio
Dim intAsgSecurity			'Protezione pubblica delle statistiche
Dim strAsgStartHits			'Visite pagine di partenza
Dim strAsgStartVisits		'Visitatori di partenza
Dim strAsgImage				'Immagine da mostrare
Dim strAsgFilterIP			'IP da filtrare
Dim strAsgTimeZone			'Fusi orari
Dim aryAsgTimeZone			'Array Fusi orari
Dim blnRefererServer		'Considera il proprio server come referer
Dim blnStripPathQS			'Taglia la QS della pagina
Dim blnMonitReferer			'Monitoraggio Referer
Dim blnMonitDaily			'Monitoraggio Giornaliero
Dim blnMonitIP				'Monitoraggio IP
Dim blnMonitHourly			'Monitoraggio per Orari
Dim blnMonitSystem			'Monitoraggio Sistemi
Dim blnMonitLanguages		'Monitoraggio Lingue Browser
Dim blnMonitPages			'Monitoraggio Pagine Visitate
Dim blnMonitEngine			'Monitoraggio Motori di Ricerca
Dim blnMonitCountry			'Monitoraggio Nazioni di Provenienza
Dim blnAsgCheckIcon			'Notifica icone non riconosciute
Dim blnAsgCheckUpdate		'Data ultimo controllo aggiornamento
Dim blnApplicationConfig	'Utilizza o meno variabili di applicazione per settaggio
Dim blnConnectionIsOpen		'Imposta a true se ha dovuto procedere all'apertura della connessione


%><!--#include file="includes/inc_config.asp" --><%


'							========================================
'---------------------------	NON MODIFICARE I DATI SOTTOSTATI!	-------------------------------------
'							========================================

'-----------------------------------------------------------------------------------------
' Adatta il fuso orario inbound
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	31.03.2004 |
' Commenti:	
'-----------------------------------------------------------------------------------------
function FormatInTimeZone(ByVal datetimevalue, ByVal inTimeZone)
	
	'Modifica l'orario in base al fuso
	If Left(inTimeZone, 1) = "+" Then
		datetimevalue = DateAdd("h", inTimeZone, datetimevalue)
	ElseIf Left(inTimeZone, 1) = "-" Then
		datetimevalue = DateAdd("h", inTimeZone, datetimevalue)
	Else
		'
	End If
	'datetimevalue = DateAdd("h", + CInt(Right(inTimeZone, Len(inTimeZone-1))), datetimevalue)
	
	'Esegui i calcoli
	'dtmAsgNow = datetimevalue
	dtmAsgDay = Day(datetimevalue)
	dtmAsgMonth = Month(datetimevalue)
	dtmAsgYear = Year(datetimevalue)
	dtmAsgDate = Year(datetimevalue) & "/" & Month(datetimevalue) & "/" & Day(datetimevalue)
	dtmAsgNow = Year(datetimevalue) & "/" & Month(datetimevalue) & "/" & Day(datetimevalue) & " " & Hour(datetimevalue) & "." & Minute(datetimevalue) & "." & Second(datetimevalue)

	'Se l'ora è a 1 cifra inserisci uno 0 davanti
	If Len(dtmAsgDay) < 2 Then dtmAsgDay = "0" & dtmAsgDay
	If Len(dtmAsgMonth) < 2 Then dtmAsgMonth = "0" & dtmAsgMonth
	
end function 'Calcola elaborazione server
startAsgElab = Timer()

'Variabili debug e sviluppo
Dim strAsgVersion
strAsgVersion = "2.1.4-dev"
Dim dtmAsgUpdate
dtmAsgUpdate = "20070809"
Const blnAsgDebugMode = False


'Inizializza variabili
blnAsgIsVisit = True
blnExitCount = False
blnExitEngine = False
blnApplicationConfig = True
intAsgVisitValue = 1
strAsgSessionID = Session.SessionID
blnConnectionIsOpen = False

'Definisci la connessione
Set objAsgConn = Server.CreateObject("ADODB.Connection")

'Definisci il recordset
set objAsgRs = Server.CreateObject("ADODB.Recordset")



'---------------------------------------------------
'	Apri connessione al Database
'---------------------------------------------------
'Before 2.x version
'objAsgConn.Open strAsgConn

'---------------------------------------------------
'	Richiama la configurazione attiva
'---------------------------------------------------

'Setta predefinite alcune variabili
intAsgSecurity = 0
strAsgStartHits = 0
strAsgStartVisits = 0


'Verifica le variabili Application per configurazione
If isEmpty(Application("blnConfig")) OR isNull(Application("blnConfig")) OR Application("blnConfig") = false OR blnApplicationConfig = false Then
	
	
	'---------------------------------------------------
	'	Apri connessione al Database
	'---------------------------------------------------
	'Apri la connessione solo se necessaria per il 
	'richiamo dei dati.
	'Se i usano variabili application puoi evitare di 
	'aprirla per risparmiare risorse.
	objAsgConn.Open strAsgConn
	blnConnectionIsOpen = True
	'---------------------------------------------------
	
	
	'Inizializza SQL
	strAsgSQL = "SELECT TOP 1 * FROM "&strAsgTablePrefix&"Config"
	
	'Apri Recordset
	objAsgRs.Open strAsgSQL, objAsgConn
	If NOT objAsgRs.EOF Then
		
		strAsgSiteName = objAsgRs("Sito_Nome")
		strAsgSiteURLremote = objAsgRs("Sito_URL_Remoto")
		strAsgSiteURLlocal = objAsgRs("Sito_URL_Locale")
		strAsgSiteEmail = objAsgRs("Sito_Email")
		strAsgSitePsw = objAsgRs("Sito_Psw")
		dtmAsgStartStats = objAsgRs("Start_Stats")
		If IsNumeric(objAsgRs("Stats_Protezione")) Then intAsgSecurity = CInt(objAsgRs("Stats_Protezione"))
		strAsgImage = objAsgRs("Image")
		strAsgFilterIP = objAsgRs("Filter_IP")
		strAsgTimeZone = objAsgRs("Time_Zone")
		If IsNumeric(objAsgRs("Start_Hits")) Then strAsgStartHits = Clng(objAsgRs("Start_Hits"))
		If IsNumeric(objAsgRs("Start_Visits")) Then strAsgStartVisits = Clng(objAsgRs("Start_Visits"))
		blnRefererServer = CBool(objAsgRs("Opt_Referer_Server"))
		blnStripPathQS = CBool(objAsgRs("Opt_Strip_Path_QS"))
		blnMonitReferer	= CBool(objAsgRs("Opt_Monit_Referer"))
		blnMonitDaily	= CBool(objAsgRs("Opt_Monit_Daily"))
		blnMonitIP	= CBool(objAsgRs("Opt_Monit_IP"))
		blnMonitHourly	= CBool(objAsgRs("Opt_Monit_Hourly"))
		blnMonitSystem	= CBool(objAsgRs("Opt_Monit_System"))
		blnMonitLanguages = CBool(objAsgRs("Opt_Monit_Languages"))
		blnMonitPages	= CBool(objAsgRs("Opt_Monit_Pages"))
		blnMonitEngine	= CBool(objAsgRs("Opt_Monit_Engine"))
		blnMonitCountry = CBool(objAsgRs("Opt_Monit_Country"))
		blnAsgCheckIcon = CBool(objAsgRs("Opt_Check_Icon"))
		blnAsgCheckUpdate = CLng(objAsgRs("Opt_Check_Update"))
		
		'Se si utilizzano le variabili (ma sono vuote) allora creale
		If blnApplicationConfig Then
				
			'Blocca l'applicazione che potrà essere modificata solo da un utente
			Application.Lock
			
			'Inserisci le variabili lette dal database
			Application("strAsgSiteName") = strAsgSiteName
			Application("strAsgSiteURLremote") = strAsgSiteURLremote
			Application("strAsgSiteURLlocal") = strAsgSiteURLlocal
			Application("strAsgSiteEmail") = strAsgSiteEmail
			Application("strAsgSitePsw") = strAsgSitePsw
			Application("dtmAsgStartStats") = dtmAsgStartStats
			Application("intAsgSecurity") = CInt(intAsgSecurity)
			Application("strAsgImage") = strAsgImage
			Application("strAsgFilterIP") = strAsgFilterIP
			Application("strAsgTimeZone") = strAsgTimeZone
			Application("strAsgStartHits") = CLng(strAsgStartHits)
			Application("strAsgStartVisits") = CLng(strAsgStartVisits)
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
			Application("blnAsgCheckUpdate") = CLng(blnAsgCheckUpdate)
		
			'Imposta la variabile di config a vera
			Application("blnConfig") = True
				
			'Sblocca l'Application
			Application.UnLock
		
		'Fine condizione .EOF
		End If

	End If
	objAsgRs.Close
	
'Se le Application sono valorizzate ed attive usale per risparmiare elaborazione al db
ElseIf blnApplicationConfig Then

		strAsgSiteName = Application("strAsgSiteName")
		strAsgSiteURLremote = Application("strAsgSiteURLremote")
		strAsgSiteURLlocal = Application("strAsgSiteURLlocal")
		strAsgSiteEmail = Application("strAsgSiteEmail")
		strAsgSitePsw = Application("strAsgSitePsw")
		dtmAsgStartStats = Application("dtmAsgStartStats")
		intAsgSecurity = CInt(Application("intAsgSecurity")) 
		strAsgImage = Application("strAsgImage")
		strAsgFilterIP = Application("strAsgFilterIP")
		strAsgTimeZone = Application("strAsgTimeZone")
		strAsgStartHits = CLng(Application("strAsgStartHits"))
		strAsgStartVisits = CLng(Application("strAsgStartVisits")) 
		blnRefererServer = CBool(Application("blnRefererServer"))
		blnStripPathQS = CBool(Application("blnStripPathQS"))
		blnMonitReferer = CBool(Application("blnMonitReferer"))
		blnMonitDaily = CBool(Application("blnMonitDaily"))
		blnMonitIP = CBool(Application("blnMonitIP"))
		blnMonitHourly = CBool(Application("blnMonitHourly"))
		blnMonitSystem = CBool(Application("blnMonitSystem"))
		blnMonitLanguages = CBool(Application("blnMonitLanguages"))
		blnMonitPages = CBool(Application("blnMonitPages"))
		blnMonitEngine = CBool(Application("blnMonitEngine"))
		blnMonitCountry = CBool(Application("blnMonitCountry"))
		blnAsgCheckIcon = CBool(Application("blnAsgCheckIcon"))
		blnAsgCheckUpdate = CLng(Application("blnAsgCheckUpdate"))

'Fine condizioni variabili Application
End If


'-----------------------------------------------------------------------------------------
' Splitta il conenuto del campo una volta per tutte
' le esecuzioni future
  If Trim("[]" & strAsgTimeZone) = "[]" Then strAsgTimeZone = "+0|+0"
  aryAsgTimeZone = Split(strAsgTimeZone, "|")
'-----------------------------------------------------------------------------------------

'Imposta Fuso Orario in
Call FormatInTimeZone(Now(), aryAsgTimeZone(0))

%>
<!--#include file="includes/functions_common.asp"-->
<%

%>