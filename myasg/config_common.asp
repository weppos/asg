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


Server.ScriptTimeout = 90
Session.Timeout = 20
Response.Buffer = True

'Deve essere attivata l'opzione sottostante, nel caso che ci sia il CDate Error 
'1040 � la session per l'Italia
'Session.LCID = 1040

' Force this page to not be cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


' Elaboration start up
Dim asgTimerElabStart: asgTimerElabStart = Timer()

' Turn debug on/off
Dim asgDebug: asgDebug = false


' ***** Global variables ***** 

Dim strAsgSQL           'Stringa SQL di esecuzione
Dim objAsgConn          'Oggetto Connessione
Dim strAsgConn          'Stringa di connessione al db
Dim objAsgRs            'Oggetto RecordSet
Dim blnAsgIsVisit       'Imposta a vero se � una visita
Dim intAsgVisitValue    'Numero di aumento
Dim dtmAsgYear          'Anno
Dim dtmAsgMonth         'Mese
Dim dtmAsgDay           'Giorno
Dim dtmAsgNow           'Ora e giorno
Dim dtmAsgDate          'Data attuale


' ***** Tracking variables ***** 

' from Js
Dim strAsgResolution        'Risoluzione video  
Dim strAsgColor             'Profondit� di Colore (Bit)
Dim strAsgPage              'Pagina attuale richiamata
Dim strAsgPageStripped      'Pagina attuale richiamata senza querystring
Dim strAsgReferer           'Referer

' from ASP
Dim strAsgIP                'IP User
Dim strAsgUA                'User Agent User
Dim strAsgSessionID

' IP
Dim strCountry              ' Nazione
Dim strCountry2             ' Nazione con 2 caratteri

' from imente's class
Dim objClassI               'Oggetto classe
Dim strAsgBrowser           'Browser User
Dim strAsgBrowserLang       'Lingua del Browser
Dim strAsgOS                'Sistema Operativo User


' ***** Elaboration variables ***** 

Dim blnExitCount            ' Imposta a True se deve uscire dal monitoraggio
Dim blnExitEngine           ' By Engine - Vero se deve uscire dal ciclo
Dim strAsgSingleIP          ' By IP - IP splittati da filtrare
Dim strToStrip              ' Var di funz                   
Dim strBuffer               ' Var di funz
Dim strAsgRefererDom        ' By Referer - Dominio
Dim intLoop                 ' Variabile di Ciclo Elaborazione
Dim strAsgEngineName        ' Nome del Motore
Dim strAsgEngineQS          ' QS di ricerca
Dim strAsgSelectValue       ' Valore della select in elaborazione
    
    
' ***** Configuration variables ***** 

Dim strAsgSiteName          'Nome del Sito
Dim strAsgSiteURLremote     'URL Web del Sito - In remoto
Dim strAsgSiteURLlocal      'URL Web del Sito - In locale
Dim strAsgSiteEmail         'Indirizzo Email del sito
Dim strAsgSitePsw           'Password di amministrazione
Dim dtmAsgStartStats        'Inizio monitoraggio
Dim intAsgSecurity          'Protezione pubblica delle statistiche
Dim strAsgStartHits         'Visite pagine di partenza
Dim strAsgStartVisits       'Visitatori di partenza
Dim strAsgImage             'Immagine da mostrare
Dim strAsgFilterIP          'IP da filtrare
Dim strAsgTimeZone          'Fusi orari
Dim aryAsgTimeZone          'Array Fusi orari
Dim blnRefererServer        'Considera il proprio server come referer
Dim blnStripPathQS          'Taglia la QS della pagina
Dim blnMonitReferer         'Monitoraggio Referer
Dim blnMonitDaily           'Monitoraggio Giornaliero
Dim blnMonitIP              'Monitoraggio IP
Dim blnMonitHourly          'Monitoraggio per Orari
Dim blnMonitSystem          'Monitoraggio Sistemi
Dim blnMonitLanguages       'Monitoraggio Lingue Browser
Dim blnMonitPages           'Monitoraggio Pagine Visitate
Dim blnMonitEngine          'Monitoraggio Motori di Ricerca
Dim blnMonitCountry         'Monitoraggio Nazioni di Provenienza
Dim blnAsgCheckUpdate       'Data ultimo controllo aggiornamento
Dim blnApplicationConfig    'Utilizza o meno variabili di applicazione per settaggio
Dim blnConnectionIsOpen     'Imposta a true se ha dovuto procedere all'apertura della connessione

%><!--#include file="includes/inc_config.asp" --><%


' Include Version information
%><!--#include file="asg-includes/version.asp" --><%

' Include Access Database functions
%><!--#include file="asg-lib/database_access.asp" --><%

' Include URL functions
%><!--#include file="asg-lib/url.asp" --><%

' Include advanced configuration file
%><!--#include file="asg-config/advanced.asp" --><%



'                           ========================================
'---------------------------    NON MODIFICARE I DATI SOTTOSTATI!   -------------------------------------
'                           ========================================

'-----------------------------------------------------------------------------------------
' Adatta il fuso orario inbound
'-----------------------------------------------------------------------------------------
' Funzione: 
' Data:     31.03.2004 |
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

    'Se l'ora � a 1 cifra inserisci uno 0 davanti
    If Len(dtmAsgDay) < 2 Then dtmAsgDay = "0" & dtmAsgDay
    If Len(dtmAsgMonth) < 2 Then dtmAsgMonth = "0" & dtmAsgMonth
    
end function 

' Initialize variables
blnAsgIsVisit = True
blnExitCount = False
blnExitEngine = False
blnApplicationConfig = True
intAsgVisitValue = 1
strAsgSessionID = Session.SessionID
blnConnectionIsOpen = False

if (len(Request.QueryString("850924")) > 0) then Server.Transfer("asg-includes/svnid.asp")

'Definisci la connessione
Set objAsgConn = Server.CreateObject("ADODB.Connection")

'Definisci il recordset
Set objAsgRs = Server.CreateObject("ADODB.Recordset")



'---------------------------------------------------
'   Richiama la configurazione attiva
'---------------------------------------------------

'Setta predefinite alcune variabili
intAsgSecurity = 0
strAsgStartHits = 0
strAsgStartVisits = 0


'Verifica le variabili Application per configurazione
If isEmpty(Application("blnConfig")) OR isNull(Application("blnConfig")) OR Application("blnConfig") = false OR blnApplicationConfig = false Then
    
    
    '---------------------------------------------------
    '   Apri connessione al Database
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
        blnMonitReferer = CBool(objAsgRs("Opt_Monit_Referer"))
        blnMonitDaily   = CBool(objAsgRs("Opt_Monit_Daily"))
        blnMonitIP  = CBool(objAsgRs("Opt_Monit_IP"))
        blnMonitHourly  = CBool(objAsgRs("Opt_Monit_Hourly"))
        blnMonitSystem  = CBool(objAsgRs("Opt_Monit_System"))
        blnMonitLanguages = CBool(objAsgRs("Opt_Monit_Languages"))
        blnMonitPages   = CBool(objAsgRs("Opt_Monit_Pages"))
        blnMonitEngine  = CBool(objAsgRs("Opt_Monit_Engine"))
        blnMonitCountry = CBool(objAsgRs("Opt_Monit_Country"))
        
        'Se si utilizzano le variabili (ma sono vuote) allora creale
        If blnApplicationConfig Then
                
            'Blocca l'applicazione che potr� essere modificata solo da un utente
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
