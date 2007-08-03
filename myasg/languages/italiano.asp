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


'-----------------------------------------------------------------------------------------
' Definition - Do not translate!
'-----------------------------------------------------------------------------------------
Const infoAsgTypeLanguage = "italiano"
'-----------------------------------------------------------------------------------------


'-----------------------------------------------------------------------------------------
' Generali
'-----------------------------------------------------------------------------------------
Const strAsgTxtOrderBy = "Ordina per"
Const strAsgTxtURL = "Indirizzo"
Const strAsgTxtHits = "Pagine Visitate"
Const strAsgTxtVisits = "Accessi Unici"
Const strAsgTxtSmHits = "Visite"
Const strAsgTxtSmVisits = "Accessi"
Const strAsgTxtByMonth = "Divisi per Mese"
Const strAsgTxtAll = "Tutti"
Const strAsgTxtPage = "Pagina"
Const strAsgTxtOf = "di"
Const strAsgTxtAsc = "Ascendente"
Const strAsgTxtDesc = "Discendente"
Const strAsgTxtLastAccess = "Ultimo Accesso"
Const strAsgTxtShow = "Mostra"
Const strAsgTxtNoRecordInDatabase = "Nessun valore presente nel database."
Const strAsgTxtGraph = "Grafico"
Const strAsgTxtStats = "Statistiche"
Const strAsgTxtOptions = "Opzioni"
Const strAsgTxtGeneral = "Generale"
Const strAsgTxtProvenance = "Provenienza"
Const strAsgTxtExtra = "Varie"
Const strAsgTxtByHits = "per Pagine Visitate"
Const strAsgTxtByVisits = "per Accessi Unici"
Const strAsgTxtNum = "Num"

'-----------------------------------------------------------------------------------------
' statistiche.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtIndexReport = "Sommario"

'-----------------------------------------------------------------------------------------
' visitors.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtDate = "Data"
Const strAsgTxtGoToPage = "Vai alla Pagina"
Const strAsgTxtVisitorsDetails = "Dettagli Visitatori"

'-----------------------------------------------------------------------------------------
' browser.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtBrowser = "Browser"

'-----------------------------------------------------------------------------------------
' os.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtOS = "Sistema Operativo"

'-----------------------------------------------------------------------------------------
' color.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtReso = "Risoluzione"
Const strAsgTxtColor = "Prof Colore"
Const strAsgTxtSmReso = "Reso"
Const strAsgTxtSmColor = "Bit"

'-----------------------------------------------------------------------------------------
' browser_lang.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtBrowserLanguages = "Lingue del Browser"

'-----------------------------------------------------------------------------------------
' system.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtSystems = "Sistemi"

'-----------------------------------------------------------------------------------------
' referer.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtReferer = "Referer"
Const strAsgTxtRefererIn = "Referer Interni"
Const strAsgTxtRefererOut = "Referer Esterni"
Const strAsgTxtRefererAll = "Tutti i Referer"
Const strAsgTxtTypology = "Tipologia"

'-----------------------------------------------------------------------------------------
' engine.asp & query.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtSearchQuery = "Query di Ricerca"
Const strAsgTxtSearchEngine = "Motori di Ricerca"
Const strAsgTxtQuery = "Query"
Const strAsgTxtEngine = "Motore"

'-----------------------------------------------------------------------------------------
' country.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtCountry = "Nazione"

'-----------------------------------------------------------------------------------------
' ip_address.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtIP = "IP"
Const strAsgTxtIPAddress = "Indirizzo IP"
Const strAsgTxtNoInformationSelectedIP = "Nessuna informazione disponibile per l'IP selezionato!"
Const strAsgTxtShowIpInformation = "Espandi le informazioni sull'IP"

'-----------------------------------------------------------------------------------------
' accessi.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtVisitsPerDay = "Accessi per Giorno"
Const strAsgTxtVisitsPerMonth = "Accessi per Mese"
Const strAsgTxtVisitsPerHour = "Accessi per Ora"

'-----------------------------------------------------------------------------------------
' settings.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtGeneralSettings = "Gestione Applicazione"
Const strAsgTxtSecuritySettings = "Controllo Sicurezza"
Const strAsgTxtResetSettings = "Esecuzione Reset"


' NEW FROM VERSION 1.2
'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------
' login.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtLogin = "Login"
Const strAsgTxtLoginCompleted = "Login eseguito con successo"
Const strAsgTxtEntryAllowed = "Accesso consentito alle statistiche"
Const strAsgTxtClickToLogout = "Clicca qui per eseguire il Logout"
Const strAsgTxtWrongPassword = "La password inserita non è corretta"
Const strAsgTxtTypePassword = "Digitare la Password"
Const strAsgTxtEntryPassword = "Password di Accesso"

'-----------------------------------------------------------------------------------------
' settings.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtSiteName = "Nome del Sito"
Const strAsgTxtSiteURLlocal = "URL del sito - LOCALE"
Const strAsgTxtSiteURLremote = "URL del sito - REMOTO"
Const strAsgTxtSiteEmail = "E-mail di riferimento per il sito"
Const strAsgTxtConfigSettings = "Impostazioni di Configurazione"
Const strAsgTxtCountSettings = "Impostazioni di Conteggio"
Const strAsgTxtMonitSettings = "Impostazioni di Monitoraggio"
Const strAsgTxtMonitOptions = "Opzioni di Monitoraggio"
Const strAsgTxtStartVisits = "Accessi Unici di partenza"
Const strAsgTxtStartHits = "Pagine Visitate di partenza"
Const strAsgTxtFilterIPaddr = "Filtro Indirizzi IP"
Const strAsgTxtEnableMonit = "Abilita monitoraggio"
Const strAsgTxtCountServerAsReferer = "Conteggia il server come un referer"
Const strAsgTxtstrAsgIPPathQS = "Elimina il contenuto della QueryString delle pagine"
Const strAsgTxtDailyMonit = "Suddivisione giornaliera"
Const strAsgTxtHourlyMonit = "Suddivisione oraria"

'-----------------------------------------------------------------------------------------
' settings_security.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtNewPassword = "Nuova Password"
Const strAsgTxtConfirmPassword = "Conferma Password"
Const strAsgTxtUpdateSuccessfullyCompleted = "Aggiornamento completato con successo!"
Const strAsgTxtStatsProtection = "Protezione Statistiche"
Const strAsgTxtStatsProtectionLevel = "Livello di protezione"
Const strAsgTxtNone = "Nessuno"
Const strAsgTxtLimited = "Limitato"
Const strAsgTxtFull = "Completo"
Const strAsgTypeOnlyToChangePassword = "Da digitare unicamente se si desidera cambiare la Password!"
Const strAsgTxtAttentionPasswordNotMatching = "ATTENZIONE: Le Password inserite non corrispondono!"


'-----------------------------------------------------------------------------------------
' settings_reset.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtErrorOccured = "Si è verificato un errore!"
Const strAsgTxtCheckTableMatching = "Controllare che il nome della tabella corrisponda nelle impostazioni."
Const strAsgTxtTableReset = "Azzera Tabella"
Const strAsgTxtDetailContent = "Contiene le informazioni generali e le statistiche degli utenti"
Const strAsgTxtSystemContent = "Contiene le informazioni sui sistemi di navigazione degli utenti"
Const strAsgTxtHourlyContent = "Contiene la suddivisione oraria delle statistiche"
Const strAsgTxtDailyContent = "Contiene la suddivisione giornaliera delle statistiche"
Const strAsgTxtIPContent = "Contiene l'elenco degli IP degli utenti"
Const strAsgTxtLanguageContent = "Contiene le lingue dei browser di navigazione degli utenti"
Const strAsgTxtRefererContent = "Contiene le informazioni sui referer diretti al sito"
Const strAsgTxtPageContent = "Contiene le pagine visitate dagli utenti"
Const strAsgTxtQueryContent = "Contiene le query ed i motori di ricerca"
Const strAsgTxtResetAllTables = "Reset di tutte le tabelle citate"


'-----------------------------------------------------------------------------------------
' settings_reset_execute.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtExecutionReport = "Report Esecuzione"
Const strAsgTxtTable = "Tabella"
Const strAsgTxtCorrectlyDeleted = "resettata correttamente!"
Const strAsgTxtDatabaseSuccessfullyCompactedOn = "Database compattato con successo su "
Const strAsgTxtDatabaseSuccessfullyRenamedTo = "Database rinominato con successo in "

Const strAsgTxtError = "Errore"
Const strAsgTxtLogout = "Logout"
Const strAsgTxtUpdate = "Aggiorna"
Const strAsgTxtAnd = "e"


'-----------------------------------------------------------------------------------------
' statistiche.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtVisitsInformations = "Informazioni Accessi"
Const strAsgTxtGeneralInformations = "Informazioni Generali"
Const strAsgTxtGeneralAverageInformations = "Medie Generali"
Const strAsgTxtYearlyInformations = "Informazioni Annuali"
Const strAsgTxtToday = "Oggi"
Const strAsgTxtYesterday = "Ieri"
Const strAsgTxtPerDay = "per Giorno"
Const strAsgTxtPerHour = "per Ora"
Const strAsgTxtLastMonth = "scorso Mese"


' NEW FROM VERSION 1.3
'-----------------------------------------------------------------------------------------

Const strAsgTxtDetails = "Dettagli"
Const strAsgTxtGoingToBeRedirected = "Stai per essere indirizzato alla pagina da cui provenivi"
Const strAsgTxtClickToRedirect = "Clicca qui se non vuoi attendere o se il browser non ti rimanda automaticamente"
Const strAsgTxtJanuary = "Gennaio"
Const strAsgTxtFebruary = "Febbraio"
Const strAsgTxtMarch = "Marzo"
Const strAsgTxtApril = "Aprile"
Const strAsgTxtMay = "Maggio"
Const strAsgTxtJune = "Giugno"
Const strAsgTxtJuly = "Luglio"
Const strAsgTxtAugust = "Agosto"
Const strAsgTxtSeptember = "Settembre"
Const strAsgTxtOctober = "Ottobre"
Const strAsgTxtNovember = "Novembre"
Const strAsgTxtDecember = "Dicembre"
Const strAsgTxtPath = "Path"
Const strAsgTxtGroupByPath = "Raggruppa per Path"
Const strAsgTxtGroupByPage = "Raggruppa per Pagina"
Const strAsgTxtGroupByEngine = "Raggruppa per Motore"
Const strAsgTxtGroupByQuery = "Raggruppa per Query"
Const strAsgTxtIPTracking = "Tracking IP"
Const strAsgTxtFor = "per"
Const strAsgTxtTime = "Ora"
Const strAsgTxtMissedDataToElab = "Alcuni dati necessari per l'elaborazione risultano mancanti!"
Const strAsgTxtCloseWindow = "Chiudi Finestra"
Const strAsgTxtView = "Visualizza"


' NEW FROM VERSION 1.4
'-----------------------------------------------------------------------------------------

Const strAsgTxtInsufficientPermission = "Non hai i permessi sufficienti per accedere alla pagina!"
Const strAsgTxtAction = "Azione"
Const strAsgTxtAddToList = "Aggiungi ad Elenco"
Const strAsgTxtResetAndAddToList = "Resetta Elenco ed Aggiungi Valore"
Const strAsgTxtMonthlyCalendar = "Calendario Mensile"

Const strAsgTxtSunday = "Domenica"


' NEW FROM VERSION 2.0
'-----------------------------------------------------------------------------------------
Const strAsgTxtStatsOfTheMonth = "Statistiche del Mese"
Const strAsgTxtStatsOfTheYear = "Statistiche dell'Anno"
Const strAsgTxtCalendar = "Calendario"
Const strAsgTxtDomain = "Dominio"
Const strAsgTxtGroupByDomain = "Raggruppa per Dominio"
Const strAsgTxtGroupByReferer = "Raggruppa per Referer"
Const strAsgTxtInformationsToExitByIpRange = "E' possibile usare il carattere * per bloccare range di indirizzi. <br />Ex. Per bloccare la range '200.200.200.0 - 255' si dovrà inserire '200.200.200.*'"
Const strAsgTxtServerinformations = "Informazioni Server"
Const strAsgTxtFullVersion = "Versione completa"
Const strAsgTxtEurope = "Europa"
Const strAsgTxtAfrica = "Africa"
Const strAsgTxtAsia = "Asia"
Const strAsgTxtAmerica = "America"
Const strAsgTxtOceania = "Oceania"
Const strAsgTxtSkinSettings = "Gestione Skin"
Const strAsgTxtProgramTools = "Strumenti Programma"
Const strAsgTxtReportUnknownIcons = "Notifica icone non riconosciute"
Const strAsgTxtSERPreports = "Controllo SERP"
Const strAsgTxtExclusionSettings = "Esclusione Conteggio"
Const strAsgTxtExitByIP = "Esclusione tramite IP"
Const strAsgTxtExitByCookie = "Esclusione tramite Cookie"
Const strAsgTxtExcludePC = "Escludi il PC dalle statistiche"
Const strAsgTxtIncludePC = "Includi il PC nelle statistiche"
Const strAsgTxtDateSettings = "Impostazioni Data"
Const strAsgTxtTimeZoneOffSet = "Differenza di fuso orario"
Const strAsgTxtOffSetClientServer = "rispetto all'ora del Server"
Const strAsgTxtOffSetServerToGMT = "tra Server ed il meridiano fondamentale di Greenwich (GMT)"
Const strAsgTxtOffSetGMTtoUser = "tra il meridiano fondamentale di Greenwich (GMT) ed il tuo Orario"
Const strAsgTxtThisPageWasGeneratedIn = "Pagina generata in"
Const strAsgTxtSeconds = "secondi"
Const strAsgTxtOn = "in"
Const strAsgTxtAt = "alle"
Const strAsgTxtServerDateTimeAre = "L'ora e la data del server sono"
Const strAsgTxtReportDateTimeAre = "L'ora e la data dei report sono"
Const strAsgTxtCountryContent = "Contiene le informazioni sulle nazioni degli utenti"
Const strAsgTxtMonth = "Mese"
Const strAsgTxtMonths = "Mesi"
Const strAsgTxtWeek = "Settimana"
Const strAsgTxtWeeks = "Settimane"
Const strAsgTxtDataReset = "Reset dei dati"
Const strAsgTxtOlderThan = "più vecchi di"
Const strAsgTxtCurrent = "questo"
Const strAsgTxtOnlineUsers = "Utenti online"
Const strAsgTxtTop = "Top"


' NEW FROM VERSION 2.1
'-----------------------------------------------------------------------------------------
Const strAsgTxtSearch = "Cerca"
Const strAsgTxtSearchFoundNoResults = "Nessun risultato per la ricerca selezionata."
Const strAsgTxtAdvice = "Avviso"
Const strAsgTxtTablesWithWarningIconNeedsReset = "Per le tabelle segnalate da icona di pericolo è consigliabile una cancellazione di dati."
Const strAsgTxtRecords = "Record"
Const strAsgTxtFrom = "da"
Const strAsgTxtTo = "a"
Const strAsgCompactDatabase = "Compatta Database"
Const strAsgTxtIISversion = "Versione IIS"
Const strAsgTxtProtocolVersion = "Versione protocollo"
Const strAsgTxtYourIpIs = "Il tuo IP"
Const strAsgTxtServerName = "Nome Server"
Const strAsgTxtThisPCisActually = "Questo pc è attualmente"
Const strAsgTxtIncluded = "incluso"
Const strAsgTxtExcluded = "escluso"
Const strAsgTxtIntoMonitoringProcess = "nel processo di conteggio"
Const strAsgTxtUsingApplication = "Utilizzo del programma"
Const strAsgTxtPageMonitoringString = "Stringa di monitoraggio da inserire nelle pagine da tracciare"

%>