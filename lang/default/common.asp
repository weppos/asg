<%

' TXT_		- textual message
' _Error_	- error message
' _info		- information or informative message
' _desc		- service description
' _conf		- confirmation message or popup
' _tip		- tooltip message
' _warning	- warning message


' Language definition - Use ISO 2 chr country value
Const INFO_Langset_Common = "en"

' ***** PAGE CHARSET *****
' Const STR_ASG_PAGE_CHARSET = "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" />"
Const STR_ASG_PAGE_CHARSET = "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"" />"

' Months
Const TXT_January = "January"
Const TXT_February = "February"
Const TXT_March = "March"
Const TXT_April = "April"
Const TXT_May = "May"
Const TXT_June = "June"
Const TXT_July = "July"
Const TXT_August = "August"
Const TXT_September = "September"
Const TXT_October = "October"
Const TXT_November = "November"
Const TXT_December = "December"

' Dictionary
Const TXT_pageviews = "Page Views"
Const TXT_pageviews_title = "Page Views"
Const TXT_visits = "Visits"
Const TXT_visits_title = "Visits"
Const TXT_lastAccess = "Last Access"
Const TXT_traffic = "Traffic"

Const TXT_Hits = "Richieste"
'Const TXT_Hits_short = "Visite"
'Const TXT_Visits_short = "Accessi"

' Shared
Const TXT_elabtime = "This page was generated in $time$ seconds."
Const TXT_administrator = "Administrator"
Const TXT_pagetop = "Top"
Const TXT_details = "More details"
Const TXT_page = "Page"
Const TXT_orderAsc = "ascending"
Const TXT_orderDesc = "descending"
Const TXT_orderBy = "order by"
Const TXT_homepage = "Homepage"
Const TXT_homepage_website = "Site Homepage"
Const TXT_homepage_stats = "Stats Homepage"
Const TXT_graph = "Chart"

Const TXT_time_at = "at"
Const TXT_date = "Date"
Const TXT_seconds = "seconds"
Const TXT_minutes = "minutes"
Const TXT_hours = "hours"

Const TXT_button_show = "Show report"

Const TXT_navigation_schema = "Page $pageCurrent$ of $pageCount$" ' DO NOT CHANGE SCHEMA
Const TXT_gotoPage_number_schema = "Go to page number $page$"
Const TXT_gotoPage_string_schema = "Go to page $string$"
Const TXT_gotoPage = "Go to page"
Const TXT_orderBy_schema = "Order by $field$ $order$"


Const TXT_Reset = "Reset"
Const TXT_Restore = "Ripristina"
Const TXT_Expand = "Espandi"
Const TXT_Collapse = "Raccogli"
Const TXT_Pages = "Pagine"
Const TXT_Kb = "Kb"
Const TXT_Bit = "bit"
Const TXT_Records = "Record"
Const TXT_Table = "Tabella"
Const TXT_Update = "Aggiorna"
Const TXT_And = "e"
'Const TXT_ThisPageWasGeneratedIn = "Pagina generata in"
Const TXT_From = "da"
Const TXT_To = "a"
Const TXT_Week = "Settimana"
Const TXT_Weeks = "Settimane"
Const TXT_Month = "mese"
Const TXT_Months = "mesi"
Const TXT_StatsOfTheMonth = "Statistiche del Mese"
Const TXT_StatsOfTheYear = "Statistiche dell'Anno"
Const TXT_Today = "Oggi"
Const TXT_Yesterday = "Ieri"
Const TXT_URL = "Indirizzo"
Const TXT_All = "Tutti"
Const TXT_Of = "di"
Const TXT_Search = "Cerca"
Const TXT_Stats = "Statistiche"
Const TXT_Options = "Opzioni"
Const TXT_Enable = "Attiva"
Const TXT_Disable = "Disattiva"
Const TXT_Include = "Includi"
Const TXT_Exclude = "Escludi"
Const TXT_On = "in"
Const TXT_Nodata_db = "Nessun valore presente nel database."
Const TXT_Nodata_search = "Nessun risultato per la ricerca selezionata."
Const TXT_Nodata_activeusers = "Nessun utente attivo negli ultimi $var1$ minuti."
Const TXT_Search_fieldquery = "Ricerca per <strong>$var1$</strong> nel campo <strong>$var2$</strong>"
Const TXT_Search_cleanfield = "Resetta i criteri di ricerca nel campo"
Const TXT_Search_delquery = "Rimuovi i criteri di ricerca"
Const TXT_Warning = "Warning"
Const TXT_Tooltip = "Suggerimento"
Const TXT_Info = "Informazione"
Const TXT_Advice = "Avviso"


' Errors
Const TXT_error = "Error!"
Const TXT_error_defaultpage = "Are you looking for something?"

Const TXT_permission = "You haven't enough permission to enter this page."
Const TXT_Error_Occured = "Si &egrave; verificato un errore!"
Const TXT_Error_Insufperm_default = "Stai cercando di accedere ad una cartella interna di <a href=""http://www.weppos.com/asg/"">ASP Stats Generator</a>.<br />L'esplorazione automatica di questa directory non &egrave; permessa!"

' Warnings
Const TXT_Setuplock_off_warning = "Blocco installazione non attivo"

' Menu Bar title
Const MENUGROUP_Main = "Main Menu"
Const MENUGROUP_Visitors = "Visitors"
Const MENUGROUP_Navigation = "Navigational Analysis"

Const MENUGROUP_Reports = "Report"
Const MENUGROUP_Marketing = "Marketing"
Const MENUGROUP_Tools = "Strumenti"
Const MENUGROUP_Administration = "Amministrazione"

' Bar title
Const MENUSECTION_Summary = "Sommario"
Const MENUSECTION_ActiveUsers = "Utenti online"
Const MENUSECTION_Logout = "Logout"
Const MENUSECTION_Login = "Login"

Const MENUSECTION_VisitorDetails = "Visitor Details"
Const MENUSECTION_VisitorSystems = "Visitor Systems"
Const MENUSECTION_Systems = "Platforms"
Const MENUSECTION_OS = "Operating Systems"
Const MENUSECTION_Browsers = "Browsers"
Const MENUSECTION_BrowsersLang = "Browser languages"
Const MENUSECTION_ResoBit = "Resolutions and Colors"
Const MENUSECTION_Reso = "Resolutions"
Const MENUSECTION_Colors = "Color depth"
Const MENUSECTION_IpAddresses = "IP Addresses"

Const MENUSECTION_Countries = "Nazioni"
Const MENUSECTION_VisitedPages = "Pagine Visitate"
Const MENUSECTION_HourlyReports = "Riepilogo Orario"
Const MENUSECTION_DailyReports = "Riepilogo Giornaliero"
Const MENUSECTION_YearlyReports = "Riepilogo Annuale"
Const MENUSECTION_MonthlyReports = "Riepilogo Mensile"
Const MENUSECTION_MonthlyCalendar = "Calendario Mensile"
Const MENUSECTION_Referers = "Referer"
Const MENUSECTION_SearchEngines = "Motori di ricerca"
Const MENUSECTION_SearchQueries = "Query di ricerca"
Const MENUSECTION_Serp = "SERP"
Const MENUSECTION_General = "Generale"
Const MENUSECTION_Email = "Email"
Const MENUSECTION_Maintenance = "Manutenzione"
Const MENUSECTION_SetupAndUpdate = "Installazione e Aggiornamenti"
Const MENUSECTION_Config = "Configurazione"
Const MENUSECTION_EmailConfig = "Configurazione Email"
Const MENUSECTION_Setuplock = "Blocco installazione"
Const MENUSECTION_Security = "Sicurezza"
Const MENUSECTION_TrackingExclusion = "Esclusione Conteggio"
Const MENUSECTION_CompactDatabase = "Compattazione Database"
Const MENUSECTION_BatchDeleteOldData = "Eliminazione vecchi dati"
Const MENUSECTION_Customize = "Personalizza"
Const MENUSECTION_SkinSettings = "Gestione skin"
Const MENUSECTION_ServerInfo = "Informazioni server"
Const MENUSECTION_ServerVariables = "Variabili server"
Const MENUSECTION_HelpContents = "Guida"
Const MENUSECTION_OnlineFaqs = "Faq Online"	' Online FAQs
Const MENUSECTION_MakeADonation = "Invia una Donazione"
Const MENUSECTION_CheckForNewVersion = "Verifica Aggiornamenti"
Const MENUSECTION_ReportBug = "Segnala un Bug"
Const MENUSECTION_TechnicalSupportForum = "Forum di Supporto Tecnico"
Const MENUSECTION_Feedback = "Invia un Commento"
Const MENUSECTION_LicenseAgreement = "Condizioni di Licenza"
Const MENUSECTION_About = "Informazioni su"

' Label
Const BARLABEL_DataSearch = "Data Search"
Const BARLABEL_DataExport = "Data Export and Print"
Const BARLABEL_DataView = "Data View"

' Labels
Const LABEL_Navigation = "Page navigation"
Const LABEL_ViewPeriod = "Monthly view"
Const LABEL_ViewYear = "Yearly view"
Const LABEL_ViewMode = "View mode"

Const LABEL_Searchform = "Modulo di Ricerca"
Const LABEL_Exec_Report = "Report di Esecuzione"
Const LABEL_Group = "Raggruppamento"
Const LABEL_Type = "Tipologie"
Const LABEL_Settings_site = "Impostazioni del sito"
Const LABEL_Settings_datetime = "Impostazioni data e ora"
Const LABEL_Settings_tracking = "Impostazioni di monitoraggio"
Const LABEL_Settings_misc = "Impostazioni aggiuntive"
Const LABEL_Settings_emailserver = "Impostazioni server email"
Const LABEL_Monitstring = "Stringa di monitoraggio"

' compact_database.asp
Const TXT_CompactAndOptimize = "Compatta ed Ottimizza"
Const TXT_Db_compact_info = "La compattazione ed ottimizzazione delle tabelle permette il ripristino di tabelle corrotte e contribuisce a mantenere efficienti le performance e la dimensione del database."
Const TXT_Db_compact_info2 = "Si consiglia di eseguire periodicamente il comando in relazione al traffico del sito."
Const TXT_Db_mysql_optimized = "Database ottimizzato con successo"
Const TXT_Db_access_compacted = "Database compattato con successo su "
Const TXT_Db_access_renamed = "Database rinominato con successo in "
Const TXT_Db_weight = "Peso del database"
Const TXT_BeforeCompating = "prima della compattazione"
Const TXT_AfterCompating = "dopo della compattazione"

' settings_skin.asp
Const TXT_SelectColor = "Seleziona un colore"

' color_palette.asp
Const TXT_ColorPalette = "Tavolozza Colori"
Const TXT_ColorPalette_WebSafePalette = "Colori Web"
Const TXT_ColorPalette_WindowsSystemPalette = "Sistema Windows"
Const TXT_ColorPalette_GreyScalePalette = "Scala di grigi"
Const TXT_ColorPalette_MacOSPalette = "Mac OS"

' batch_delete_old_data.asp
Const TXT_Tbldescr_all = "Tutte le tabelle elencate"
Const TXT_Tbldescr_detail = "Contiene le informazioni generali e le statistiche degli utenti"
Const TXT_Tbldescr_system = "Contiene le informazioni sui sistemi di navigazione degli utenti"
Const TXT_Tbldescr_hourly = "Contiene la suddivisione oraria delle statistiche"
Const TXT_Tbldescr_daily = "Contiene la suddivisione giornaliera delle statistiche"
Const TXT_Tbldescr_ip = "Contiene l'elenco degli IP degli utenti"
Const TXT_Tbldescr_language = "Contiene le lingue dei browser di navigazione degli utenti"
Const TXT_Tbldescr_referer = "Contiene le informazioni sui referer diretti al sito"
Const TXT_Tbldescr_page = "Contiene le pagine visitate dagli utenti"
Const TXT_Tbldescr_query = "Contiene le query ed i motori di ricerca"
Const TXT_Tbldescr_country = "Contiene le informazioni sulle nazioni degli utenti"
Const TXT_Error_CheckTableMatching = "Controllare che il nome della tabella corrisponda nelle impostazioni."
Const TXT_Deldata = "Eliminazione dei dati"
Const TXT_Deldata_conf = "Confermi l\'eliminazione definitiva dei record selezionati?"
Const TXT_Deldata_all = "reset completo"
Const TXT_Deldata_completed = "I dati selezionati sono stati eliminati."
Const TXT_Deldata_OlderThan_weekC = "pi&#249; vecchi della settimana corrente"
Const TXT_Deldata_OlderThan_week1 = "pi&#249; vecchi di 1 settimana"
Const TXT_Deldata_OlderThan_weekN = "pi&#249; vecchi di $var1$ settimane"
Const TXT_Deldata_OlderThan_monthC = "pi&#249; vecchi del mese corrente"
Const TXT_Deldata_OlderThan_month1 = "pi&#249; vecchi di 1 mese"
Const TXT_Deldata_OlderThan_monthN = "pi&#249; vecchi di $var1$ mesi"


' login.asp
Const TXT_login = "Login"
Const TXT_cookiesMustBeEnabled = "Cookies must be enabled past this point."
Const TXT_password = "Password"
Const TXT_password_desc = "Type the password to enter administration area."
Const TXT_password_wrong = "Wrong password."
Const TXT_password_forgotten = "Forgotten your password?"
Const TXT_autologin = "Remember me"
Const TXT_login_completed = "Login succesfully completed."
Const TXT_login_entryAllowed = "Now you can enter and manage your traffic information."
Const TXT_login_redirectPreviousPage = "In a few seconds you'll be redirected to the previous page."
Const TXT_logout = "Logout"
Const TXT_logout_execute = "Click here to logout."
Const TXT_logout_conf = "Are you sure you wish to log out?"
Const TXT_login_clickAndGo = "Click here if your browser does not automatically redirect you."


' browser.asp
Const TXT_browser = "Browser"

' browser_lang.asp
Const TXT_browser_language = "Browser language"

' color.asp
Const TXT_reso = "Resolution"
Const TXT_reso_short = "Reso"
Const TXT_color = "Color depth"
Const TXT_color_short = "Bit"

' os.asp
Const TXT_os = "Operating System"

' country.asp
Const TXT_Country = "Nazione"

' pages.asp
Const TXT_Path = "Path"

' engine.asp & query.asp
Const TXT_search_engine = "Search Engine"

Const TXT_Query = "Query"
Const TXT_Queries = "Query"
Const TXT_Searchquery = "Query di Ricerca"
Const TXT_Searchengine = "Motore di Ricerca"
' Const TXT_search_engines = "Motori"
' Const TXT_Searchqueries = "Query di Ricerca"
' Const TXT_Searchengines = "Motore di Ricerca"

' referer.asp
Const TXT_referer = "Referer"
Const TXT_domain = "Domain"
Const TXT_referer_type1 = "Direct request"
Const TXT_referer_type2 = "Internal referer"
Const TXT_referer_type3 = "Mirror or Alternative domain"
Const TXT_referer_type4 = "External referer"
Const TXT_referer_type5 = "Search Engine"

Const TXT_referer_longer250_warning = "L'indirizzo &egrave; stato accorciato in fase di monitoraggio poich&egrave; superiore alla dimensione massima consentita di 250 caratteri. Potrebbe non essere possibile visualizzarlo nel browser accedendo dal collegamento automatico."
Const TXT_referer_debugengine_advice = "E' stata individuata una corrispondenza per la quale il referer potrebbe risultare un motore di ricerca non individuato. Per aiutare lo sviluppo e l'aggiornamento del programma &egrave; possibile verificare ed inviare una segnalazione automatica di aggiornamento a www.weppos.com ."
Const TXT_referer_tip1 = "E' possibile abilitare il <strong>debug automatico dei motori di ricerca</strong>. <br />Se attivo il programma provveder&agrave; a verificare ciascun referer esterno con un algoritmo di ricerca e tenter&agrave; di individuare possibili motori non inseriti nelle definizioni.</p><p style=""text-align: justify;"">Per attivare la funzione &egrave; sufficiente abilitare a <em>true</em> la funzione nel file inc_config_advanced.asp impostando il valore a <pre style=""text-align: center;"">Const ASG_DEBUG_SEARCHENGINES = true</pre>"

' sysinfo.asp
Const TXT_VbsEngine = "Motore VbScript"
Const TXT_Server_bados_warning = "Il programma potrebbe non funzionare correttamente con questo sistema operativo. <br />Si consiglia se possibile l'upgrade ad un sistema operativo pi&ugrave; aggiornato."

' settings_exitcount.asp
Const TXT_ExitByIP = "Esclusione tramite IP"
Const TXT_ExitByCookie = "Esclusione tramite Cookie"
Const TXT_Exclmex = "Questo pc &egrave; attualmente $v1$ nel processo di conteggio."
Const TXT_Exclpc = "$v1$ il PC dalle statistiche"
Const TXT_Included = "incluso"
Const TXT_Excluded = "escluso"
Const TXT_Filtered_IPs = "Indirizzi IP filtrati"

' settings_security.asp
Const TXT_Password_new = "Nuova Password"
Const TXT_Password_confirm = "Conferma Password"
Const TXT_Update_Completed = "Aggiornamento completato con successo!"
Const TXT_StatsProtection = "Protezione Statistiche"
Const TXT_Seclevel = "Livello di protezione"
Const TXT_Seclevel_None = "Nessuno"
Const TXT_Seclevel_Limited = "Limitato"
Const TXT_Seclevel_Full = "Completo"
Const TXT_Error_PasswordNotMatching = "Le Password inserite non corrispondono!"

' settings_common.asp
Const TXT_SiteName = "Nome del Sito"
Const TXT_SiteURL = "URL del Sito"
Const TXT_Option_refserver = "Conteggia il server come un referer"
Const TXT_Debug_icons = "Notifica icone non riconosciute"
Const TXT_Datetime_servernow = "Ora e la data del server sono"
Const TXT_Datetime_offset = "Differenza di fuso orario"
Const TXT_Datetime_offsetbetw = "rispetto all'ora del Server"

' settings_email.asp
Const TXT_Email_address = "Indirizzo email"
Const TXT_Email_server = "Server SMTP in uscita"
Const TXT_Email_object = "Componente email"
Const BUTTON_Object_test = "Test componente"

' asg_visitor.asp
Const TXT_user_system = "User platform"
Const TXT_user_agent = "User agent"
Const TXT_currentpage = "Current page"
Const TXT_visit_length = "Visit length"
Const TXT_visit_length_schema = "<strong>$hours$</strong> $hours_label$, <strong>$minutes$</strong> $minutes_label$, $seconds$ $seconds_label$" ' DO NOT CHANGE SCHEMA
Const TXT_active_range_schema = "from $startDate$ at $startTime$ to $endDate$ at $endTime$" ' DO NOT CHANGE SCHEMA





' asg_ip_address.asp
Const TXT_IP = "IP"
Const TXT_IPaddress = "IP address"

' main.asp
Const TXT_BoxTitle_TrafficSummary = "Riepilogo generale"
Const TXT_BoxTitle_TrafficSummary_Year = "Riepilogo traffico annuale"
Const TXT_BoxTitle_TrafficSummary_Average = "Riepilogo traffico medio"
Const TXT_ServerInfo = "Informazioni server"
Const TXT_LastMonth = "scorso mese"
Const TXT_ThisMonth = "questo mese"
Const TXT_BeginningOfStats = "Inizio monitoraggio il"


Const TXT_Cont_europe = "Europa"
Const TXT_Cont_africa = "Africa"
Const TXT_Cont_asia = "Asia"
Const TXT_Cont_america = "America"
Const TXT_Cont_oceania = "Oceania"




Const TXT_StartVisits = "Accessi Unici di partenza"
Const TXT_StartHits = "Pagine Visitate di partenza"

'-----------------------------------------------------------------------------------------
' asg_ip_address.asp
'-----------------------------------------------------------------------------------------
Const TXT_NoInformationSelectedIP = "Nessuna informazione disponibile per l'IP selezionato!"
Const TXT_ShowIpInformation = "Espandi le informazioni sull'IP"

'-----------------------------------------------------------------------------------------
' main.asp
'-----------------------------------------------------------------------------------------
Const TXT_PerDay = "per giorno"
Const TXT_PerHour = "per ora"
Const TXT_GroupByPath = "Raggruppa per Path"
Const TXT_GroupByPage = "Raggruppa per Pagina"
Const TXT_GroupByEngine = "Raggruppa per Motore"
Const TXT_GroupByQuery = "Raggruppa per Query"
Const TXT_IPTracking = "Tracking IP"
Const TXT_For = "per"
Const TXT_Time = "Ora"
Const TXT_MissedDataToElab = "Alcuni dati necessari per l'elaborazione risultano mancanti!"
Const TXT_CloseWindow = "Chiudi Finestra"
Const TXT_View = "Visualizza"
Const TXT_InsufficientPermission = "Non hai i permessi sufficienti per accedere alla pagina!"
Const TXT_Action = "Azione"
Const TXT_AddToList = "Aggiungi ad Elenco"
Const TXT_ResetAndAddToList = "Resetta Elenco ed Aggiungi Valore"
Const TXT_Sunday = "Domenica"
Const TXT_GroupByDomain = "Raggruppa per Dominio"
Const TXT_GroupByReferer = "Raggruppa per Referer"
Const TXT_FullVersion = "Versione completa"
Const TXT_Current = "questo"
Const TXT_OnlineUsers = "Utenti online"
Const TXT_IISversion = "Versione IIS"
Const TXT_ProtocolVersion = "Versione protocollo"
Const TXT_YourIpIs = "Il tuo IP"
Const TXT_ServerName = "Nome Server"
Const TXT_UsingApplication = "Utilizzo del programma"
Const TXT_Divided = "Dividi"
Const TXT_PerMonth = "per mese"



%>