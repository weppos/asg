<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'

' Elaboration timeout
Server.ScriptTimeout = 90
Session.Timeout = 20
Response.Buffer = true

' In case of CDate Error check this value
' 1040 : Italy
Session.LCID = 1040

' Do not save in cache
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


' Database connection and recordset
Dim strAsgSQL			' SQL string
Dim objAsgConn			' Connection
Dim strAsgConn			' Database connection string
Dim objAsgRs			' Recordset
Dim blnAsgConnIsOpen	' Set to true if the connection is open
	
' Configuration variables (application)
Dim appAsgSiteName			' Holds the site name
Dim appAsgSiteURL			' Holds the site URL
Dim appAsgSitePsw			' Holds the password to login
Dim appAsgProgramSetup		' Holds the date when the program has been installed
Dim appAsgSecurity			' Holds the security level
Dim appAsgStartHits			' Holds the value of starting hits
Dim appAsgStartVisits		' Holds the value of starting visits
Dim appAsgImage				' Holds the name of the image to redirect to
Dim appAsgFilteredIPs		' Holds the comma separated list of filtered IPs
Dim appAsgTimeZone			' Holds Time Zone settings

Dim appAsgRefererServer		'Considera il proprio server come referer
Dim appAsgTrackReferer			'Monitoraggio Referer
Dim appAsgTrackDaily			'Monitoraggio Giornaliero
Dim appAsgTrackIP				'Monitoraggio IP
Dim appAsgTrackHourly			'Monitoraggio per Orari
Dim appAsgTrackSystem			'Monitoraggio Sistemi
Dim appAsgTrackLang		'Monitoraggio Lingue Browser
Dim appAsgTrackPages			'Monitoraggio Pagine Visitate
Dim appAsgTrackEngine			'Monitoraggio Motori di Ricerca
Dim appAsgTrackCountry			'Monitoraggio Nazioni di Provenienza
Dim appAsgDebugIcon			'Notifica icone non riconosciute
Dim appAsgCheckUpdate		'Data ultimo controllo aggiornamento
Dim appAsgEmailAddress		' Holds the site email address
Dim appAsgEmailServer		' Holds the outgoing SMTP mail server
Dim appAsgEmailComponent	' Holds the email component

' Date time settings
Dim dtmAsgNow				' Date Time
Dim dtmAsgDate				' Date
Dim dtmAsgYear				' Year
Dim dtmAsgMonth				' Month
Dim dtmAsgDay				' Day

' Server elaboration time
Dim startAsgElab			' Holds beginning time
startAsgElab = Timer()

' Constant variables
Const ASG_VERSION = "3.0 alpha"
Const ASG_VERSION_BUILD = "20050821"
Const ASG_VERSION_ID = 0

'-------------------------------------------------------------------------------'
' Include config settings
'-------------------------------------------------------------------------------'
%>
<!--#include file="config/config.inc.asp" -->
<!--#include file="config/config.advanced.inc.asp" -->
<!--#include file="config/database.inc.asp" -->
<%


'---------------------------------------------------
' Database connection and recordset settings
'---------------------------------------------------

' Set database connection as closed
blnAsgConnIsOpen = false

' Set database connection and recordset
Set objAsgConn = Server.CreateObject("ADODB.Connection")
set objAsgRs = Server.CreateObject("ADODB.Recordset")


' Check config
if isEmpty(Application(ASG_APPLICATION_PREFIX & "Config")) OR isNull(Application(ASG_APPLICATION_PREFIX & "Config")) OR Application(ASG_APPLICATION_PREFIX & "Config") = false OR blnApplicationConfig = false then
	
	' Open the database connection only if it's necessary.
	' If application variables are enabled then keep it closed.
	objAsgConn.Open strAsgConn
	blnAsgConnIsOpen = true
	
	' Initialise SQL string to select configuration
	if ASG_USE_MYSQL then
		strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "config LIMIT 1"
	else
		strAsgSQL = "SELECT TOP 1 * FROM " & ASG_TABLE_PREFIX & "config"
	end if
	
	' Open Rs to collect configuration variables
	objAsgRs.Open strAsgSQL, objAsgConn
	if not objAsgRs.EOF then
		
		appAsgSiteName = objAsgRs("conf_site_name")
		appAsgSiteURL = objAsgRs("conf_site_url")
		appAsgSitePsw = objAsgRs("conf_site_psw")
		appAsgProgramSetup = objAsgRs("conf_setup_date")
		if isNumeric(objAsgRs("conf_security_level")) then 
			appAsgSecurity = Cint(objAsgRs("conf_security_level"))
		else
			appAsgSecurity = 0
		end if
		appAsgImage = objAsgRs("conf_image")
		appAsgFilteredIPs = objAsgRs("conf_filtered_ips")
		appAsgTimeZone = objAsgRs("conf_time_zone")
		if isNumeric(objAsgRs("conf_start_hits")) then 
			appAsgStartHits = Clng(objAsgRs("conf_start_hits"))
		else
			appAsgStartHits	= 0
		end if
		if isNumeric(objAsgRs("conf_start_visits")) then 
			appAsgStartVisits = Clng(objAsgRs("conf_start_visits"))
		else
			appAsgStartVisits = 0
		end if
		appAsgRefererServer = CBool(objAsgRs("conf_referer_server"))
		appAsgDebugIcon = CBool(objAsgRs("conf_debug_icon"))
		appAsgCheckUpdate = CLng(objAsgRs("conf_check_update"))
		appAsgTrackReferer	= CBool(objAsgRs("track_referer"))
		appAsgTrackDaily	= CBool(objAsgRs("track_daily"))
		appAsgTrackIP	= CBool(objAsgRs("track_ip"))
		appAsgTrackHourly	= CBool(objAsgRs("track_hourly"))
		appAsgTrackSystem	= CBool(objAsgRs("track_system"))
		appAsgTrackLang = CBool(objAsgRs("track_lang"))
		appAsgTrackPages	= CBool(objAsgRs("track_page"))
		appAsgTrackEngine	= CBool(objAsgRs("track_engine"))
		appAsgTrackCountry = CBool(objAsgRs("track_country"))
		appAsgEmailAddress = objAsgRs("conf_email_address")
		appAsgEmailComponent = objAsgRs("conf_email_component")
		appAsgEmailServer = objAsgRs("conf_email_server")
		
		' if the program uses application variables give them the right value
		if blnApplicationConfig then
				
			' Lock application variables to keep them safe during the update
			Application.Lock
			
			' Read the configuration details from the database
			Application(ASG_APPLICATION_PREFIX & "site_name") = appAsgSiteName
			Application(ASG_APPLICATION_PREFIX & "site_url") = appAsgSiteURL
			Application(ASG_APPLICATION_PREFIX & "SitePsw") = appAsgSitePsw
			Application(ASG_APPLICATION_PREFIX & "ProgramSetup") = appAsgProgramSetup
			Application(ASG_APPLICATION_PREFIX & "Security") = CInt(appAsgSecurity)
			Application(ASG_APPLICATION_PREFIX & "Image") = appAsgImage
			Application(ASG_APPLICATION_PREFIX & "FilteredIPs") = appAsgFilteredIPs
			Application(ASG_APPLICATION_PREFIX & "TimeZone") = appAsgTimeZone
			Application(ASG_APPLICATION_PREFIX & "StartHits") = CLng(appAsgStartHits)
			Application(ASG_APPLICATION_PREFIX & "StartVisits") = CLng(appAsgStartVisits)
			Application(ASG_APPLICATION_PREFIX & "RefererServer") = CBool(appAsgRefererServer)
			Application(ASG_APPLICATION_PREFIX & "TrackReferer") = CBool(appAsgTrackReferer)
			Application(ASG_APPLICATION_PREFIX & "TrackDaily") = CBool(appAsgTrackDaily)
			Application(ASG_APPLICATION_PREFIX & "TrackIP") = CBool(appAsgTrackIP)
			Application(ASG_APPLICATION_PREFIX & "TrackHourly") = CBool(appAsgTrackHourly)
			Application(ASG_APPLICATION_PREFIX & "TrackSystem") = CBool(appAsgTrackSystem)
			Application(ASG_APPLICATION_PREFIX & "TrackLang") = CBool(appAsgTrackLang)
			Application(ASG_APPLICATION_PREFIX & "TrackPages") = CBool(appAsgTrackPages)
			Application(ASG_APPLICATION_PREFIX & "TrackEngine") = CBool(appAsgTrackEngine)
			Application(ASG_APPLICATION_PREFIX & "TrackCountry") = CBool(appAsgTrackCountry)
			Application(ASG_APPLICATION_PREFIX & "DebugIcon") = CBool(appAsgDebugIcon)
			Application(ASG_APPLICATION_PREFIX & "CheckUpdate") = CLng(appAsgCheckUpdate)
			Application(ASG_APPLICATION_PREFIX & "EmailAddress") = appAsgEmailAddress
			Application(ASG_APPLICATION_PREFIX & "EmailServer") = appAsgEmailServer
			Application(ASG_APPLICATION_PREFIX & "EmailComponent") = appAsgEmailComponent
		
			' Set application variables to true
			Application(ASG_APPLICATION_PREFIX & "Config") = true
				
			' Unlock the application
			Application.UnLock
		
		end if	' application variables

	end if	' not .EOF
	objAsgRs.Close
	
' Get configuration from application variables
elseif blnApplicationConfig then

	appAsgSiteName = Application(ASG_APPLICATION_PREFIX & "site_name")
	appAsgSiteURL = Application(ASG_APPLICATION_PREFIX & "site_url")
	appAsgSitePsw = Application(ASG_APPLICATION_PREFIX & "SitePsw")
	appAsgProgramSetup = Application(ASG_APPLICATION_PREFIX & "ProgramSetup")
	appAsgSecurity = Cint(Application(ASG_APPLICATION_PREFIX & "Security")) 
	appAsgImage = Application(ASG_APPLICATION_PREFIX & "Image")
	appAsgFilteredIPs = Application(ASG_APPLICATION_PREFIX & "FilteredIPs")
	appAsgTimeZone = Application(ASG_APPLICATION_PREFIX & "TimeZone")
	appAsgStartHits = Clng(Application(ASG_APPLICATION_PREFIX & "StartHits"))
	appAsgStartVisits = Clng(Application(ASG_APPLICATION_PREFIX & "StartVisits")) 
	appAsgRefererServer = CBool(Application(ASG_APPLICATION_PREFIX & "RefererServer"))
	appAsgDebugIcon = CBool(Application(ASG_APPLICATION_PREFIX & "DebugIcon"))
	appAsgCheckUpdate = Clng(Application(ASG_APPLICATION_PREFIX & "CheckUpdate"))
	appAsgTrackDaily = CBool(Application(ASG_APPLICATION_PREFIX & "TrackDaily"))
	appAsgTrackIP = CBool(Application(ASG_APPLICATION_PREFIX & "TrackIP"))
	appAsgTrackHourly = CBool(Application(ASG_APPLICATION_PREFIX & "TrackHourly"))
	appAsgTrackSystem = CBool(Application(ASG_APPLICATION_PREFIX & "TrackSystem"))
	appAsgTrackLang = CBool(Application(ASG_APPLICATION_PREFIX & "TrackLang"))
	appAsgTrackPages = CBool(Application(ASG_APPLICATION_PREFIX & "TrackPages"))
	appAsgTrackEngine = CBool(Application(ASG_APPLICATION_PREFIX & "TrackEngine"))
	appAsgTrackCountry = CBool(Application(ASG_APPLICATION_PREFIX & "TrackCountry"))
	appAsgEmailAddress = Application(ASG_APPLICATION_PREFIX & "EmailAddress")
	appAsgEmailComponent = Application(ASG_APPLICATION_PREFIX & "EmailComponent")
	appAsgEmailServer = Application(ASG_APPLICATION_PREFIX & "EmailServer")

end if	' application variables check

' Update the time with time zone settings
if not Len(appAsgTimeZone) > 0 then appAsgTimeZone = "+0"
Call formatTimeZone(Now(), appAsgTimeZone)

%>
<!--#include file="lib/functions_common.asp"-->
<!--#include file="lib/utils.common.asp"-->
<!--#include file="lib/utils.datetime.asp"-->