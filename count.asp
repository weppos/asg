<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config_common.asp" -->
<!--#include file="wbstat/wbstat3_class.asp"-->
<!--#include file="lib/functions_count.asp" -->
<!--#include file="lib/utils.count.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


' Close database connection
if blnAsgConnIsOpen then
	' Se non si usavano variabili application chiudi 
	' connessione dato che è stata aperta per 
	' richiamare i dati.
	objAsgConn.Close
end If
	
	
sub Log()	

	'---------------------------------------------------
	' Tracking variables
	'---------------------------------------------------
	
	' From Js
	Dim strAsgClientReso			' Holds the Video reso
	Dim strAsgClientColor			' Holds the Video color
	Dim strAsgReferer				' Holds the Referer
	Dim strAsgPage					' Holds the current page (Path + Qs)
	Dim strAsgPagePath				' Holds the page path
	Dim strAsgPageQs				' Holds the page querystring
	Dim strAsgPageTitle				' Holds the page title
	Dim strAsgFontSmoothing			'						-- NON IMPLEMENTATO ATTUALMENTE --
	Dim strAsgBrowserJavaEnabled	' Java Abilitati		-- NON IMPLEMENTATO ATTUALMENTE --
	
	' From ASP
	Dim strAsgClientIP				' Holds the IP address
	Dim strAsgBrowserUA				' Holds the client User Agent
	
	' From class
	Dim objClassI					' Class object
	Dim strAsgClientOS				' Holds the OS name
	Dim strAsgBrowser				' Holds the browser name
	Dim strAsgBrowserLang			' Holds the browser language
	Dim strAsgBrowserLangCode		' Holds the browser language code
	' Dim strAsgBrowserActCookie	'						-- NON IMPLEMENTATO ATTUALMENTE --
	
	' From IP
	Dim strCountry					' Holds the country name
	Dim strCountry2					' Holds the 2chr country name
	Dim aryCountry
	
	
	'---------------------------------------------------
	' Elaboration variables
	'---------------------------------------------------

	Dim strAsgRefererDom			' Hold the referer domain
	Dim intAsgRefererType			' Holds the referer type
	Dim strAsgEngineName			' Holds the search engine name
	Dim strAsgEngineTLD				' Holds the search engine top level domain
	Dim strAsgEngineQuery			' Holds the search engine query
	Dim intAsgEnginePage			' Holds the number of the serp
	Dim ii								' Variabile di Ciclo Elaborazione
	

	' On Error Resume Next
	

	Dim blnAsgIsVisit		' Set true if it's an unique visitor
	blnAsgIsVisit = true

	Dim intAsgVisitValue	' Holds the counter value for the unique visitor
	intAsgVisitValue = 1

	Dim strAsgSessionID		' Holds the server session ID
	strAsgSessionID = Session.SessionID
	

	'-----------------------------------------------------------------------------------------
	' Get client details
	'-----------------------------------------------------------------------------------------
	
	' From Js
	strAsgClientReso = Request("w") & "x" & Request("h")	
	strAsgClientColor = Request("c")
	strAsgPage = Request("u")
	strAsgPageTitle = Request("t")
	strAsgReferer = Request("r")	
	strAsgFontSmoothing = Request("fs")
	strAsgBrowserJavaEnabled = Request("j")

	' From ASP
	strAsgClientIP = Request.ServerVariables("REMOTE_ADDR")
	strAsgBrowserUA = Request.ServerVariables("HTTP_USER_AGENT")
	

	'-----------------------------------------------------------------------------------------
	' Check if the current PC is excluded from tracking
	'-----------------------------------------------------------------------------------------

	Dim blnExitCount			' Set true to exit

	' Cookie check
	blnExitCount = exitCountByCookie()
	if blnExitCount then exit sub

	' IP check
	if Len(appAsgFilteredIPs) > 0 then 
		blnExitCount = exitCountByIP(strAsgClientIP)
		if blnExitCount then exit sub
	end if


	'-----------------------------------------------------------------------------------------
	' Client details
	'-----------------------------------------------------------------------------------------
	
	' Check resolution value
	strAsgClientReso = checkValue(strAsgClientReso, "x", ASG_UNKNOWN)
	
	' Check color deep value
	strAsgClientColor = checkValue(strAsgClientColor, "", ASG_UNKNOWN)
	
	' Check page value
	strAsgPage = Replace(strAsgPage, "%3F", "?")
	strAsgPage = checkValue(strAsgPage, "", ASG_UNKNOWN)
	strAsgPage = trimValue(strAsgPage, 240)

	' Check trailing slash
	if Len(strAsgPage) > 0 then
		strAsgPage = compareURLsForFinalSlash(strAsgPage, getURLdomain(appAsgSiteURL, false))
	end if
	
	' Strip page querystring	// Disabled since version 3.0
	' if blnStripPathQS then 
	'	strAsgPage = stripURLquerystring(strAsgPage)
	' end if
		
	' Anti-Aliasing Fonts
	'If strAsgFontSmoothing = "true" then
	'	strAsgFontSmoothing = "True"
	'Else
	'	strAsgFontSmoothing = "False"
	'End if
	
	' Check referer value
	strAsgReferer = checkValue(strAsgReferer, "", Request.ServerVariables("HTTP_REFERER"))
	' Some Antivirus Programs and Browsers if it's a directs request put the page URL
	' in the Referer value
	' if strAsgReferer = strAsgPage then strAsgReferer = ""
	' If the referer is empty kept it as empty
	' strAsgReferer = checkValue(strAsgReferer, "", ASG_UNKNOWN)
	
	' If the page is an internal web page and the option is disabled remove the referer
	if appAsgRefererServer = false then
		if InStr(stripURLquerystring(strAsgReferer), appAsgSiteURL) then
			strAsgReferer = ASG_OWNSERVER
		end if
	end if
			
	' Get the referer domain
	If strAsgReferer = ASG_UNKNOWN OR Len(strAsgReferer) < 1 Then
		strAsgRefererDom = ASG_UNKNOWN
	Else
		strAsgRefererDom = getURLdomain(strAsgReferer, false)
	End If
	
	' If the referer has a proper value check the final /
	if strAsgReferer <> ASG_UNKNOWN AND strAsgReferer <> ASG_OWNSERVER AND Len(strAsgReferer) > 0 then
		strAsgReferer = compareURLsForFinalSlash(strAsgReferer, strAsgRefererDom)
	end if


	'-----------------------------------------------------------------------------------------
	' Search engines tracking elaboration
	'-----------------------------------------------------------------------------------------
		
		if appAsgTrackEngine AND strAsgReferer <> ASG_UNKNOWN then
			
			%><!--#include file="write_permission/def/def_search_engines.asp" --><%			
			
			Dim blnAsgIsEngine
			Const STR_ASG_QUERY_PATTERN = "[\?.*&|\?]$var1$([^&|\/]+)"
			Const STR_ASG_PAGE_PATTERN = "[\?.*&|\?]$var1$(\d+)"
			
			blnAsgIsEngine = false
			
			for ii = 1 to UBound(aryAsgEngine, 2)
				
				blnAsgIsEngine = regexpTest(aryAsgEngine(1, ii), "http://" & strAsgRefererDom & "/")
				
				strAsgEngineTLD = regexpExecuteEngine(aryAsgEngine(1, ii), "http://" & strAsgRefererDom & "/")
				if Right(strAsgEngineTLD, 1) = "." then strAsgEngineTLD = Left(strAsgEngineTLD, Len(strAsgEngineTLD) - 1)
				
				' The referer is a search engine
				if blnAsgIsEngine then
					
					' Get the search engine name
					strAsgEngineName = aryAsgEngine(2, ii)
					
					' 
					if not IsNull(aryAsgEngine(3, ii)) then
						' Get the query
						strAsgEngineQuery = regexpExecuteEngine(Replace(STR_ASG_QUERY_PATTERN, "$var1$", aryAsgEngine(3, ii)), strAsgReferer)
						' Decode query
						strAsgEngineQuery = URLDecode(strAsgEngineQuery, true)
						' Clean query
						strAsgEngineQuery = filterSQLinput(strAsgEngineQuery, true, true)
					else
						strAsgEngineQuery = ""
					end if
					
					' 
					if not IsNull(aryAsgEngine(4, ii)) then
						' Get the raw serp value
						intAsgEnginePage = regexpExecuteEngine(Replace(STR_ASG_PAGE_PATTERN, "$var1$", aryAsgEngine(4, ii)), strAsgReferer)
						' if the lenght is longer than 1 then filter the value
						' and keep the returned value
						if Len(intAsgEnginePage) > 0 then
							' Filter the serp value
							intAsgEnginePage = getSearchResultPage(aryAsgEngine(5, ii), intAsgEnginePage)
						' no page info, then it could be one of this 2 cases:
						' - no page info so we are on page number 1
						' - no numerical page info or error
						' Follow the first case.
						else
							intAsgEnginePage = 1
						end if
					else
						intAsgEnginePage = -1
					end if
					
					blnAsgIsEngine = true
					
				' regexp test
				end if
										
				' Exit for if this is the right search engine
				if blnAsgIsEngine then exit for
			
			' engine loop	
			next
		
		end if

	
	' Get the referer type
	intAsgRefererType = getRefererType(strAsgReferer, strAsgRefererDom, strAsgEngineName)

	'-----------------------------------------------------------------------------------------
	' Get last user details
	'-----------------------------------------------------------------------------------------
		
	' Class Object
	Set objClassI = CreateWBstat("wbstat/wbstat3_spec/", false, ASG_UNKNOWN, 1, 0, True, False, False, True, True, True, True, True, False, True , True, True, True, False, False)
		
	strAsgBrowser = objClassI("Browser")
	strAsgBrowserLang = objClassI("Browser.Language")
	strAsgBrowserLangCode = objClassI("Browser.Language.Code")
	strAsgClientOS = objClassI("OS")
	' strAsgBrowserActCookie = objClassI("Browser.Act.Cookie")

	' Release Object
	Set objClassI = Nothing
	
	' Filter language code
	strAsgBrowserLangCode = Left(strAsgBrowserLangCode, 5)
	strAsgBrowserLangCode = Lcase(strAsgBrowserLangCode)
	strAsgBrowserLangCode = Trim(strAsgBrowserLangCode)
	strAsgBrowserLangCode = Replace(strAsgBrowserLangCode, ",", "-")
		
	'-----------------------------------------------------------------------------------------
	' Filter input for malicious SQL code
	'-----------------------------------------------------------------------------------------
	
	strAsgClientOS = filterSQLinput(strAsgClientOS, true, true)
	strAsgClientReso = filterSQLinput(strAsgClientReso, true, true)
	strAsgClientColor = filterSQLinput(strAsgClientColor, true, true)
	strAsgBrowser = filterSQLinput(strAsgBrowser, true, true)
	strAsgBrowserLang = filterSQLinput(strAsgBrowserLang, true, true)
	strAsgReferer = filterSQLinput(strAsgReferer, false, true)
	strAsgPage = filterSQLinput(strAsgPage, false, true)
	strAsgPageTitle = filterSQLinput(strAsgPageTitle, true, true)

		
	'-----------------------------------------------------------------------------------------
	' Trim long values
	'-----------------------------------------------------------------------------------------

	strAsgReferer = trimValue(strAsgReferer, 240)
	strAsgPageTitle = trimValue(strAsgPageTitle, 100)
		
	'-----------------------------------------------------------------------------------------
	' Open database connection
	'-----------------------------------------------------------------------------------------
	
	' Open database connection
	objAsgConn.Open strAsgConn

		Dim strAsgSQLtmp
		Dim strAsgSQLtmpid	' Holds the temp ID of the record to update
		Dim lngAsgUserID

		
	'-----------------------------------------------------------------------------------------
	' Check if the current hits is a unique visitor
	'-----------------------------------------------------------------------------------------

		'
		if ASG_USE_MYSQL then
			strAsgSQL = "SELECT user_id, visitor_id, user_useragent, user_last_access, user_country_2chr " &_
				"FROM " & ASG_TABLE_PREFIX & "user " &_
				"WHERE user_ip = '" & strAsgClientIP & "' AND user_useragent = '" & strAsgBrowserUA & "' " &_
				"ORDER BY user_last_access DESC " &_
				"LIMIT 1"
		else
			strAsgSQL = "SELECT TOP 1 user_id, visitor_id, user_useragent, user_last_access, user_country_2chr " &_
				"FROM " & ASG_TABLE_PREFIX & "user " &_
				"WHERE user_ip = '" & strAsgClientIP & "' AND user_useragent = '" & strAsgBrowserUA & "' "
		end if
		lngAsgUserID = -1
		'-----------------------------------------------------------------------------------------

		' strAsgSQL = "SELECT Detail_date, Visitor_ID, User_Agent, Country2 FROM "&ASG_TABLE_PREFIX&"detail_old WHERE IP = '" & strAsgClientIP & "' AND User_Agent = '" & strAsgBrowserUA & "' ORDER BY Details_ID DESC LIMIT 1"

		objAsgRs.Open strAsgSQL, objAsgConn
		if not objAsgRs.EOF then
			' Following conditions determine a non unique visitor
			' 23.11.2003 The recordset is not empty
			' 23.11.2003 Last hit is not higher than 6 hours
			' 23.11.2003 Same session ID
			Dim dtmDiffVisitTime

			dtmDiffVisitTime = DateDiff("h", CDate(objAsgRs("user_last_access")), dtmAsgNow)
			if IsNumeric(dtmDiffVisitTime) then dtmDiffVisitTime = Clng(dtmDiffVisitTime)

			if dtmDiffVisitTime < 6 OR objAsgRs("visitor_id") = strAsgSessionID then
				' Check that the day is the same also if the difference is smaller than 6 hours
				if not Day(objAsgRs("user_last_access")) <> Day(dtmAsgNow) then
					' Not unique visitor
					blnAsgIsVisit = false
					intAsgVisitValue = 0
					' Get some shared values
					strAsgSessionID = objAsgRs("visitor_id")
					lngAsgUserID = objAsgRs("user_id")
					strCountry2 = objAsgRs("user_country_2chr")
				end if
			end If

		end If
		objAsgRs.Close
		
		if not blnAsgIsVisit then

			if ASG_USE_MYSQL then
				strAsgSQL = "SELECT user_id " &_
					"FROM " & ASG_TABLE_PREFIX & "user " &_
					"WHERE user_ip = '" & strAsgClientIP & "' AND user_useragent = '" & strAsgBrowserUA & "' " &_
					"ORDER BY user_last_access DESC " &_
					"LIMIT 1"
			else
				strAsgSQL = "SELECT TOP 1 user_id " &_
					"FROM " & ASG_TABLE_PREFIX & "user " &_
					"WHERE user_ip = '" & strAsgClientIP & "' AND user_useragent = '" & strAsgBrowserUA & "' " &_
					"ORDER BY user_last_access DESC"
			end if

			objAsgRs.Open strAsgSQL, objAsgConn
				if not objAsgRs.EOF then
					lngAsgUserID = Clng(objAsgRs("user_id"))
				else
					lngAsgUserID = -1
				end if
			objAsgRs.Close

		end if

		
	'-----------------------------------------------------------------------------------------
	' Insert into the database
	'-----------------------------------------------------------------------------------------

		
	'-----------------------------------------------------------------------------------------
	' Main counter
	'-----------------------------------------------------------------------------------------
	strAsgSQL = "SELECT counter_id " &_
		"FROM " & ASG_TABLE_PREFIX & "counter " &_
		"WHERE counter_periody = " & dtmAsgYear
		
	objAsgRs.Open strAsgSQL, objAsgConn
	if objAsgRs.EOF then 
		strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "counter (counter_periody, counter_hits, counter_visits) "
		strAsgSQL = strAsgSQL & " VALUES (" & dtmAsgYear & ", 1, " & intAsgVisitValue & ")"
	else
		strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "counter SET " &_
			"counter_hits = counter_hits + 1 , " &_
			"counter_visits = counter_visits + " & intAsgVisitValue & " " &_
			"WHERE counter_periody = " & dtmAsgYear
	end if
	objAsgRs.Close

	objAsgConn.Execute(strAsgSQL)


	'-----------------------------------------------------------------------------------------
	' Referer logging
	'-----------------------------------------------------------------------------------------
	
	if appAsgTrackReferer AND Len(strAsgReferer) > 0 then
		
		strAsgSQL = "SELECT referer_id " &_
			"FROM " & ASG_TABLE_PREFIX & "referer " &_
			"WHERE referer_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' AND referer_url = '" & strAsgReferer & "' "
			
		objAsgRs.Open strAsgSQL, objAsgConn
		if objAsgRs.EOF then 
				strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "referer (referer_url, referer_domain, referer_type, referer_last_access, referer_hits, referer_visits, referer_period) "
			if ASG_USE_MYSQL then
					strAsgSQL = strAsgSQL & "VALUES ('" & strAsgReferer & "', '" & strAsgRefererDom & "', " & intAsgRefererType & ", '" & dtmAsgNow & "', 1, " & intAsgVisitValue & ", '" & dtmAsgMonth & "-" & dtmAsgYear & "')"
				else
					strAsgSQL = strAsgSQL & "VALUES ('" & strAsgReferer & "', '" & strAsgRefererDom & "', " & intAsgRefererType & ", #" & dtmAsgNow & "#, 1, " & intAsgVisitValue & ", '" & dtmAsgMonth & "-" & dtmAsgYear & "')"
				end if
		else
			if ASG_USE_MYSQL then
					strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "referer SET " &_
						"referer_hits = referer_hits + 1 , " &_
						"referer_visits = referer_visits + " & intAsgVisitValue & " , " &_
						"referer_last_access = '" & dtmAsgNow & "' " &_
						"WHERE referer_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' AND referer_url = '" & strAsgReferer & "' "
			else
					strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "referer SET " &_
						"referer_hits = referer_hits + 1 , " &_
						"referer_visits = referer_visits + " & intAsgVisitValue & " , " &_
						"referer_last_access = #" & dtmAsgNow & "# " &_
						"WHERE referer_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' AND referer_url = '" & strAsgReferer & "' "
			end if
		end if
		objAsgRs.Close
			
		objAsgConn.Execute(strAsgSQL)
			
	end if  ' appAsgTrackReferer


	'-----------------------------------------------------------------------------------------
	' Daily
	'-----------------------------------------------------------------------------------------

	if appAsgTrackDaily then
			
		if ASG_USE_MYSQL then
			strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "daily " &_
				"WHERE daily_date = '" & dtmAsgDate & "' "
		else
			strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "daily " &_
				"WHERE daily_date = #" & dtmAsgDate & "# "
		end if

		objAsgRs.Open strAsgSQL, objAsgConn
		if objAsgRs.EOF then 
			strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "daily (daily_date, referer_type_" & intAsgRefererType & ", daily_period, daily_hits, daily_visits) "
			if ASG_USE_MYSQL then
				strAsgSQL = strAsgSQL & " VALUES ('" & dtmAsgDate & "', 1, '" & dtmAsgMonth & "-" & dtmAsgYear & "', 1, " & intAsgVisitValue & ")"
			else
				strAsgSQL = strAsgSQL & " VALUES (#" & dtmAsgDate & "#, 1, '" & dtmAsgMonth & "-" & dtmAsgYear & "', 1, " & intAsgVisitValue & ")"
			end if
		else
			if ASG_USE_MYSQL then
				strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "daily SET " &_
					"referer_type_" & intAsgRefererType & " = referer_type_" & intAsgRefererType & " + 1 , " &_
					"daily_hits = daily_hits + 1 , " &_
					"daily_visits = daily_visits + " & intAsgVisitValue & " " &_
					"WHERE daily_date = '" & dtmAsgDate & "' "
			else
				strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "daily SET " &_
					"referer_type_" & intAsgRefererType & " = referer_type_" & intAsgRefererType & " + 1 , " &_
					"daily_hits = daily_hits + 1 , " &_
					"daily_visits = daily_visits + " & intAsgVisitValue & " " &_
					"WHERE daily_date = #" & dtmAsgDate & "# "
			end if
		end if  ' EOF
		objAsgRs.Close
			
		objAsgConn.Execute(strAsgSQL)
			
	end if  ' appAsgTrackDaily


	'-----------------------------------------------------------------------------------------
	' Hourly
	'-----------------------------------------------------------------------------------------

	if appAsgTrackHourly then
		
		strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "hourly " &_
			"WHERE hourly_hour = " & Hour(dtmAsgNow) & " AND hourly_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' "
			
		objAsgRs.Open strAsgSQL, objAsgConn
		if objAsgRs.EOF then 
			strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "hourly (hourly_hour, hourly_hits, hourly_visits, hourly_period) " &_
				"VALUES (" & Hour(dtmAsgNow) & ", 1, " & intAsgVisitValue & ", '" & dtmAsgMonth & "-" & dtmAsgYear & "' )"
		else
			strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "hourly SET " &_
				"hourly_hits = hourly_hits + 1 , " &_
				"hourly_visits = hourly_visits + " & intAsgVisitValue & " " &_
				"WHERE hourly_hour = " & Hour(dtmAsgNow) & " AND hourly_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' "
		end if
		objAsgRs.Close
			
		objAsgConn.Execute(strAsgSQL)
			
	end if  ' appAsgTrackHourly


	'-----------------------------------------------------------------------------------------
	' Pages
	'-----------------------------------------------------------------------------------------

	Dim aryAsgPage

	' Split path from querystring
	aryAsgPage = Split(strAsgPage, "?", 2)
	strAsgPagePath = aryAsgPage(0)
	if Ubound(aryAsgPage) > 0 then strAsgPageQs = aryAsgPage(1)

	if appAsgTrackPages then
		
		strAsgSQL = "SELECT page_id " &_
			"FROM " & ASG_TABLE_PREFIX & "page " &_
			"WHERE page_path = '" & strAsgPagePath & "' AND page_qs = '" & strAsgPageQs & "' AND page_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' "

		objAsgRs.Open strAsgSQL, objAsgConn
		if objAsgRs.EOF then 
				strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "page (page_path, page_qs, page_title, page_hits, page_visits, page_period) "
				strAsgSQL = strAsgSQL & " VALUES ('" & strAsgPagePath & "', '" & strAsgPageQs & "', '" & strAsgPageTitle & "', 1, " & intAsgVisitValue & ", '" & dtmAsgMonth & "-" & dtmAsgYear & "' )"
		else
			strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "page SET " &_
				"page_hits = page_hits + 1 , " &_
				"page_visits = page_visits + " & intAsgVisitValue & " " &_
				"WHERE page_path = '" & strAsgPagePath & "' AND page_qs = '" & strAsgPageQs & "' AND page_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' "
		end if
		objAsgRs.Close
			
		objAsgConn.Execute(strAsgSQL)
			
	end if


	'-----------------------------------------------------------------------------------------
	' Search Engines
	'-----------------------------------------------------------------------------------------

	if appAsgTrackEngine AND strAsgReferer <> ASG_UNKNOWN then
			
		' Check values
		if "[]" & strAsgEngineName <> "[]" AND "[]" & strAsgEngineQuery <> "[]" then
			
			strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "query " &_
				"WHERE query_keyphrase = '" & strAsgEngineQuery & "' AND engine_name = '" & strAsgEngineName & "." & strAsgEngineTLD & "' AND query_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' "

			objAsgRs.Open strAsgSQL, objAsgConn
			if objAsgRs.EOF then 
				strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "query (query_keyphrase, engine_name, engine_tld, engine_lang_code, query_hits, query_visits, query_period, query_serp_page) " &_
					"VALUES ('" & strAsgEngineQuery & "', '" & strAsgEngineName & "." & strAsgEngineTLD & "', '" & strAsgEngineTLD & "', '', 1, " & intAsgVisitValue & ", '" & dtmAsgMonth & "-" & dtmAsgYear & "', " & intAsgEnginePage & " )"
			else
				strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "query SET " &_
					"query_hits = query_hits + 1 , " &_
					"query_visits = query_visits + " & intAsgVisitValue & " , " &_
					"query_serp_page = " & intAsgEnginePage & " " &_
					"WHERE query_keyphrase = '" & strAsgEngineQuery & "' AND engine_name = '" & strAsgEngineName & "." & strAsgEngineTLD & "' AND query_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' "
			end if
			objAsgRs.Close
			
			objAsgConn.Execute(strAsgSQL)
			
		end if  ' values
		
	end if  ' appAsgTrackEngine


	'-----------------------------------------------------------------------------------------
	' IP Address
	'-----------------------------------------------------------------------------------------

	if appAsgTrackIP then
		
		if ASG_USE_MYSQL then
			strAsgSQLtmp = "UPDATE " & ASG_TABLE_PREFIX & "ip SET " &_
				"ip_hits = ip_hits + 1 , " &_
				"ip_visits = ip_visits + " & intAsgVisitValue & " " &_
				"WHERE ip_address = '" & strAsgClientIP & "' "
				' ADD the following string to WHERE condition
				' to daily track IP addresses
				' AND ip_last_access = '" & dtmAsgDate & "'
		else
			strAsgSQLtmp = "UPDATE " & ASG_TABLE_PREFIX & "ip SET " &_
				"ip_hits = ip_hits + 1 , " &_
				"ip_visits = ip_visits + " & intAsgVisitValue & " " &_
				"WHERE ip_address = '" & strAsgClientIP & "' "
				' ADD the following string to WHERE condition
				' to daily track IP addresses
				' ip_last_access = #" & dtmAsgDate & "#
		end if

		' Full mode
		if blnAsgIsVisit then
			
			if ASG_USE_MYSQL then
				strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "ip " &_
					"WHERE ip_address = '" & strAsgClientIP & "' "
			else
				strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "ip " &_
					"WHERE ip_address = '" & strAsgClientIP & "' "
			end if
			objAsgRs.Open strAsgSQL, objAsgConn
			if objAsgRs.EOF then 
				strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "ip (ip_address, ip_last_access, ip_hits, ip_visits, ip_period) "
				if ASG_USE_MYSQL then
					strAsgSQL = strAsgSQL & " VALUES ('" & strAsgClientIP & "', '" & dtmAsgNow & "', 1, " & intAsgVisitValue & ", '" & dtmAsgMonth & "-" & dtmAsgYear & "')"
				else
					strAsgSQL = strAsgSQL & " VALUES ('" & strAsgClientIP & "', #" & dtmAsgNow & "#, 1, " & intAsgVisitValue & ", '" & dtmAsgMonth & "-" & dtmAsgYear & "')"
				end if
			else
				strAsgSQL = strAsgSQLtmp
			end if
			objAsgRs.Close

		' Short mode
		else	
			strAsgSQL = strAsgSQLtmp
		end if

		objAsgConn.Execute(strAsgSQL)
			
	end if  ' appAsgTrackIP


	'-----------------------------------------------------------------------------------------
	' Country
	'-----------------------------------------------------------------------------------------

	if appAsgTrackCountry then
		
		' Unique visitor, detect the country
		if blnAsgIsVisit then 
			aryCountry = getCountry(strAsgClientIP)
			if aryCountry(0) = false then
				strCountry = ASG_UNKNOWN
				strCountry2 = "xx"
			else
				strCountry = aryCountry(0)
				strCountry2 = aryCountry(1)
			end if
		end if

		strAsgSQLtmp = "UPDATE " & ASG_TABLE_PREFIX & "country SET " &_
			"country_hits = country_hits + 1 , " &_
			"country_visits = country_visits + " & intAsgVisitValue & " " &_
			"WHERE country_name = '" & strCountry & "' AND country_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' "

		' Full mode
		if blnAsgIsVisit then
			
			strAsgSQL = "SELECT country_id " &_
				"FROM " & ASG_TABLE_PREFIX & "country " &_
				"WHERE country_name = '" & strCountry & "' AND country_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "'"
			objAsgRs.Open strAsgSQL, objAsgConn
				if objAsgRs.EOF then 
					strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "country (country_name, country_2chr, country_hits, country_visits, country_period) "
					strAsgSQL = strAsgSQL & "VALUES ('" & strCountry & "', '" & strCountry2 & "', 1, " & intAsgVisitValue & ", '" & dtmAsgMonth & "-" & dtmAsgYear & "' )"
				else
					strAsgSQL = strAsgSQLtmp
				end if
			objAsgRs.Close

		' Short mode
		else	
			strAsgSQL = strAsgSQLtmp
		end if

		objAsgConn.Execute(strAsgSQL)
			
	end if  ' appAsgTrackCountry


	'-----------------------------------------------------------------------------------------
	' System
	'-----------------------------------------------------------------------------------------

	if appAsgTrackSystem then
		
		strAsgSQLtmp = "UPDATE " & ASG_TABLE_PREFIX & "system SET " &_
			"system_hits = system_hits + 1 , " &_
			"system_visits = system_visits + " & intAsgVisitValue & " " &_
			"WHERE system_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' AND system_os = '" & strAsgClientOS & "' AND system_browser = '" & strAsgBrowser & "' AND system_reso = '" & strAsgClientReso & "' AND system_color = '" & strAsgClientColor & "' "

		' Full mode
		if blnAsgIsVisit then
			
			strAsgSQL = "SELECT * FROM " & ASG_TABLE_PREFIX & "system " &_
				"WHERE system_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' AND system_os = '" & strAsgClientOS & "' AND system_browser = '" & strAsgBrowser & "' AND system_reso = '" & strAsgClientReso & "' AND system_color = '" & strAsgClientColor & "' "

			objAsgRs.Open strAsgSQL, objAsgConn
			if objAsgRs.EOF then 
				strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "system (system_os, system_browser, system_reso, system_color, system_hits, system_visits, system_period) "
				strAsgSQL = strAsgSQL & "VALUES ('" & strAsgClientOS & "', '" & strAsgBrowser & "', '" & strAsgClientReso & "', '" & strAsgClientColor & "', 1, " & intAsgVisitValue & ", '" & dtmAsgMonth & "-" & dtmAsgYear & "' )"
			else
				strAsgSQL = strAsgSQLtmp
			end if
			objAsgRs.Close

		' Short mode
		else	
			strAsgSQL = strAsgSQLtmp
		end if

		objAsgConn.Execute(strAsgSQL)
			
	end if  ' appAsgTrackSystem


	'-----------------------------------------------------------------------------------------
	' Browser language
	'-----------------------------------------------------------------------------------------

	if appAsgTrackLang then

		strAsgSQLtmp = "UPDATE " & ASG_TABLE_PREFIX & "language SET " &_
			"lang_hits = lang_hits + 1 , " &_
			"lang_visits = lang_visits + " & intAsgVisitValue & " " &_
			"WHERE lang_code = '" & strAsgBrowserLangCode & "' AND lang_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' "

		' Full mode
		if blnAsgIsVisit then
			
			strAsgSQL = "SELECT lang_id " &_
				"FROM " & ASG_TABLE_PREFIX & "language " &_
				"WHERE lang_code = '" & strAsgBrowserLangCode & "' AND lang_period = '" & dtmAsgMonth & "-" & dtmAsgYear & "' "
				objAsgRs.Open strAsgSQL, objAsgConn
			if objAsgRs.EOF then 
					strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "language (lang_name, lang_code, lang_code_main, lang_hits, lang_visits, lang_period) "
				strAsgSQL = strAsgSQL & "VALUES ('" & strAsgBrowserLang & "', '" & strAsgBrowserLangCode & "', '" & Left(strAsgBrowserLangCode, 2) & "', 1, " & intAsgVisitValue & ", '" & dtmAsgMonth & "-" & dtmAsgYear & "' )"
			else
				strAsgSQL = strAsgSQLtmp
			end if
			objAsgRs.Close

		' Short mode
		else	
			strAsgSQL = strAsgSQLtmp
		end if

		objAsgConn.Execute(strAsgSQL)
			
	end if  ' appAsgTrackLang


		'-----------------------------------------------------------------------------------------
		' Unique user tracking
		'-----------------------------------------------------------------------------------------
		if blnAsgIsVisit then

			'SQL string to track users
			strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "user "
			strAsgSQL = strAsgSQL & "(visitor_id, user_first_access, user_last_access, user_hits, user_ip, user_country, user_country_2chr, user_useragent, user_os, user_browser, user_browser_lang, user_browser_lang_code, user_reso, user_color, user_referer_url, user_last_page, user_search_query, user_search_engine, user_cached, user_cache) "
			if ASG_USE_MYSQL then
				strAsgSQL = strAsgSQL & " VALUES ('" & strAsgSessionID & "', '" & dtmAsgNow & "', '" & dtmAsgNow & "', 1, '" & strAsgClientIP & "', '" & strCountry & "', '" & strCountry2 & "', '" & strAsgBrowserUA & "', '" & strAsgClientOS & "', '" & strAsgBrowser & "', '" & strAsgBrowserLang & "', '" & strAsgBrowserLangCode & "', '" & strAsgClientReso & "', '" & strAsgClientColor & "', '" & strAsgReferer & "', '" & strAsgPage & "', '" & strAsgEngineQuery & "', '" & strAsgEngineName & "', 1, 1) "
			else
				strAsgSQL = strAsgSQL & " VALUES ('" & strAsgSessionID & "', #" & dtmAsgNow & "#, #" & dtmAsgNow & "#, 1, '" & strAsgClientIP & "', '" & strCountry & "', '" & strCountry2 & "', '" & strAsgBrowserUA & "', '" & strAsgClientOS & "', '" & strAsgBrowser & "', '" & strAsgBrowserLang & "', '" & strAsgBrowserLangCode & "', '" & strAsgClientReso & "', '" & strAsgClientColor & "', '" & strAsgReferer & "', '" & strAsgPage & "', '" & strAsgEngineQuery & "', '" & strAsgEngineName & "', 1, 1) "
			end if
	
			'Execute the query
			objAsgConn.Execute(strAsgSQL)
	
'			if ASG_USE_MYSQL then
				' Great job MySQL staff! http://dev.mysql.com/doc/mysql/en/mysql_insert_id.html
'			else
				strAsgSQL = "SELECT @@identity FROM " & ASG_TABLE_PREFIX & "user "
				objAsgRs.Open strAsgSQL, objAsgConn
					lngAsgUserID = objAsgRs(0)
				objAsgRs.Close
'			end if

		else

			' SQL string to update user information
			strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "user SET "
			if ASG_USE_MYSQL then
				strAsgSQL = strAsgSQL &	"user_last_access = '" & dtmAsgNow & "' , "
			else
				strAsgSQL = strAsgSQL &	"user_last_access = #" & dtmAsgNow & "# , "
			end if
			strAsgSQL = strAsgSQL &_
				"user_last_page = '" & strAsgPage & "' , " &_
				"user_cached = user_cached + 1 , " &_
				"user_hits = user_hits + 1 " &_
				"WHERE user_id = " & lngAsgUserID
	
			' Execute the query
			objAsgConn.Execute(strAsgSQL)
	
		end if

	'-----------------------------------------------------------------------------------------
	' Light details
	'-----------------------------------------------------------------------------------------

	strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "detail (detail_user_id, detail_date, detail_referer_url, detail_page_url, detail_cache) "
	if ASG_USE_MYSQL then
		strAsgSQL = strAsgSQL & "VALUES (" & lngAsgUserID & ", '" & dtmAsgNow & "' , '" & strAsgReferer & "', '" & strAsgPage & "', 1) "
	else
		strAsgSQL = strAsgSQL & "VALUES (" & lngAsgUserID & ", #" & dtmAsgNow & "# , '" & strAsgReferer & "', '" & strAsgPage & "', 1) "
	end if

	objAsgConn.Execute(strAsgSQL)


	'-----------------------------------------------------------------------------------------
	' Error debug
	'-----------------------------------------------------------------------------------------

	if Request.QueryString("errnumber") = 1 then
		Response.Write("<br />Error number : " & Err.Number)
	end if
		
	if Err.Number <> 0 Then

		strAsgSQL = "INSERT INTO " & ASG_TABLE_PREFIX & "debug_error "
		strAsgSQL = strAsgSQL & "(error_date, error_number, error_description, error_source) "
		if ASG_USE_MYSQL then
			strAsgSQL = strAsgSQL & "VALUES ('" & dtmAsgNow & "' , '" & filterSQLinput(Err.Number, true, true) & "', '" & filterSQLinput(Err.Description) & "', '" & filterSQLinput(Err.Source, true, true) & "') "
		else
			strAsgSQL = strAsgSQL & "VALUES (#" & dtmAsgNow & "# , '" & filterSQLinput(Err.Number) & "', '" & filterSQLinput(Err.Description, true, true) & "', '" & filterSQLinput(Err.Source, true, true) & "') "
		end if
	
		objAsgConn.Execute(strAsgSQL)
		'Err.Clear
		
	end if
		

	objAsgConn.Close
	Set objAsgConn = Nothing

end sub
   
' Execute tracking
Call Log()

' Redirect to the image URL
Response.Redirect(appAsgImage)

if Request.QueryString("elabtime") = 1 then
	Response.Write("<br />Elab time : " & FormatNumber(Timer() - startAsgElab, 4))
end if
	
%>

