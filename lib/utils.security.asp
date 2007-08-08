<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'/**
' * Collect user information, check page privileges and allow or deny 
' * the user to browse the page.
' *
' * @since 		1.x
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function checkPermission(levelNone, levelLimited, levelFull, securitylevel)
	
	' Autologin user
	Dim blnAutologin
	blnAutologin = checkPermissionCookie()
	
	Dim aryAllow(2)	' Holds the security levels
	
	aryAllow(0) = CBool(levelNone)
	aryAllow(1) = CBool(levelLimited)
	aryAllow(2) = CBool(levelFull)
	
	' The selected level security deny access to
	' not logged in users
	if aryAllow(securitylevel) = false then
	
		if Session("asgLogin") <> "Logged" then

			Set objAsgRs = Nothing
			objAsgConn.Close
			Set objAsgConn = Nothing
			Response.Redirect("login.asp?backto=" & Server.URLEncode(Request.ServerVariables("URL")))
		
		end If
		
	end if

end function

'/**
' * Check autologin cookie and autologin user if the cookie contains right values or
' * reset old values if the cookie contains bad values.
' * 
' * @return 	(bool) true if the user can be logged in,
' *				false otherwise.
' * 
' * @since 		3.0
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function checkPermissionCookie()
	
	Dim strPassword		' Holds the program password
	Dim return			' Holds a tmp return value
	strPassword = Request.Cookies(ASG_COOKIE_PREFIX & "autologin")("psw")
	
	' Check if the cookie is active and the user
	' not logged in
	if Request.Cookies(ASG_COOKIE_PREFIX & "autologin")("autologin") = "true" AND Len(strPassword) > 0 AND Session("asgLogin") <> "Logged" then
		' User logged in
		if StrComp(strPassword, appAsgSitePsw) = 0 then
			Session("asgLogin") = "Logged"
			Response.Cookies(ASG_COOKIE_PREFIX & "autologin")("autologin") = "true"
			Response.Cookies(ASG_COOKIE_PREFIX & "autologin")("psw") = strPassword
			Response.Cookies(ASG_COOKIE_PREFIX & "autologin").Expires = dateAdd("yyyy", 1, date)
			return = true
		else
			Response.Cookies(ASG_COOKIE_PREFIX & "autologin")("autologin") = ""
			Response.Cookies(ASG_COOKIE_PREFIX & "autologin")("psw") = ""
			return = false
		end if
	end if
	
	checkPermissionCookie = return
	
end function

'/**
' * Check the session ID value to prevent unauthorized visitors.
' * 
' * @param 		(string) lngSessionID	- the collected session ID value.
' * @return 	(bool) true if the session ID value matches,
' *				false otherwise.
' * 
' * @since 		3.0
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function checkSessionID(lngSessionID)

	' Check to see if the session ID collected from the form
	' matches to the current one 
	if lngSessionID <> Session.SessionID then

		Set objAsgRs = Nothing
'		objAsgConn.Close
		Set objAsgConn = Nothing
		Response.Redirect("login.asp?backto=" & Server.URLEncode(Request.ServerVariables("URL")))
	
	end if

end function


%>