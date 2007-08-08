<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'/**
' * Check if the cookie to exclude the pc is set.
' * 
' * @return 	bool § if the cookie exists, 
' *				false otherwise.
' *
' * @since 		2.x
' * @version	1.1
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function exitCountByCookie()
	
	if Request.Cookies(ASG_COOKIE_PREFIX & "exitcount") = "excludepc" then
		exitCountByCookie = true
	else
		exitCountByCookie = false
	end If
	
end function

'/**
' * Strip the querystring from the URL.
' * 
' * @param 		string § url - the full URL to filter
' * @return 	string § the path without querystring.
' *
' * @since 		1.0
' * @version	1.0.2 , 2005-09-03
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function stripURLquerystring(url)

	Dim buffer
	Dim return
	
	buffer = InStr(url, "?")
	if buffer then 
		return = left(url, buffer - 1) 
	else 
		return = url
	end if
	
	stripURLquerystring = return
	
end function

'/**
' * Check the string and return a default value if the input is empty.
' * 
' * @param 		string § input - the raw input
' * @param 		string § emptyValue - the value that evalutates the input as empty
' * @param 		string § defaultValue	- the value to assign to an empty input
' * @return 	string § the raw input if the string is not empty,
' *				the default value otherwise.
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function checkValue(input, emptyValue, defaultValue)
	
	if input = emptyValue then 
		input = defaultValue
	end if
	
	checkValue = input

end function

'/**
' * Convert the current IP to ip2country format.
' * 
' * @param 		string § ip - the IP address
' * @return 	string § the dotted IP if the argument is a valid IP,
' *				false otherwise.
' *
' * @since 		1.0
' * @version	1.0.1 , 2005-09-03
' * @see		http://ip-to-country.webhosting.info/node/view/55
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function dottedIp(ip)
	
	Dim aryIp
	Dim strIp

	if Trim("[]" & ip) <> "[]" then
		aryIp = Split(ip, ".")
		strIp = aryIp(0) * 16777216 + aryIp(1) * 65536 + aryIp(2) * 256 + aryIp(3)
	else
		strIp = false
	end if
	
	dottedIp = strIp

end function

'/**
' * Detect country from an IP address.
' * Thanks to http://www.ip-to-country.com/ for providing country/ip database.
' * 
' * @param 		string § ip - dotted IP
' * @return 	array § an array with 2 indexes containing country information,
' *				an array with index 0 = false if the IP is invalid.
' *
' * @since 		1.0
' * @version	1.1 , 2005-09-03
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function getCountry(ip)
	
	Dim strIP2Conn
	Dim objIP2Conn
	Dim objIP2Rs
	Dim strIP2SQL
		
	Dim return(1)
	Dim strDottedIp
	strDottedIp = dottedIp(ip)

	if strDottedIp <> false then

		Set objIP2Conn = Server.CreateObject("ADODB.Connection")
		Set objIP2Rs = Server.CreateObject("ADODB.Recordset")

		if ASG_USE_MYSQL then
			if ASG_IP2C_SAMEDATABASE then
				strIP2Conn = "driver=Mysql ODBC 3.51 Driver;server=" & ASG_MYSQL_SERVER & ";uid=" & ASG_MYSQL_USER & ";pwd=" & ASG_MYSQL_PASSWORD & ";database=" & ASG_MYSQL_DATABASE
			else
				strIP2Conn = "driver=Mysql ODBC 3.51 Driver;server=" & ASG_IP2C_SERVER & ";uid=" & ASG_IP2C_USER & ";pwd=" & ASG_IP2C_PASSWORD & ";database=" & ASG_IP2C_DATABASE
			end if
		else
			strIP2Conn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strAsgMapPathIP
		end if			
				
		objIP2Conn.Open strIP2Conn

		strIP2SQL = "SELECT ip_country, ip_country_2chr " &_
			"FROM " & ASG_IP2C_TABLE & " " &_
			"WHERE ip_from <= " & strDottedIp & " AND ip_to >= " & strDottedIp & ""
				
		objIP2Rs.Open strIP2SQL, objIP2Conn
		if objIP2Rs.EOF then
			return(0) = false
		else
			return(0) = objIP2Rs("ip_country")
			return(1) = objIP2Rs("ip_country_2chr")
		end if
		objIP2Rs.Close
					
		Set objIP2Rs = Nothing
		objIP2Conn.Close
		Set objIP2Conn = Nothing
				
	else 
		return(0) = false
	end If
	
	getCountry = return
	
end function

%>