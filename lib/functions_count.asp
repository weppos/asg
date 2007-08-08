<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'




'-----------------------------------------------------------------------------------------
' Strip the protocol from the URL returning just the second or third level domain.
' http://www.weppos.com || http://weppos.com --> weppos.com		
' http://www.weppos.com/ || http://weppos.com/ --> weppos.com		
' http://www.domain.weppos.com || http://domain.weppos.com --> domain.weppos.com		
' http://www.domain.weppos.com/ || http://domain.weppos.com/ --> domain.weppos.com		
'
' @since 1.0
' @version 1.01 , 20050220
'-----------------------------------------------------------------------------------------
public function stripURLprotocol(argUrl)

	Dim lvString 
	Dim lvTmp
	
	Dim strToStrip, strBuffer
	
	lvString = Instr(argUrl, "://")
	if lvString then lvTmp = Right(argUrl, Len(argUrl) - (3 + lvString - 1)) else lvTmp = argUrl
	' Remove the www .
	if left(lvTmp, 4) = "www." then lvTmp = Right(lvTmp, Len(lvTmp) - 4)
	' Remove the trailing / .
	if Right(lvTmp, 1) = "/" then lvTmp = Left(lvTmp, Len(lvTmp) - 1)
	
	' Return function
	stripURLprotocol = lvTmp
	
end function


'-----------------------------------------------------------------------------------------
' Ricava dominio	
'-----------------------------------------------------------------------------------------
' Function:	Ricava l'esclusivo dominio da un URL di partenza
' Date: 	01.09.03 | 01.09.03
' Comment:	Funziona anche se non è presente lo slash finale!
'			http://www.weppos.com || http://www.weppos.com/ || http://www.weppos.com/.../ --> www.weppos.com		
'			http://weppos.com || http://weppos.com/ || http://weppos.com/.../ --> weppos.com
'-----------------------------------------------------------------------------------------
function getURLdomain(url, full)

	Dim strToStrip, strBuffer, strDomain
	
	strToStrip = Instr(url, "://")
	if strToStrip then 
		strBuffer = right(url, len(url) - (3 + strToStrip - 1)) 
	else 
		strBuffer = url
	end if
	strToStrip = Instr(strBuffer, "/")
	if strToStrip > 0 Then
		strDomain = Left(strBuffer, strToStrip)
	else
		strDomain = strBuffer
	end If
	' Remove the final / .
	if Right(strDomain, 1) = "/" then strDomain = Left(strDomain, Len(strDomain) - 1)
	' Full domain
	if full then strDomain = "http://" & strDomain & "/"

	' Return the function
	getURLdomain = strDomain
	
End Function


'-----------------------------------------------------------------------------------------
' Escludi By IP	
'-----------------------------------------------------------------------------------------
' Funzione: Esclude l'User dalle Statistiche in base agli IP
' Date: 	01.09.03 | 13.02.04
' Commenti: 		
'-----------------------------------------------------------------------------------------
Function exitCountByIP(controllaIP)

	Dim strAsgCheckIpRange
	
	' Get IP to filter
	strAsgSingleIP = Split(Trim(appAsgFilteredIPs), "," )
	exitCountByIP = false
	
	'Controlla ogni IP
	For Each appAsgFilteredIPs In strAsgSingleIP

		'Controlla se è necessario Bannare 1 solo IP o una Range
		'// Verifica se è presente un * per una Range di IP
		'// Rileva Range nel modello xxx.xxx.xxx.*
		If Right(appAsgFilteredIPs, 1) = "*" Then
			
			'Elimina *
			strAsgCheckIpRange = Replace(appAsgFilteredIPs, "*", "", 1, -1, 1)
		
			'Taglia l'IP in funzione alla lunghezza del presente oper verificare corrispondenza
			controllaIP = Left(controllaIP, Len(strAsgCheckIpRange))
			
			'Verifica le 2 stringhe ed Imposta a True se corrisponde la range
			If strAsgCheckIpRange = controllaIP then exitCountByIP = true
			
		'// Controlla intero indirizzo
		Else

			'Imposta a True se corrisponde l'IP
			If appAsgFilteredIPs = controllaIP then exitCountByIP = true

		End If
	
	Next 
		
End Function




'-----------------------------------------------------------------------------------------
' Format Empty String	
'-----------------------------------------------------------------------------------------
' Funzione: Formatta le stringhe in output per determinare un valore standard nel caso
'			sia nullo.
' Date: 	10.03.2004 |
' Commenti: 		
'-----------------------------------------------------------------------------------------
Function formatEmptyString(ByVal stringToFormat, ByVal stringType)
	
	Dim tmpValue
	
	'Esegui pulizia se la stringa è numerica
	If stringType = "Numeric" Then
		If NOT Len(stringToFormat) > 0 Then tmpValue = 0
	'Esegui pulizia se la stringa è testuale
	ElseIf stringType = "Text" Then
		If NOT Len(stringToFormat) > 0 Then tmpValue = ASG_UNKNOWN
	'Esegui pulizia se la stringa è testuale
	'// Formato a 2 caratteri!
	ElseIf stringType = "Text2chr" Then
		If NOT Len(stringToFormat) > 0 Then tmpValue = "1k"
	End If
	
	formatEmptyString = tmpValue

End Function


'-----------------------------------------------------------------------------------------
' Return the SERP value depending on the search engine and the referer type.
'-----------------------------------------------------------------------------------------
public function getSearchResultPage(argPagetype, argPagenumber)

	Dim lvTmp	

	if not IsNumeric(argPagenumber) Then argPagetype = 0

	Select case argPagetype
		
		' first page then usually no querystring value
		case 0
			lvTmp = -1
		case 1
			lvTmp = argPagenumber
		case 2
			lvTmp = (argPagenumber / 10) + 1
		case 3
			lvTmp = (argPagenumber + 2 / 10) + 1
		case 4
			lvTmp = argPagenumber + 1
		case 5
			lvTmp = (argPagenumber - 1 / 10)
		case 6
			lvTmp = (argPagenumber + 1 / 10) + 1
		case 7
			lvTmp = ((argPagenumber - 1) / 10) + 1
		case else
			lvTmp = -1
	
	end select
	
	' Return the function
	getSearchResultPage = lvTmp
		
End Function 


'-----------------------------------------------------------------------------------------
' Trim a string if the lenght is higher than the max length allowed and
' add [...] at the end of the returned string.
'
' @since 3.0
'-----------------------------------------------------------------------------------------
public function trimValue(argString, argMaxLength)

	if Len(argString) > argMaxLength then argString = Trim(Mid(argString, 1, argMaxLength)) & " [...]"
	
	' Return the function
	trimValue = argString
	
end function
	

	

'-----------------------------------------------------------------------------------------
' Get the type of the referer and return a numerical value.
' 1 - Direct request
' 2 - Internal referer : server
' 3 - Internal referer : other domain / mirror
' 4 - External referer : normal referer
' 5 - External referer : search engine
'
' @since 3.0
'-----------------------------------------------------------------------------------------
public function getRefererType(argURL, argDomain, argEngine)
		
	Dim lvType

	' Direct request
	if not Len(argURL) > 0 then
		lvType = 1 
	' Internal referer
	elseif argURL = ASG_OWNSERVER OR "http://" & argDomain & "/" = appAsgSiteURL then
		lvType = 2 
	' External referer
	else
		' Search engine result
		if Len(argEngine) > 0 then
			lvType = 5 
		else
			lvType = 4 
		end if
	end if
	
	' Return the function
	getRefererType = lvType
		
end function
	

'-----------------------------------------------------------------------------------------
' Compare two URLs to check if they are different and return a new URL if the
' URLs are the same but one has a final / .
' http://www.weppos.com = http://www.weppos.com/
' http://weppos.com = http://weppos.com/
'-----------------------------------------------------------------------------------------
public function compareURLsForFinalSlash(url, domain)
	
	Dim strDiff, newUrl, fulldomain
	' Build the full domain
	fulldomain = "http://" & domain & "/"
	newUrl = url

	' Remove the URL domain
	strDiff = Replace(url, "http://" & domain, "")

	' "http://" & domain == URL then add a /
	if Len(strDiff) = 0 then
		newUrl = url & "/"
	' 
	' elseif Len(strDiff) > 0 then
	end if
	' Return function
	compareURLsForFinalSlash = newUrl
	
end function
	

	

'-----------------------------------------------------------------------------------------
'
'
' @since 3.0
'-----------------------------------------------------------------------------------------
public function regexpExecuteEngine(argPattern, argString)
	  
	Dim objRegexp
	Dim objMatches
	Dim objMatch
	Dim objSubMatch
	Dim blnFound
	DIm lvTmp
	Set objRegexp = New RegExp
	  
	objRegexp.Pattern = argPattern
	objRegexp.IgnoreCase = true
	objRegexp.Global = true
	  
	Set objMatches = objRegexp.Execute(argString)
	if objMatches.Count > 0 then
		Set objMatch = objMatches(0)
		lvTmp = objMatch.SubMatches(0)
	end if

	' Return the falue
	regexpExecuteEngine = lvTmp
	
end function

%>
