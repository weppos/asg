<!--#include file="utils.datetime.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'



'-----------------------------------------------------------------------------------------
' FUNZIONI DI OUTPUT STATISTICO
'-----------------------------------------------------------------------------------------

Dim intStsOreDiOggi			'In uso nelle funzioni
Dim intStsGiorniPerMese
	

'-----------------------------------------------------------------------------------------
' Giorni per Mese
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	26.11.2003 | 26.11.2003
' Comment:	
'-----------------------------------------------------------------------------------------
Function GiorniPerMese(ByVal mese)
	
	'Conta i giorni dei mesi
	Select Case CInt(mese)
		Case 1 
			intStsGiorniPerMese = 31
		Case 2 
			'Controllo Anni Bisestili
			if IsDate("29/02/" & Year(Date())) then
				intStsGiorniPerMese = 29
			else
				intStsGiorniPerMese = 28
			end if 		
		Case 3 
			intStsGiorniPerMese = 31
		Case 4 
			intStsGiorniPerMese = 30
		Case 5 
			intStsGiorniPerMese = 31
		Case 6 
			intStsGiorniPerMese = 30
		Case 7 
			intStsGiorniPerMese = 31
		Case 8 
			intStsGiorniPerMese = 31
		Case 9 
			intStsGiorniPerMese = 30
		Case 10 
			intStsGiorniPerMese = 31
		Case 11 
			intStsGiorniPerMese = 30
		Case 12 
			intStsGiorniPerMese = 31
	End Select
	
End Function


'-----------------------------------------------------------------------------------------
' Media Giornaliera
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	16.11.2003 | 16.11.2003
' Comment:	
'-----------------------------------------------------------------------------------------
Function MediaGiorno(Accessi, Tipo, Cronologia)
	
	Dim intTmp
	If Cronologia = 1 Then		'Oggi
		intStsOreDiOggi = CInt(Hour(dtmAsgNow))
		'Da 00.00 a 00.59 dovrebbe dividere per 0
		'impostare ad 1 dato che Acchiardi insegna che per 0 non si divide! ;oP
		If intStsOreDiOggi = 0 Then intStsOreDiOggi = 1
		
		intTmp = FormatNumber(Accessi/intStsOreDiOggi, 1)
	ElseIf Cronologia = 2 Then	'Ieri
		intTmp = FormatNumber(Accessi/24, 1)
	End If
	
	MediaGiorno = intTmp

End Function


'-----------------------------------------------------------------------------------------
' Media Mensile
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	16.11.2003 | 16.11.2003
' Comment:	
'-----------------------------------------------------------------------------------------
Function MediaMese(Accessi, Tipo, Cronologia)
	
	Dim dtmTmp, intTmp
	If Cronologia = 1 Then		'Mese Corrente
		
		'Calcolo dei giorni
		GiorniPerMese(dtmAsgMonth)
			
		If Tipo = 1 Then		' x/Ora
			'Calcola ore dei giorni passati
			dtmTmp = 24*(CInt(dtmAsgDay) - 1)
			'Aggiungi le ore di oggi
			dtmTmp = dtmTmp + Hour(dtmAsgNow)
			'NON SI PUO' DIVIDERE PER O!
			'Il bug si presenta la prima ora del primo mese
			If dtmTmp = 0 Then dtmTmp = 1
			'Dividi gli accessi per le ore cacolate
			intTmp = FormatNumber(Accessi/dtmTmp, 1)
		
		ElseIf Tipo = 2 Then	' x/Giorno
			intTmp = FormatNumber(Accessi/CInt(dtmAsgDay), 1)
		End If
		
	ElseIf Cronologia = 2 Then	'Mese Scorso
		
		'Imposta la variabile temporanea
		dtmTmp = CInt(dtmAsgMonth) - 1
		If dtmTmp = 0 Then dtmTmp = 12
		
		'Calcolo dei giorni
		Call GiorniPerMese(dtmTmp)
			
		If Tipo = 1 Then		' x/Ora
			'Ore del mese passato
			dtmTmp = 24*(CInt(intStsGiorniPerMese))
			'Calcola
			intTmp = FormatNumber(Accessi/CInt(dtmTmp), 1)
		ElseIf Tipo = 2 Then	' x/Giorno
			intTmp = FormatNumber(Accessi/CInt(intStsGiorniPerMese), 1)
		End If
		
	End If
	
	MediaMese = intTmp

End Function


'-----------------------------------------------------------------------------------------
' Calculate the percentage value of the selected item.
'-----------------------------------------------------------------------------------------
public function calcPercValue(argTotalValue, argCurrentValue)

	Dim lvTmp
	if Clng(argTotalValue) = 0 then 
		lvTmp = FormatPercent(0, 2)
	else
		lvTmp = FormatPercent(argCurrentValue / argTotalValue, 2)
	end If
   
	' Return the function
	calcPercValue = lvTmp

end function


'-----------------------------------------------------------------------------------------
' Dichiarazioni Paginazione Avanzata Risultati
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	19.11.2003 | 19.11.2003
' Comment:			
'-----------------------------------------------------------------------------------------
Dim page
Dim RecordsPerPage
Dim loopAdvDataSorting

public function dimAdvDataSorting()
	
	page = Request.QueryString("page")
	' Allow user to choose the number of records in the page
	if IsNumeric(Request.QueryString("perpage")) AND Len(Request.QueryString("perpage")) > 0 then
		RecordsPerPage = Clng(Request.QueryString("perpage"))
	else
		RecordsPerPage = 30
	end if
		
	if len(page) > 0 And IsNumeric(page) Then
		page = CLng(page)
	else
		page = 1
	end If
		
end function


'-----------------------------------------------------------------------------------------
' Dichiarazioni Paginazione Avanzata Risultati Dettagli
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	19.11.2003 | 19.11.2003
' Comment:			
'-----------------------------------------------------------------------------------------
Dim detpage
Dim detRecordsPerPage
Dim loopAdvDetDataSorting

public function dimAdvDetDataSorting()
	
	detpage = Request.QueryString("detpage")
	' Allow user to choose the number of records in the page
	if IsNumeric(Request.QueryString("detperpage")) AND Len(Request.QueryString("detperpage")) > 0 then
		detRecordsPerPage = Clng(Request.QueryString("detperpage"))
	else
		detRecordsPerPage = 25
	end if
	
	if len(detpage) > 0 And IsNumeric(detpage) Then
		detpage = CLng(detpage)
	else
		detpage = 1
	end if
	
end function

'-----------------------------------------------------------------------------------------
' Append to the querystring the old values with one changed
'-----------------------------------------------------------------------------------------
public function appendToQuerystring(argNoappend)

	Dim lvTmp
	Dim lvNoappend
	Dim objItem
	
	lvTmp = "page=" & page
'	if Len(noappend) > 0 then 
		lvNoappend = "||" & argNoappend &  "||"
'	else
'		lvNoappend = "||"
'	end if

	if Len(Request.ServerVariables("QUERY_STRING")) > 0 then
		for each objItem in Request.QueryString
			if not Instr(lvNoappend, "||" & objItem & "||") > 0 AND objItem <> "page" then 
				lvTmp = lvTmp & "&amp;" & objItem & "=" & Request.QueryString(objItem) & ""
			end if
		next
	end if
	
	' Return the function
	appendToQuerystring = lvTmp

end function


'-----------------------------------------------------------------------------------------
' Read from querystrig the selected value, filter and format it following some
' basical rules.
'-----------------------------------------------------------------------------------------
public function formatSetting(qsfield, defaultvalue)
	
	Dim tmpValue
	tmpValue = Request.QueryString(qsfield)

	' Different possibilities
	select case qsfield
		
		' Period part
		case "periodm", "periody"
			if IsNumeric(tmpValue) AND Len(tmpValue) > 0 then
				tmpValue = Cint(tmpValue)
			else
				tmpValue = defaultvalue
			end if

		' Period part
		case "period"
			tmpValue = Right("0" & intAsgPeriodM, 2) & "-" & intAsgPeriodY

		' Sort order
		case "sortorder"
			if tmpValue <> "ASC" then 
				tmpValue = defaultvalue
			end if
		
		' Else
		case else
			if not Len(tmpValue) > 0 then
				tmpValue = defaultvalue
			end if
	
	end select
	
	' Return function
	formatSetting = tmpValue
	
end function


'-----------------------------------------------------------------------------------------
' Check if the database is mysql and return the right field value
' depending on current settings.
' It's used to return the righ SQL condition in the ORDER BY syntax
' without spending lots of lines describing the condition.
'
' @since 2.0
' @version 1.00 , 20050224
'-----------------------------------------------------------------------------------------
public function formatSortingField(argMysqlField, argAccessField, isMySql)
	
	Dim lvField
	
	if isMySql then
		lvField = argMysqlField
	else
		lvField = argAccessField
	end if
	
	' Return function
	formatSortingField = lvField
	
end function


'-----------------------------------------------------------------------------------------
' Check if the searching mode is enable and add to the SQL string
' the condition to search the database.
' It the query has no WHERE condition before this one the argument 'argFirstcondition'
' may prevent the error addint the proper WHERE syntax.
'
' @since 2.0
' @version 1.01 , 20050220
'-----------------------------------------------------------------------------------------
public function searchFor(argSQL, argFirstcondition)
	
	'Read for keywords to search for
	asgSearchfor = Trim(Request.QueryString("searchfor"))
	
	'Read for a field to search in
	asgSearchin = Trim(Request.QueryString("searchin"))

	'If there are keywords to search for and a field to search in then add SQL search string
	if Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 then
		
		' If this isn't the first WHERE condition add the string using AND operator
		if argFirstcondition = false then
			argSQL = argSQL & " AND " & asgSearchin & " LIKE '%" & filterSQLinput(asgSearchfor, true, true) & "%' "
		' If this is the first WHERE condition add the string using WHERE operator
		else
			argSQL = argSQL & " WHERE " & asgSearchin & " LIKE '%" & filterSQLinput(asgSearchfor, true, true) & "%' "
		end If

	'If there are no enough information to query the database then return the normal SQL query	
	Else

		argSQL = argSQL

	End If

	' Return function
	searchFor = argSQL
	
end function

'/**
' * Highlight searched keywords.
' * 
' * @param		
' * @param		
' * @return 	string § string with keywords highlighted.
' *
' * @since 		3.0
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function searchTerms(input, databaseField, searchFor, searchIn)

	' If some data has been searched and this is the database 
	' where you have searched in then highlight search terms
	if Len(searchFor) > 0 AND Len(searchIn) > 0 AND searchIn = databaseField then
		input = Replace(input, searchFor, "<span class=""highlighted"">" & searchFor & "</span>", 1, -1, vbTextCompare)
	end If
	
	' Return function
	'searchTerms = Server.HTMLEncode(argString)
	searchTerms = input

end function

'/**
' * If the string is longer than a fixed value it will be trimmed
' * X chrs on the left and Y chrs on the right, where X and Y are numerical values.
' * The trimmed part of the string will be replaced with '...' .
' * 
' * @param 		string § ip - dotted IP
' * @return 	array § an array with 2 indexes containing country information,
' *				an array with index 0 = false if the IP is invalid.
' *
' * @since 		2.0
' * @version	1.01 , 2005-02-20
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function stripValueTooLong(input, maxLenght, trimLeft, trimRight)

	Dim return
	
	if Len(input) > maxLenght then 
		return = Left(input, trimLeft) & "..." & Right(input, trimRight)
	else
		return = input
	end If

	stripValueTooLong = return

end function


'-----------------------------------------------------------------------------------------
' Icona di competenza dominio
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	12.03.2004 | 
' Comment:			
'-----------------------------------------------------------------------------------------
Function chooseDomainIcon(ByVal outputPage, ByVal prefixType)

	Dim strTmp
	strTmp = outputPage

If prefixType = "classic" Then

	'Taglia tutto il prefisso sito + http:// se non è una pagina sconosciuta
	'Versioni precedenti 1.2
	'If Mid(asgOutputPage, 1, Len("http://" & appAsgSiteURL)) = "http://" & appAsgSiteURL Then asgOutputPage = Mid(asgOutputPage, Len("http://" & appAsgSiteURL) + 1) 
	If Mid(strTmp, 1, Len(appAsgSiteURL)) = appAsgSiteURL Then strTmp = Mid(strTmp, Len(appAsgSiteURL) + 1) 

ElseIf prefixType = "visitors" Then

	'Taglia tutto il prefisso sito + http:// se non è una pagina sconosciuta
	'NB. La formula originale prevedeva
	'	'Taglia tutto il prefisso sito + http:// se non è una pagina sconosciuta
		'If Mid(asgOutputPage, 1, Len("http://" & appAsgSiteURL)) = "http://" & appAsgSiteURL Then asgOutputPage = Mid(asgOutputPage, Len("http://" & appAsgSiteURL) + 1) 
	If Mid(strTmp, 1, Len("http://" & appAsgSiteURL)) = "http://" & appAsgSiteURL Then strTmp = Mid(strTmp, Len("http://" & appAsgSiteURL) + 1) 

End If
	
	'Mostra una icona appropriata in base alla corrispondenza
	If outputPage <> strTmp Then
		chooseDomainIcon = vbCrLf & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "home.gif"" alt=""" & appAsgSiteURL & """ align=""absmiddle"" border=""0"" />"
	Else
		chooseDomainIcon = vbCrLf & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "arrow_small_dx.gif"" alt=""" & outputPage & """ align=""absmiddle"" border=""0"" />"
	End If
	
	'Ritorna il valore della variabile asgOutputPage
	'per consentire l'uso della successiva funzione di taglio stringa
	asgOutputPage = strTmp

End Function


'-----------------------------------------------------------------------------------------
' Manda a capo le stringhe troppo lunghe
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	15.03.2004
' Comment:	Tratto dal sito di Mems (www.oscarjsweb.com) - forum HTML.it		
'-----------------------------------------------------------------------------------------
Function shareWords(tempTXT, maxlenght)
	
	Dim Limit, arrTxt, tempLenght, start, intTmp
	Dim i, j
	
	Limit = maxlenght
	arrTXT = Split(tempTXT)
	
	For i = 0 To UBound(arrTXT)
	
	tempLenght = Len(arrTXT(i))
	If tempLenght > Limit Then

		intTmp = tempLenght / Limit
		If intTmp - CInt(intTmp) <> 0 Then
			intTmp = intTmp + 1
		End If
		start = 1
		
		For j = 1 To intTmp
			Response.Write Mid(arrTXT(i),start,Limit) & " "
			start = start + Limit
		Next
	Else
		Response.Write arrTXT(i) & " "
	End If
	
	Next
	
End Function


'-----------------------------------------------------------------------------------------
' Richiamo ultima versione da sito
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	30.03.2004 | 
' Comment:	Thanks to ToroSeduto
'-----------------------------------------------------------------------------------------
Function getLastVersion(ByVal siteUrl)
	
	Dim objXMLHTTP
	
	Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
	objXMLHTTP.Open "GET", siteUrl, false
	objXMLHTTP.Send     
	GetLastVersion = CStr(objXMLHTTP.ResponseText)
	Set objXMLHTTP = Nothing 
	
End Function


'-----------------------------------------------------------------------------------------
' Controllo nuove versioni
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	30.03.2004 | 
' Comment:	
'-----------------------------------------------------------------------------------------
Function checkUpdate(ByVal asgVersion, ByVal asgUpdate)
	
	'Dim strAsgLastVersion			'Ultima Versione dal sito
	Dim aryAsgLastVersion			'Array con info ultima versione
	
	strAsgLastVersion = GetLastVersion("http://www.weppos.com/asg/checkversion/check_update.asp?host=" & Server.URLEncode(Request.ServerVariables("HTTP_HOST")))
	aryAsgLastVersion = Split(strAsgLastVersion, "|")

	strAsgLastVersion = aryAsgLastVersion(0)
	dtmAsgLastUpdate = aryAsgLastVersion(1)
	urlAsgLastUpdate = aryAsgLastVersion(2)
	
	If LCase(asgVersion) <> LCase(strAsgLastVersion) then 	
		
		'Versioni differenti
		intAsgLastUpdate = 2
	
	Else
		
		If Clng(asgUpdate) <> Clng(dtmAsgLastUpdate) then 	
			'Versioni uguali ma differente aggiornamento
			intAsgLastUpdate = 3
		Else
			'Tutto combacia
			intAsgLastUpdate = 1
		End If
	
	End If

End Function



'-----------------------------------------------------------------------------------------
' Get information about the VbScript Engine installed on the server.
'
' @return	script engine type, version and release.
'
' @since 	2.0
' @version	1.01 , 2005-02-20
'-----------------------------------------------------------------------------------------
public function getScriptEngineInfo()
   
	Dim lvTmp
	lvTmp = ScriptEngine & "&nbsp;"
	lvTmp = lvTmp & ScriptEngineMajorVersion & "."
	lvTmp = lvTmp & ScriptEngineMinorVersion & "."
	lvTmp = lvTmp & ScriptEngineBuildVersion
   
	' Return the function
	getScriptEngineInfo = lvTmp

end function

%>