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

			
			'********** 		NOTE DI SVILUPPO 		*********
			' Inizio conversione nomi e variabili in inglese	'
			'****************************************************
			

'-----------------------------------------------------------------------------------------
' FUNZIONI DI CONTEGGIO
'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------
' Taglia QueryString	
'-----------------------------------------------------------------------------------------
' Funzione:	Taglia i caratteri della QueryString URL
' Data: 	01.09.03 | 01.09.03
' Commenti:	Tratto dal sito di Mems (www.oscarjsweb.com) forum HTML.it		
'-----------------------------------------------------------------------------------------
function StripURLquerystring(strURL)

	strToStrip = instr(strURL, "?")
	if strToStrip then strBuffer = left(strURL, strToStrip-1) else strBuffer = strURL
	StripURLquerystring = strBuffer
	
end function


'-----------------------------------------------------------------------------------------
' Taglia Protocollo	
'-----------------------------------------------------------------------------------------
' Funzione:	Taglia il protocollo completo dell'URL restituendo dominio di I, II e III liv.
' Data: 	01.09.03 | 01.09.03
' Commenti:	Taglia http:// | http://www. mantenendo però le path successive
'			http://www.weppos.com | http://weppos.com --> weppos.com		
'			http://www.weppos.sonoio.com | http://weppos.sonoio.com --> weppos.sonoio.com		
'-----------------------------------------------------------------------------------------
function StripURLprotocol(strURL)

	strToStrip = instr(strURL, "://")
	if strToStrip then strBuffer = right(strURL, len(strURL) - (3 + strToStrip - 1)) else strBuffer = strURL
	if left(strBuffer, 4) = "www." then strBuffer = right(strBuffer, len(strBuffer) - 4)
	StripURLprotocol = strBuffer
	
end function


'-----------------------------------------------------------------------------------------
' Ricava dominio	
'-----------------------------------------------------------------------------------------
' Funzione:	Ricava l'esclusivo dominio da un URL di partenza
' Data: 	01.09.03 | 01.09.03
' Commenti:	Funziona anche se non è presente lo slash finale!
'			http://www.weppos.com | http://www.weppos.com/ | http://www.weppos.com/.../ --> www.weppos.com		
'			http://weppos.com | http://weppos.com/ | http://weppos.com/.../ --> weppos.com		
'-----------------------------------------------------------------------------------------
function GetURLdomain(strURL)
	
	strToStrip = instr(strURL, "://")
	If strToStrip then strBuffer = right(strURL, len(strURL) - (3 + strToStrip - 1)) else strBuffer = strURL
	strToStrip = instr(strBuffer, "/")
	If strToStrip > 0 Then
		GetURLdomain = Left(strBuffer, strToStrip)
	Else
		GetURLdomain = strBuffer & "/"
	End If
	
end function


'-----------------------------------------------------------------------------------------
' Escludi By IP	
'-----------------------------------------------------------------------------------------
' Funzione: Esclude l'User dalle Statistiche in base agli IP
' Data: 	01.09.03 | 13.02.04
' Commenti: 		
'-----------------------------------------------------------------------------------------
function ExitCountByIP(ByVal controllaIP)

	'Richiama gli IP da filtrare e imposta come False la corrispondenza
	strAsgSingleIP = Split(Trim(strAsgFilterIP), "," )
	blnExitCount = False
	Dim strAsgCheckIpRange
	
	'Controlla ogni IP
	For Each strAsgFilterIP In strAsgSingleIP

		'Controlla se è necessario Bannare 1 solo IP o una Range
		'// Verifica se è presente un * per una Range di IP
		'// Rileva Range nel modello xxx.xxx.xxx.*
		If Right(strAsgFilterIP, 1) = "*" Then
			
			'Elimina *
			strAsgCheckIpRange = Replace(strAsgFilterIP, "*", "", 1, -1, 1)
		
			'Taglia l'IP in funzione alla lunghezza del presente oper verificare corrispondenza
			controllaIP = Left(controllaIP, Len(strAsgCheckIpRange))
			
			'Verifica le 2 stringhe ed Imposta a True se corrisponde la range
			If strAsgCheckIpRange = controllaIP Then blnExitCount = True
			
		'// Controlla intero indirizzo
		Else

			'Imposta a True se corrisponde l'IP
			If strAsgFilterIP = controllaIP Then blnExitCount = True

		End If
	
	Next 
		
end function


'-----------------------------------------------------------------------------------------
' Escludi By Cookie	
'-----------------------------------------------------------------------------------------
' Funzione: Esclude l'User dalle Statistiche in base al cookie impostato
' Data: 	01.09.03 | 28.03.04
' Commenti: 
'-----------------------------------------------------------------------------------------
function ExitCountByCookie()
	
	blnExitCount = False

	If Request.Cookies(strAsgCookiePrefix& "exitcount") = "excludepc" Then
		blnExitCount = True
	End If
	
end function


'-----------------------------------------------------------------------------------------
' DottedIp
'-----------------------------------------------------------------------------------------
' Funzione: Converte l'IP dell'Utente in un formato leggibile nel Db
' Data: 	06.12.03 | 15.02.04
' Commenti: http://ip-to-country.webhosting.info/node/view/55
'-----------------------------------------------------------------------------------------
Dim arrAsgIp, strDottedIp

Public function DottedIp(ByVal userIP)

	If Trim("[]" & userIP) <> "[]" Then
		arrAsgIp = Split(userIP,".")
		strDottedIp = arrAsgIp(0)*16777216 + arrAsgIp(1)*65536 + arrAsgIp(2)*256 + arrAsgIp(3)
	Else
		strDottedIp = "noIp"
	End If
	
	'Ritorna Funzione
	DottedIp = strDottedIp

end function


'-----------------------------------------------------------------------------------------
' Format Empty String	
'-----------------------------------------------------------------------------------------
' Funzione: Formatta le stringhe in output per determinare un valore standard nel caso
'			sia nullo.
' Data: 	10.03.2004 |
' Commenti: 		
'-----------------------------------------------------------------------------------------
function FormatEmptyString(ByVal stringToFormat, ByVal stringType)
	
	'Esegui pulizia se la stringa è numerica
	If stringType = "Numeric" Then
		If NOT Len(stringToFormat) > 0 Then strtmp = 0
	'Esegui pulizia se la stringa è testuale
	ElseIf stringType = "Text" Then
		If NOT Len(stringToFormat) > 0 Then strtmp = "(unknown)"
	'Esegui pulizia se la stringa è testuale
	'// Formato a 2 caratteri!
	ElseIf stringType = "Text2long" Then
		If NOT Len(stringToFormat) > 0 Then strtmp = "1k"
	End If
	
	FormatEmptyString = strtmp

end function


'-----------------------------------------------------------------------------------------
' Ottieni Pagina	
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	23.03.2004 | 
' Commenti:	Funzione sviluppata da ToroSeduto	
'-----------------------------------------------------------------------------------------
function GetSearchResultPage(Tipo, Numero)
	
	If Not IsNumeric(Numero) Then Tipo = -1
	
	Select case Tipo
		
		case 0
			strtmp = -1
		case 1
			strtmp = Numero
		case 2
			strtmp = (Numero / 10) + 1
		case 3
			strtmp = (Numero+2 / 10) + 1
		case 4
			strtmp = Numero+1
		case 5
			strtmp = (Numero-1 / 10)
		case 6
			strtmp = (Numero+1 / 10) + 1
		case else
			strtmp = -1
	
	End select
	
	GetSearchResultPage = strtmp
		
end function 


%>
