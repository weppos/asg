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
