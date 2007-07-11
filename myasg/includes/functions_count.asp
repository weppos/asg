<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright 2003-2006 - Carletti Simone										'
'-------------------------------------------------------------------------------'
'																				'
'	Autore:																		'
'	--------------------------													'
'	Simone Carletti (weppos)													'
'																				'
'	Collaboratori 																'
'	[che ringrazio vivamente per l'impegno ed il tempo dedicato]				'
'	--------------------------													'
'	@ imente 			- www.imente.it | www.imente.org						'
'	@ ToroSeduto		- www.velaforfun.com									'
'																				'
'	Hanno contribuito															'
'	[anche a loro un grazie speciale per le idee apportate]						'
'	--------------------------													'
'	@ Gli utenti del forum con consigli e segnalazioni							'
'	@ subxus (suggerimento generazione grafica dei report)						'
'																				'
'	Verifica le proposte degli utenti, implementate o da implementare al link	'
'	http://www.weppos.com/forum/forum_posts.asp?TID=140&PN=1					'
'																				'
'-------------------------------------------------------------------------------'
'																				'
'	Informazioni sulla Licenza													'
'	--------------------------													'
'	Questo è un programma gratuito; potete modificare ed adattare il codice		'
'	(a vostro rischio) in qualsiasi sua parte nei termini delle condizioni		'
'	della licenza che lo accompagna.											'
'																				'
'	Non è consentito utilizzare l'applicazione per conseguire ricavi 			'
'	personali, distribuirla, venderla o diffonderla come una propria 			'
'	creazione anche se modificata nel codice, senza un esplicito e scritto 		'
'	consenso dell'autore.														'
'																				'
'	Potete modificare il codice sorgente (a vostro rischio) per adattarlo 		'
'	alle vostre esigenze o integrarlo nel sito; nel caso le funzioni possano	'
'	essere di utilità pubblica vi invitiamo a comunicarlo per poterle 			'
'	implementare in una futura versione e per contribuire allo sviluppo 		'
'	del programma.																'
'																				'
'	In nessun caso l'autore sarà responsabile di danni causati da una 			'
'	modifica, da un uso non corretto o da un uso qualsiasi 						'
'	dell'applicazione.															'
'																				'
'	Nell'utilizzo devono rimanere intatte tutte le informazioni sul 			'
'	copyright; è possibile modificare o rimuovere unicamente le indicazioni 	'
'	espressamente specificate.													'
'																				'
'	Numerose ore sono state impiegate nello sviluppo del progetto e, anche 		'
'	se non vincolante ai fini dell'uso, sarebbe gratificante l'inserimento		'
'	di un link all'applicazione sul vostro sito.								'
'																				'
'	NESSUNA GARANZIA															'
'	------------------------- 													'
'	Questo programma è distribuito nella speranza che possa essere utile ma 	'
'	senza GARANZIA DI ALCUN GENERE.												'
'	L'utente si assume tutte le responsabilità nell'uso.						'
'																				'
'-------------------------------------------------------------------------------'

'********************************************************************************'
'*																				*'	
'*	VIOLAZIONE DELLA LICENZA													*'
'*	 																			*'
'*	L'utilizzo dell'applicazione violando le condizioni di licenza comporta la 	*'
'*	perdita immediata della possibilità d'uso ed è PERSEGUIBILE LEGALMENTE!		*'
'*																				*'
'********************************************************************************'
			
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
