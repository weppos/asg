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
' Costruisci Riga Tabella Contenuti - Nessun Record
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	10.03.2004 | 
' Commenti:	
'-----------------------------------------------------------------------------------------
function BuildTableContNoRecord(ByVal colspanValue, ByVal message)
	
	'Verifica se è presente un messaggio alternativo.
	'Nel caso non sia definito usa quello
	'standard.
	If message = "standard" Then 
		message = strAsgTxtNoRecordInDatabase
	ElseIf message = "search" Then
		message = strAsgTxtSearchFoundNoResults
	End If 
			
	Response.Write(vbCrLf & "<tr class=""smalltext"" bgcolor=""" & strAsgSknTableContBgColour & """>")
	Response.Write(vbCrLf & "  <td colspan=""" & colspanValue & """ background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"">" & message & "</td>")
	Response.Write(vbCrLf & "</tr>")

end function
			

'-----------------------------------------------------------------------------------------
' Costruisci Riga Tabella Contenuti - Debug automatico icone non riconosciute
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	14.03.2004 | 
' Commenti:	
'-----------------------------------------------------------------------------------------
function BuildTableContCheckIcon(ByVal colspanValue, ByVal iconType, ByVal pageNum)
	
	Dim strAsgTableContent
	strAsgTableContent = ""
	
	strAsgTableContent = strAsgTableContent & vbCrLf & "<!-- Informazioni icone non riconosciute -->"
	strAsgTableContent = strAsgTableContent & vbCrLf & "<tr class=""smalltext"" align=""center"" valign=""top"">"
	strAsgTableContent = strAsgTableContent & vbCrLf & "  <td width=""100%"" colspan=""" & colspanValue & """><br /><img src=""" & strAsgSknPathImage & iconType & ".asp?icon=checkicon&page=" & pageNum & """ alt="""" /><br /></td>"
	strAsgTableContent = strAsgTableContent & vbCrLf & "</tr>"
	strAsgTableContent = strAsgTableContent & vbCrLf & "<!-- Informazioni icone non riconosciute -->"
			  
			
	If iconType = "browser" AND Session("blnAsgIconBrowser" & pageNum) <> "notified" AND blnAsgCheckIcon Then
	
		Response.Write(strAsgTableContent)
	
	ElseIf iconType = "os" AND Session("blnAsgIconOs" & pageNum) <> "notified" AND blnAsgCheckIcon Then
	
		Response.Write(strAsgTableContent)
	
	ElseIf iconType = "engine" AND Session("blnAsgIconEngine" & pageNum) <> "notified" AND blnAsgCheckIcon Then
	
		Response.Write(strAsgTableContent)
	
	End If
			
end function
			

'-----------------------------------------------------------------------------------------
' Costruisci Riga Tabella Contenuti - Spaziatore finale
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	14.03.2004 | 
' Commenti:	
'-----------------------------------------------------------------------------------------
function BuildTableContEndSpacer(ByVal colspanValue)

	Response.Write(vbCrLf & "<tr class=""smalltext"" bgcolor=""" & strAsgSknTableTitleBgColour & """>")
	Response.Write(vbCrLf & "  <td colspan=""" & colspanValue & """ background=""" & strAsgSknPathImage & strAsgSknTableTitleBgImage & """ height=""2""></td>")
	Response.Write(vbCrLf & "</tr>")

end function
			

'-----------------------------------------------------------------------------------------
' Costruisci Riga Footer - Linea Bordo
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	10.05.2004 | 
' Commenti:	
'-----------------------------------------------------------------------------------------
function BuildFooterBorderLine()

	Response.Write(vbCrLf & "<tr bgcolor=""" & strAsgSknTableLayoutBorderColour & """>")
	Response.Write(vbCrLf & "  <td align=""center"" height=""1""></td>")
	Response.Write(vbCrLf & "</tr>")

end function
			

%>