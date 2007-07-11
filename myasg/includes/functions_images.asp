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
			

'-----------------------------------------------------------------------------------------
' Icona Lingua	
'-----------------------------------------------------------------------------------------
' Funzione:	Restituisce una icona in base al nome della lingua
' Data: 	19.11.2003 | 25.02.2004
' Commenti:	Da Classe 2.0.4 versione restituita in 
'			[code] - Lingua Estesa
'-----------------------------------------------------------------------------------------
function ShowIconLanguage(languages)
	
	'Italiano
	'// Classe 2.0.4
	If InStr(1, languages, "Italiano", 1) > 0  Then
		Response.Write "it.png"
	
	'Inglese
	'// Classe 2.0.4
	ElseIf InStr(1, languages, "Inglese", 1) > 0  Then
		Response.Write "gb.png"
	
	'Tedesco
	'// Classe 2.0.4
	ElseIf InStr(1, languages, "Tedesco", 1) > 0  Then
		Response.Write "de.png"
	
	'Spagnolo
	'// Classe 2.0.4
	ElseIf InStr(1, languages, "Spagnolo", 1) > 0  Then
		Response.Write "es.png"
	
	'Irlandese
	'// Classe 2.0.4
	ElseIf InStr(1, languages, "Irlandese", 1) > 0  Then
		Response.Write "ir.png"
	
	'Russo
	'// Classe 2.0.4
	ElseIf InStr(1, languages, "Russo", 1) > 0  Then
		Response.Write "ru.png"
	
	'Giapponese
	'// Classe 2.0.4
	ElseIf InStr(1, languages, "Giapponese", 1) > 0  Then
		Response.Write "jp.png"
	
	'Olandese
	'// Classe 2.0.4c
	ElseIf InStr(1, languages, "Olandese", 1) > 0  Then
		Response.Write "nl.png"
	
	'Francese
	'// Classe 2.0.4c
	ElseIf InStr(1, languages, "Francese", 1) > 0  Then
		Response.Write "fr.png"
	
	'Portoghese
	'// Classe 2.0.4d
	ElseIf InStr(1, languages, "Portoghese", 1) > 0  Then
		Response.Write "pt.png"
	
	'Coreano
	'// Classe 2.0.4d
	ElseIf InStr(1, languages, "Coreano", 1) > 0  Then
		Response.Write "kr.png"
	
	'Norvegese
	'// Classe 2.1
	ElseIf InStr(1, languages, "Norvegese", 1) > 0  Then
		Response.Write "no.png"
	
	'Rumeno
	'// Classe 2.1
	ElseIf InStr(1, languages, "Rumeno", 1) > 0  Then
		Response.Write "ro.png"
	
	'Danese
	'// Classe 2.1
	ElseIf InStr(1, languages, "Danese", 1) > 0  Then
		Response.Write "dk.png"
	
	'Svedese
	'// Classe 2.1
	ElseIf InStr(1, languages, "Svedese", 1) > 0  Then
		Response.Write "se.png"
	
	'Cinese
	'// Classe 2.1
	ElseIf InStr(1, languages, "Cinese", 1) > 0  Then
		Response.Write "cn.png"
	
	'Ebreo
	'// Classe 2.1
	ElseIf InStr(1, languages, "Ebreo", 1) > 0  Then
		Response.Write "il.png"
	
	'Turco
	'// Classe 2.1
	ElseIf InStr(1, languages, "Turco", 1) > 0  Then
		Response.Write "tr.png"
	
	'Polacco
	'// Classe 3.x
	ElseIf InStr(1, languages, "Polacco", 1) > 0  Then
		Response.Write "pl.png"
	
	'Sloveno
	'// Classe 3.x
	ElseIf InStr(1, languages, "Sloveno", 1) > 0  Then
		Response.Write "sk.png"
	
	'Ceco
	'// Classe 3.x
	ElseIf InStr(1, languages, "Ceco", 1) > 0  Then
		Response.Write "cz.png"
	
	'Finlandese
	'// Classe 3.x
	ElseIf InStr(1, languages, "Finlandese", 1) > 0  Then
		Response.Write "fi.png"
	
	'Croato
	'// Classe 3.x
	ElseIf InStr(1, languages, "Croato", 1) > 0  Then
		Response.Write "hr.png"
	
	'Bulgaro
	'// Classe 3.x
	ElseIf InStr(1, languages, "Bulgaro", 1) > 0  Then
		Response.Write "bg.png"
	
	'Arabo
	'// Classe 3.x
	ElseIf InStr(1, languages, "Arabo", 1) > 0  Then
		Response.Write "sa.png"
	
	'Indiano
	'// Classe 3.x
	ElseIf InStr(1, languages, "Indiano", 1) > 0  Then
		Response.Write "in.png"
	
	'Ungherese
	'// Classe 3.x
	ElseIf InStr(1, languages, "Ungherese", 1) > 0  Then
		Response.Write "hu.png"
	
	'Greco
	'// Classe 3.x
	ElseIf InStr(1, languages, "Greco", 1) > 0  Then
		Response.Write "gr.png"
	

	'Mostra Sconosciuto
	Else
		Response.Write "unknown.png"
	
	End If

end function


'-----------------------------------------------------------------------------------------
' Icona Filtro indirizzo
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	06.04.2004
' Commenti:	
'-----------------------------------------------------------------------------------------
function ShowIconFilterIp(ByVal ipaddress)
					
	'Filter IP
	'// Link PopUp
	Response.Write(vbCrLf & "<a href=""JavaScript:openWin('popup_filter_ip.asp?IP=" & ipaddress & "','Filter','toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=200')"" title=""" & strAsgTxtFilterIPaddr & """>")
								
	'// L'IP è escluso
	If InStr(1, strAsgFilterIP, ipaddress, 1) > 0 Then
										
		'// Icona esclusione
		Response.Write(vbCrLf & "<img src=""" & strAsgSknPathImage & "locked_icon.gif"" alt=""" &  strAsgTxtFilterIPaddr & """ border=""0"" align=""absmiddle"" />")
									
	'// L'IP è escluso
	Else
									
		'// Icona esclusione
		Response.Write(vbCrLf & "<img src=""" & strAsgSknPathImage & "unlocked_icon.gif"" alt=""" &  strAsgTxtFilterIPaddr & """ border=""0"" align=""absmiddle"" />")
								
	End If
								
	'// Chiudi Link PopUp
	Response.Write("</a>")
	
end function

%>