<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
' Copyright 2003-2006 - Carletti Simone										'
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


'SPERIMENTALE!
Const strAsgSknPathImage      = "images/" 'Percorso della cartella immagini
'-----------------------------------------------------------------------------------------
' Generali	
'-----------------------------------------------------------------------------------------
Const strAsgSknPageBgColour    = "#F9F9F9" 'Colore di sfondo delle pagine
Const strAsgSknPageBgImage    = "" 'Immagine di sfondo delle pagine
Const strAsgSknPageWidth      = "900" 'Larghezza delle pagine <br />&nbsp;ATTENZIONE! Variare solo se necessario! Una bassa risoluzione causerà problemi nella visualizzazione
'-----------------------------------------------------------------------------------------
' Tabelle di Layout
'-----------------------------------------------------------------------------------------
Const strAsgSknTableLayoutBorderColour    = "#999999" 'Colore del bordo della tabella Layout
Const strAsgSknTableLayoutBgColour    = "#F6F6F6" 'Colore di sfondo della tabella Layout
Const strAsgSknTableLayoutBgImage    = "" 'Immagine di sfondo della tabella Layout
Const strAsgSknTableBarBgColour    = "" 'Colore di sfondo della Barra Header e Footer
Const strAsgSknTableBarBgImage    = "layout/bar_bg.jpg" 'Immagine di sfondo della Barra Header e Footer
'-----------------------------------------------------------------------------------------
' Tabelle Dati
'-----------------------------------------------------------------------------------------
Const strAsgSknTableTitleBgColour    = "" 'Colore di sfondo del titolo della tabella contenuti
Const strAsgSknTableTitleBgImage    = "layout/small_bar_bg.jpg" 'Immagine di sfondo del titolo della tabella contenuti
Const strAsgSknTableContBgColour    = "#ECEAF2" 'Colore di sfondo della tabella contenuti
Const strAsgSknTableContBgImage    = "" 'Immagine di sfondo della tabella contenuti
'-----------------------------------------------------------------------------------------
' Configurazioni opzionali
'-----------------------------------------------------------------------------------------
'Mostra il tempo di elaborazione
Const blnAsgElabTime = True 'Mostra il tempo di elaborazione.<br />&nbsp;Accetta valori <strong>True</strong> o <strong>False</strong>

%>
