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


'Server MapPath
Dim strAsgMapPath
Dim strAsgMapPathTo
Dim strAsgMapPathIP  



'							========================================
'---------------------------   		Percorsi di collegamento		 -------------------------------------
'							========================================

'-----------------------------------------------------------------------------------------
' strAsgPathFolderDb
'-----------------------------------------------------------------------------------------
' Contiene il percorso alla cartella del server contenente i database
' NB. E' necessaria una cartella speciale con attivi i permessi di scrittura
' ed in genere è prevista da ogni provider che supporti l'uso di Access.
Const strAsgPathFolderDb = "mdb/" 
'
' Ecco alcuni esempi:
'
' // Sottocartella applicazione (myasg) '/mdb' - Include relativi
' Const strAsgPathFolderDb = "mdb/"
'
' // Sottocartella applicazione (myasg) '/mdb' - Include assoluti
' Const strAsgPathFolderDb = "/myasg/mdb/"
'
' // Sottocartella root  '/mdb' - Include relativi
' Const strAsgPathFolderDb = "../mdb/"
'
' // Sottocartella root  '/mdb' - Include assoluti
' Const strAsgPathFolderDb = "/mdb/"
'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------
' strAsgPathFolderWr
'-----------------------------------------------------------------------------------------
' Contiene il percorso alla cartella del server con attivi i permessi
' di scrittura e dove inserire il file inc_skin_file.asp
' E' possibile specificare la stessa del database.
Const strAsgPathFolderWr = "mdb/"




'							========================================
'---------------------------   			Nomi dei database			 -------------------------------------
'							========================================

'-----------------------------------------------------------------------------------------
' Statistiche
'-----------------------------------------------------------------------------------------
'
Const strAsgDatabaseSt = "dbstats" 

'-----------------------------------------------------------------------------------------
' IP
'-----------------------------------------------------------------------------------------
'
Const strAsgDatabaseIp = "ip-to-country" 


' Prefisso dei campi della tabella statistiche
' Utile se si vuole integrare l'applicazione in una tabella esistente per
' evitare conflitti
Const strAsgTablePrefix = "tblst_"


' Prefisso dei cookie
Const strAsgCookiePrefix = "ASG"




'							========================================
'---------------------------   		Collegamento al database		-------------------------------------
'							========================================

'-----------------------------------------------------------------------------------------
' Specifiche MapPath al database
'-----------------------------------------------------------------------------------------
strAsgMapPath = Server.MapPath(strAsgPathFolderDb & strAsgDatabaseSt & ".mdb")
strAsgMapPathTo = Server.MapPath(strAsgPathFolderDb & strAsgDatabaseSt & ".bak")
strAsgMapPathIP = Server.MapPath(strAsgPathFolderDb & strAsgDatabaseIp & ".mdb")




'							========================================
'---------------------------   Stringhe di connessione al database	-------------------------------------
'							========================================

'-----------------------------------------------------------------------------------------
' Microsoft Access 97	
'-----------------------------------------------------------------------------------------

' Driver Specifico
'-----------------
'strAsgConn = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & strAsgMapPath

' Driver OLEDB più veloce ed affidabile
'--------------------------------------
'strAsgConn = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & strAsgMapPath

'-----------------------------------------------------------------------------------------
' Microsoft Access 2000 - 2002 - 2003	
'-----------------------------------------------------------------------------------------

' Driver OLEDB più veloce ed affidabile
'--------------------------------------
strAsgConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strAsgMapPath



'-----------------------------------------------------------------------------------------
' Stampa di Debug stringa
'-----------------------------------------------------------------------------------------
If Request.QueryString("print") = "true" Then Response.Write(strAsgPathFolderDb & "<br />" & strAsgPathFolderWr & "<br />")


%>
