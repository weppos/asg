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


'-------------------------------------------------------------------------------'
'																				'
'	Questo file è stato creato per consentire l'aumento delle prestazioni in	'
'	elaborazione ed allo stesso tempo ovviare al problema dell'uso dei percorsi	'
'	relativi a path superiori su Win2003 Server.								'
'																				'
'	Qui saranno/sono inclusi tutti i file che necessiteranno di essere eseguiti	'
'	dall'applicazione generale ma non al file di conteggio mentre il file 		'
'	config_common.asp è stato dedicato all'uso principale dal file di conteggio.'
'	In questo modo il file di conteggio non sarà obbligato a caricare in		'
'	memoria inutili variabili di elaborazione utili solo al processo di report	' 
'	rallenterebbero l'applicazione.												'
'																				'
'-------------------------------------------------------------------------------'



'-------------------------------------------------------------------------------'
' Includi il file di skin!
'-------------------------------------------------------------------------------'
'
' 									>>	 DA ADATTARE!	<<
'
' Enter the path of the folder where the skin file 
' has been uploaded.

%><!--#include file="mdb/inc_skin_file.asp" --><%
'
'
'-------------------------------------------------------------------------------'


'-------------------------------------------------------------------------------'
' Includi le informazioni generiche di configurazione
'-------------------------------------------------------------------------------'
%><!--#include file="config_common.asp" --><%

'-------------------------------------------------------------------------------'
' Includi le informazioni di gestione dei report statistici
'-------------------------------------------------------------------------------'
%><!--#include file="includes/functions_stats.asp" --><%

'-------------------------------------------------------------------------------'
' Includi le informazioni sullo sviluppo del layout tramite funzioni
'-------------------------------------------------------------------------------'
%><!--#include file="includes/functions_layout.asp" --><%

	
	'---------------------------------------------------
	'	Dimension variables : show icons
	'---------------------------------------------------
	Dim strAsgIconaTemp
	Dim Index
	
	'---------------------------------------------------
	'	Dimension variables : sorting records
	'---------------------------------------------------
	Dim strAsgSortBy
	Dim strAsgSortByFld
	Dim strAsgSortOrder
	
	'---------------------------------------------------
	'	Dimension variables : report data search engine
	'---------------------------------------------------
	Dim strAsgSQLsearchstring			'Holds the search string to query database
	Dim asgSearchfor					'Holds the keywords to search for
	Dim asgSearchIn						'Holds the name of the table to search in
	
	'---------------------------------------------------
	'	Dimension variables : other elaborations
	'---------------------------------------------------
	Dim strAsgUnknownIcon				'Holds unknown icons information
	
	'---------------------------------------------------
	'	Controllo aggiornamenti
	'---------------------------------------------------
	Dim strAsgLastVersion		'
	Dim dtmAsgLastUpdate		'
	Dim intAsgLastUpdate		'Casiste possibili:
		intAsgLastUpdate =	0 	'  - variabile non in calcolo
								'1 - tutto combacia
								'2 - differente versione e data
								'3 - uguale versione ma differente data
	Dim urlAsgLastUpdate		'
	

'-------------------------------------------------------------------------------'
' Includi le informazioni sulla traduzione in uso
'-------------------------------------------------------------------------------'

'//	Italiano %>
<!--include file="languages/italiano.asp" --><%

'//	English %>
<!--#include file="languages/english.asp" --><%

'//	Espanol %>
<!--include file="languages/espanol.asp" --><%


'---------------------------------------------------
'	Apri connessione al Database
'---------------------------------------------------
If blnConnectionIsOpen = False Then
	'Se si usano variabili application apri la connessione
	'dato che non è stata aperta per gestire le risorse
	'del file di conteggio.
	objAsgConn.Open strAsgConn
'---------------------------------------------------
End If
'---------------------------------------------------


'-----------------------------------------------------------------------------------------
' Check version for update!
'-----------------------------------------------------------------------------------------
' Controllo differenza data ed esecuzione solo se amministratore
If Clng(Clng(Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)) - blnAsgCheckUpdate) > 7 AND Session("AsgLogin") = "Logged" Then

	'Esegui controllo versione
	Call CheckUpdate(strAsgVersion, dtmAsgUpdate)
	
	'Aggiornamento informazioni sul controllo
	strAsgSQL = "UPDATE "&strAsgTablePrefix&"Config SET Opt_Check_Update = " & Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & ""
	objAsgConn.Execute(strAsgSQL)
	
	'Se si utilizzano le variabili Application aggiornale
	If blnApplicationConfig Then
						
		'Aggiorna Variabili Application
		Application("blnAsgCheckUpdate") = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)

	End If
	
End If

%>
