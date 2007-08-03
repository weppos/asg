<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Option Explicit %>
<!--#include file="config.asp" -->
<!--#include file="includes/inc_array_table.asp" -->
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
'	QuefileTo è un programma gratuito; potete modificare ed adattare il codice		'
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
'	QuefileTo programma è distribuito nella speranza che possa essere utile ma 	'
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


'// ATTENZIONE! Protezione statistiche.
'	Modificare solo se necessario e se sicuri.
'	Impostazioni errate possono compromettere la privacy.
Call AllowEntry("False", "False", "False", intAsgSecurity)


'Inizializza variabili
Dim strAsgTable		'Tabella da resettare
Dim strAsgTimerange
Dim strAsgWeekrange
Dim strAsgMsg		'Messaggio Reset
Dim blnAsgDoNotExecute
DIm dtmResetDateTime		'Holds a DateTime variable used to reset data


'-----------------------------------------------------------------------------------------
' Reset tabella
'-----------------------------------------------------------------------------------------
' Funzione:	Reset della tabella in base a parametri passati
' Data: 	03.04.2004
' Commenti:		
'-----------------------------------------------------------------------------------------
function ResetTableIndicated(ByVal databasetable)
	
	'Componi la parte di delete della stringa SQL
	strAsgSQL = "DELETE * FROM " & strAsgTablePrefix & databasetable & " "
	
	
	'-----------------------------------------------------------------------------------------
	' Reset settimanali e particolari
	'-----------------------------------------------------------------------------------------
	If databasetable = "Detail" AND IsNumeric(strAsgWeekrange) Then
			
			'Consenti il reset
			blnAsgDoNotExecute = False
			'Trasforma il valore in numerico
			strAsgWeekrange = CInt(strAsgWeekrange)
			'Calculate reset date
			dtmResetDateTime = DateAdd("ww", -strAsgWeekrange, dtmAsgDate)
			dtmResetDateTime = Year(dtmResetDateTime) & "/" & Month(dtmResetDateTime) & "/" & Day(dtmResetDateTime)
			'Crea la riga di condizione
			strAsgSQL = strAsgSQL & "WHERE Data < #" & dtmResetDateTime & "# "

	End If 


	'-----------------------------------------------------------------------------------------
	' Reset mensili e classici
	'-----------------------------------------------------------------------------------------
	'Reset completo
	If strAsgTimerange = "full" Then
			'
			'Consenti il reset
			blnAsgDoNotExecute = False
	'Reset escluso mese corrente
	ElseIf strAsgTimerange = "0" Then
		
		'Controllo coerenza reset mensile
		If databasetable <> "Detail" AND databasetable <> "IP" Then 
			strAsgSQL = strAsgSQL & "WHERE Mese <> '" & dtmAsgMonth & "-" & dtmAsgYear & "' "
			'Consenti il reset
			blnAsgDoNotExecute = False
		ElseIf databasetable = "Detail" Then 
			'Calculate reset date
			dtmResetDateTime = Date() 'DateAdd("m", -1, dtmAsgDate)
			dtmResetDateTime = Year(dtmResetDateTime) & "/" & Month(dtmResetDateTime) & "/" & 1
			'Crea la riga di condizione
			strAsgSQL = strAsgSQL & "WHERE Data < #" & dtmResetDateTime & "# "
			'Consenti il reset
			blnAsgDoNotExecute = False
		Else
			'Ferma il reset
			blnAsgDoNotExecute = True
		End If
	
	'Reset escluso numero di mesi
	ElseIf IsNumeric(strAsgTimerange) AND CInt(strAsgTimerange) > 0 Then
		
		Dim dateloop
		
		'Controllo coerenza reset mensile
		If databasetable <> "Detail" AND databasetable <> "IP" Then 
			
			'Consenti il reset
			blnAsgDoNotExecute = False
			'Crea la radice di condizione
			strAsgSQL = strAsgSQL & "WHERE Mese <> '" & Right("0" & Month(dtmAsgDate), 2) & "-" & Year(dtmAsgDate) & "' "

			'Trasforma il valore in numerico
			strAsgTimerange = Cint(strAsgTimerange)
			
			For dateloop = 1 to strAsgTimerange
				'Calcola la differenza di tempo
				dtmAsgDate = DateAdd("m", -1, dtmAsgDate)
				'Prepara la condizione
				strAsgSQL = strAsgSQL & "AND Mese <> '" & Right("0" & Month(dtmAsgDate), 2) & "-" & Year(dtmAsgDate) & "' "
			Next

		ElseIf databasetable = "Detail" Then 
			
			'Consenti il reset
			blnAsgDoNotExecute = False
			'Trasforma il valore in numerico
			strAsgTimerange = Cint(strAsgTimerange)
			'Calculate reset date
			dtmResetDateTime = DateAdd("m", -strAsgTimerange, dtmAsgDate)
			dtmResetDateTime = Year(dtmResetDateTime) & "/" & Month(dtmResetDateTime) & "/" & Day(dtmResetDateTime)
			'Crea la riga di condizione
			strAsgSQL = strAsgSQL & "WHERE Data < #" & dtmResetDateTime & "# "

		Else
			'Ferma il reset
			blnAsgDoNotExecute = True
		End If
		
	Else
			'Ferma il reset
			blnAsgDoNotExecute = True
	End If

'Response.Write(strAsgSQL) : Response.End()
	If Not blnAsgDoNotExecute = True Then
		objAsgConn.Execute(strAsgSQL)
		strAsgMsg = strAsgMsg & strAsgTxtTable & "&nbsp;" & databasetable & "&nbsp;" & strAsgTxtCorrectlyDeleted & " <br />"
	End If

end function 

strAsgTable = Trim(Request.QueryString("table"))
strAsgTimerange = Trim(Request.QueryString("timerange"))
strAsgWeekrange = Trim(Request.QueryString("weekrange"))
'On Error Resume Next


'Reset generale
If strAsgTable = 0 Then
		
	For looptmp = 1 to Ubound(aryAsgTable)
	
		Call ResetTableIndicated(aryAsgTable(looptmp, 1))
	
	Next
	
'Reset selezionato
ElseIf strAsgTable <> 0 AND Len(strAsgTable) > 0 Then
	
	Call ResetTableIndicated(aryAsgTable(strAsgTable, 1))
	
Else
	
	'Reset Server Objects
	Set objAsgRs = Nothing
	objAsgConn.Close
	Set objAsgConn = Nothing
	Response.Redirect "settings_reset.asp?msg=error"

End If

'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing


'-----------------------------------------------------------------------------------------
' Rinomina File
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	27.12.2003 |
' Commenti:	
'-----------------------------------------------------------------------------------------
function RinominaFile(fileFrom, fileTo)

	Dim objFso, objFile
	
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.GetFile(fileFrom)
	objFile.Copy fileTo, True
	objFile.Delete True
		
	Set objFso = Nothing
	Set objFile = Nothing

end function '-----------------------------------------------------------------------------------------
' Ripristina File
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	27.12.2003 | 27.12.2003
' Commenti:	
'-----------------------------------------------------------------------------------------
function RipristinaFile(fileFrom, fileTo)

	Dim objFso
	
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")
	objFso.CopyFile fileTo, fileTo & ".bak", true 
	objFso.DeleteFile fileTo
	objFso.MoveFile fileFrom, fileTo
	objFso.DeleteFile fileTo & ".bak"

	Set objFso = Nothing

end function '-----------------------------------------------------------------------------------------
' Compatta il Database
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	|
' Commenti:	
'-----------------------------------------------------------------------------------------
function CompactAccessDatabase()
	
	Dim strAsgDb, strAsgDbTo
	Dim objAsgJro
	
	strAsgDb = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strAsgMapPath
	strAsgDbTo = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strAsgMapPathTo
	
	set objAsgJro = CreateObject("jro.JetEngine") 
	objAsgJro.CompactDatabase strAsgDb, strAsgDbTo
	Set objAsgJro = Nothing 
	
end function 'Compatta il database dopo i reset	
Call CompactAccessDatabase()

'Dopo aver compattato il database
'ripristina la versione precedente
'con quella compattata
Call RinominaFile(strAsgMapPathTo, strAsgMapPath)
'Call RipristinaFile(strAsgMapPathTo, strAsgMapPath)


'Nel caso si siano verificati errori valorizza una variabile
'e mostrali poi a video continuando l'esecuzione
If err <> 0 then 
    strAsgMsg = strAsgMsg & strAsgTxtError & ": <br />" & err.description & "<br />" 
Else 
    strAsgMsg = strAsgMsg & "<br />" & strAsgTxtDatabaseSuccessfullyCompactedOn & "<br /><span class=""notetext"">" & strAsgMapPathTo & "</span><br />"
    strAsgMsg = strAsgMsg & strAsgTxtDatabaseSuccessfullyRenamedTo & "<br /><span class=""notetext"">" & strAsgMapPath & "</span><br />"
End if 


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%= strAsgSiteName %> | ASP Stats Generator <%= strAsgVersion %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<meta name="copyright" content="Copyright (C) 2003-2004 Carletti Simone" />
<link href="stile.css" rel="stylesheet" type="text/css" />

<!-- 	ASP Stats Generator <%= strAsgVersion %> è una applicazione gratuita 
		per il monitoraggio degli accessi e dei visitatori ai siti web 
		creata e sviluppata da Simone Carletti.
		
		Puoi scaricarne una copia gratuita sul sito ufficiale http://www.weppos.com/ -->

</head>
<!--#include file="includes/header.asp" -->
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
		  <tr align="center" valign="middle">
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtResetSettings) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<table width="70%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" align="center" height="16"><%= UCase(strAsgTxtExecutionReport) %></td>
		  </tr>
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
			<td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="100%" align="center"><br />
			<%= strAsgMsg %><br /><br />
			</td>
		  </tr>
		<%
				
		'// Row - End table spacer			
		Call BuildTableContEndSpacer(2)

		%>
		</table><br />
<%

' Footer
Response.Write(vbCrLf & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"">")
'// Row - Footer Border Line
Call BuildFooterBorderLine()

' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write("<tr align=""center"" valign=""middle"">")
Response.Write("<td align=""center"" background=""" & strAsgSknPathImage & strAsgSknTableBarBgImage & """ bgcolor=""" & strAsgSknTableBarBgColour & """ height=""20"" class=""footer""><a href=""http://www.asp-stats.com/"" class=""linkfooter"" title=""ASP Stats Generator Homepage"">ASP Stats Generator</a> [" & strAsgVersion & "] - &copy; 2003-2007 <a href=""http://www.weppos.com/"" class=""linkfooter"" title=""Weppos.com Homepage"">weppos</a>")
If blnAsgElabTime Then Response.Write(" - " & strAsgTxtThisPageWasGeneratedIn & "&nbsp;" & FormatNumber(Timer() - startAsgElab, 4) & "&nbsp;" & strAsgTxtSeconds)
Response.Write("</td>")
Response.Write("</tr>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write("</table>")
Response.Write("</td></tr>")
Response.Write("</table>")
Response.Write("</td></tr>")
Response.Write("</table>")

%>
<!-- footer -->
<!--#include file="includes/footer.asp" -->
</body></html>