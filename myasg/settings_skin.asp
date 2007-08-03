<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="config.asp" -->
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
Call AllowEntry("False", "False", "False", intAsgProtezione)


'Reset Server Objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing

Dim strAsgInput



'on error resume next
pathFrom = Server.MapPath(strAsgPathFolderWr & "inc_skin_file.asp")
pathFrom2 = Server.MapPath(strAsgPathFolderWr & "inc_skin_file_temp.asp")

Set objFso = CreateObject("Scripting.FileSystemObject")

objFso.CreateTextFile pathFrom2


if request.querystring = "update" then


	Set ts2 = objFso.OpenTextFile(pathFrom2, 2)
	Set ts = objFso.OpenTextFile(pathFrom, 1)


	Do While ts.AtEndOfStream <> True
			riga = ts.ReadLine

	if not left(riga,1)="'" AND mid(riga,2,1)<>"%" AND mid(riga,1,1)<>"%" then
	if riga<>"" then
	temp1=right(riga,(len(riga)-instr(riga,"'")+1))
'response.write "temp1 "&temp1&"<br>"
	temp2=left(riga,instr(riga,"=")-1)
'response.write "temp2 "&temp2&"<br>"
	temp3=trim(replace(temp2,"Const ",""))
'response.write "temp3 "&temp3&"<br>"
	strAsgInput = Request.Form(temp3)
	strAsgInput = Replace(strAsgInput, ":", "")
	ts2.writeline(temp2&" = """ & strAsgInput & """ "&temp1 )

	else
	ts2.writeline(riga)

	end if

	else
	ts2.writeline(riga)

	end if

	Loop
	ts.Close
	Set ts = Nothing
	ts2.Close
	Set ts2 = Nothing

objFso.deleteFile pathFrom
objFso.CopyFile pathFrom2, pathFrom
objFso.deleteFile pathFrom2

end if


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
			<td align="center" background="<%= strAsgSknPathImage & strAsgSknTableBarBgImage %>" bgcolor="<%= strAsgSknTableBarBgColour %>" height="20" class="bartitle"><%= UCase(strAsgTxtSkinSettings) %></td>
		  </tr>
		  <tr bgcolor="<%= strAsgSknTableLayoutBorderColour %>">
			<td align="center" height="1"></td>
		  </tr>
		</table><br />
		<form method="post" name="skin" action="settings_skin.asp?update">
		<table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr bgcolor="<%= strAsgSknTableTitleBgColour %>" class="normaltitle">
			<td background="<%= strAsgSknPathImage & strAsgSknTableTitleBgImage %>" align="center" height="16" colspan="2"><%= UCase(strAsgTxtSkinSettings) %></td>
		  </tr>
		<%
		
		Set ts = objFso.OpenTextFile(pathFrom, 1)

		Do While ts.AtEndOfStream <> True
			
			riga = ts.ReadLine
		
		if not left(riga,1)="'" AND mid(riga,2,1)<>"%" AND mid(riga,1,1)<>"%" then
		temp=right(riga,(len(riga)-instr(riga,"=")))
	
		if temp <> "" then
		
		%>	
		  <tr class="smalltext" bgcolor="<%= strAsgSknTableContBgColour %>">
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="25%" align="right"><input type="text" name="<%=replace(trim(left(riga,instr(riga,"=")-1)),"Const ","")%>" value="<%=trim(replace(left(temp,instr(temp,"'")-1),"""",""))%>" />&nbsp;</td>
            <td background="<%= strAsgSknPathImage & strAsgSknTableContBgImage %>" width="75%">&nbsp;<%=replace(right(temp,(len(temp)-instr(temp,"'"))),"""","")%></td>
          </tr>
		<%
		
		end if
	
		end if
	
		Loop
		ts.Close
		Set ts = Nothing
	
		Set objFso = Nothing

		'// Row - End table spacer			
		Call BuildTableContEndSpacer(2)
	
		%>
		  <tr class="smalltext">
            <td colspan="2" align="center"><br /><input type="submit" value="<%= strAsgTxtUpdate %>" name="invia" /></td>
          </tr>
		</table><br />
		</form>
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