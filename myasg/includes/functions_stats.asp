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
' FUNZIONI DI OUTPUT STATISTICO
'-----------------------------------------------------------------------------------------

Dim intStsOreDiOggi			'In uso nelle funzioni
Dim intStsGiorniPerMese
	

'-----------------------------------------------------------------------------------------
' Giorni per Mese
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	26.11.2003 | 26.11.2003
' Commenti:	
'-----------------------------------------------------------------------------------------
function GiorniPerMese(ByVal mese)
	
	'Conta i giorni dei mesi
	Select Case CInt(mese)
		Case 1 
			intStsGiorniPerMese = 31
		Case 2 
			'Controllo Anni Bisestili
			If IsDate("29/02/" & Year(Date())) then
				intStsGiorniPerMese = 29
			Else
				intStsGiorniPerMese = 28
			End if 		
		Case 3 
			intStsGiorniPerMese = 31
		Case 4 
			intStsGiorniPerMese = 30
		Case 5 
			intStsGiorniPerMese = 31
		Case 6 
			intStsGiorniPerMese = 30
		Case 7 
			intStsGiorniPerMese = 31
		Case 8 
			intStsGiorniPerMese = 31
		Case 9 
			intStsGiorniPerMese = 30
		Case 10 
			intStsGiorniPerMese = 31
		Case 11 
			intStsGiorniPerMese = 30
		Case 12 
			intStsGiorniPerMese = 31
	End Select
	
end function


'-----------------------------------------------------------------------------------------
' Media Giornaliera
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	16.11.2003 | 16.11.2003
' Commenti:	
'-----------------------------------------------------------------------------------------
function MediaGiorno(Accessi, Tipo, Cronologia)
	

	If Cronologia = 1 Then		'Oggi
		intStsOreDiOggi = CInt(Hour(dtmAsgNow))
		'Da 00.00 a 00.59 dovrebbe dividere per 0
		'impostare ad 1 dato che Acchiardi insegna che per 0 non si divide! ;oP
		If intStsOreDiOggi = 0 Then intStsOreDiOggi = 1
		
		inttmp = FormatNumber(Accessi/intStsOreDiOggi, 1)
	ElseIf Cronologia = 2 Then	'Ieri
		inttmp = FormatNumber(Accessi/24, 1)
	End If
	
	MediaGiorno = inttmp

end function


'-----------------------------------------------------------------------------------------
' Media Mensile
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	16.11.2003 | 16.11.2003
' Commenti:	
'-----------------------------------------------------------------------------------------
function MediaMese(Accessi, Tipo, Cronologia)
	
	If Cronologia = 1 Then		'Mese Corrente
		
		'Calcolo dei giorni
		GiorniPerMese(dtmAsgMonth)
			
		If Tipo = 1 Then		' x/Ora
			'Calcola ore dei giorni passati
			dtmtmp = 24*(CInt(dtmAsgDay) - 1)
			'Aggiungi le ore di oggi
			dtmtmp = dtmtmp + Hour(dtmAsgNow)
			'NON SI PUO' DIVIDERE PER O!
			'Il bug si presenta la prima ora del primo mese
			If dtmtmp = 0 Then dtmtmp = 1
			'Dividi gli accessi per le ore cacolate
			inttmp = FormatNumber(Accessi/dtmtmp, 1)
		
		ElseIf Tipo = 2 Then	' x/Giorno
			inttmp = FormatNumber(Accessi/CInt(dtmAsgDay), 1)
		End If
		
	ElseIf Cronologia = 2 Then	'Mese Scorso
		
		'Imposta la variabile temporanea
		dtmtmp = CInt(dtmAsgMonth) - 1
		If dtmtmp = 0 Then dtmtmp = 12
		
		'Calcolo dei giorni
		Call GiorniPerMese(dtmtmp)
			
		If Tipo = 1 Then		' x/Ora
			'Ore del mese passato
			dtmtmp = 24*(CInt(intStsGiorniPerMese))
			'Calcola
			inttmp = FormatNumber(Accessi/CInt(dtmtmp), 1)
		ElseIf Tipo = 2 Then	' x/Giorno
			inttmp = FormatNumber(Accessi/CInt(intStsGiorniPerMese), 1)
		End If
		
	End If
	
	MediaMese = inttmp

end function


'-----------------------------------------------------------------------------------------
' Calcola Valore Percentuale
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	24.11.2003 | 24.11.2003
' Commenti:	
'-----------------------------------------------------------------------------------------
function CalcolaPercentuale(ByVal totale, ByVal valore)

	If Clng(totale) = 0 Then 
		CalcolaPercentuale = FormatPercent(0, 2)
	Else
		CalcolaPercentuale = FormatPercent(valore/totale, 2)
	End If

end function


'-----------------------------------------------------------------------------------------
' Dichiarazioni Paginazione Avanzata Risultati
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	19.11.2003 | 19.11.2003
' Commenti:			
'-----------------------------------------------------------------------------------------
Dim page
Dim RecordsPerPage
Dim PaginazioneLoop

function DimPaginazioneAvanzata()
	
	page = Request.QueryString("page")
	RecordsPerPage = 40
		
	If len(page) > 0 And IsNumeric(page) Then
		page = CLng(page)
	Else
		page = 1
	End If
		
end function


'-----------------------------------------------------------------------------------------
' Paginazione Avanzata Risultati
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	19.11.2003 | 19.11.2003
' Commenti:	Tratto dal sito di Mems (www.oscarjsweb.com) - forum HTML.it		
'-----------------------------------------------------------------------------------------
function PaginazioneAvanzata(ByVal linkto, ByVal linktoQS)
	
	Dim MaxPage, Midpage, UltimaPagina, indice, fromindice, toindice
	
	
	' Version 2.1
	'----------------------------------------------
	' From this version variable from function has been
	' changed with small script to help to keep 
	' all querystring info
	'
	Dim QueryStringItem 
	linktoQS = Null
	'
	' Run and build the new querystring checking all
	' items stored at the moment in the querystring
	For Each QueryStringItem in Request.QueryString
		If Not QueryStringItem = "page" Then 
			linktoQS = linktoQS & "&" & QueryStringItem & "=" & Request.QueryString(QueryStringItem)
		End If
	Next

	
	Response.Write (strAsgTxtPage & "&nbsp;" & page & "&nbsp;" & strAsgTxtOf & "&nbsp;" & objAsgRs.pageCount & "<br />" )
	Maxpage=10 
	Midpage=(Maxpage\2)+1
	UltimaPagina=objAsgRs.pageCount
	
	if Maxpage>UltimaPagina then
		for indice = 1 to UltimaPagina
			if CInt(indice)=CInt(page) then
			Response.Write ("[<a href=""" & linkto & "?page=" & indice & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;" & indice & """>" & indice & "</a>] " )
			else
			Response.Write ("[<a href=""" & linkto & "?page=" & indice & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;" & indice & """>" & indice & "</a>] " )
			end if
		next 
	else
		if CInt(Midpage)<CInt(page) then
		
			fromindice=page-Midpage+1
			toindice=page+Midpage-1
			if toindice>UltimaPagina then 
				toindice=UltimaPagina
				fromindice=toindice-Maxpage+1
			end if
		else
			fromindice=1
			toindice=Maxpage
		end if
		
		if fromindice <> 1 then
			Response.Write ("[<a href=""" & linkto & "?page=1" & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;1"">&lt;&lt;</a>] " )
		end if
		for indice = fromindice to toindice
			if CInt(indice)=CInt(page) then
				Response.Write ("[<a href=""" & linkto & "?page=" & indice & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;" & indice & """>" & indice & "</a>] " )
			else
				Response.Write ("[<a href=""" & linkto & "?page=" & indice & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;" & indice & """>" & indice & "</a>] " )
			end if
		next 
		if fromindice<UltimaPagina-Maxpage then
			Response.Write ("[<a href=""" & linkto & "?page=" & UltimaPagina & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;" & UltimaPagina & """>&gt;&gt;</a>] " )
		end if
	end if

end function


'-----------------------------------------------------------------------------------------
' Dichiarazioni Paginazione Avanzata Risultati Dettagli
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	19.11.2003 | 19.11.2003
' Commenti:			
'-----------------------------------------------------------------------------------------
Dim detpage
Dim detRecordsPerPage
Dim detPaginazioneLoop
	
function DimPaginazioneAvanzataDettagli()
	
	detpage = Request.QueryString("detpage")
	detRecordsPerPage = 25
	
	If len(detpage) > 0 And IsNumeric(detpage) Then
		detpage = CLng(detpage)
	Else
		detpage = 1
	End If
	
end function


'-----------------------------------------------------------------------------------------
' Paginazione Avanzata Risultati Dettagli
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	19.11.2003 | 19.11.2003
' Commenti:	Tratto dal sito di Mems (www.oscarjsweb.com) - forum HTML.it		
'-----------------------------------------------------------------------------------------
function PaginazioneAvanzataDettagli(ByVal linkto, ByVal linktoQS)
	
	Dim detMaxPage, detMidpage, detUltimaPagina, detIndice, detfromindice, dettoindice
	
	
	' Version 2.1
	'----------------------------------------------
	' From this version variable from function has been
	' changed with small script to help to keep 
	' all querystring info
	'
	Dim QueryStringItem 
	linktoQS = Null
	'
	' Run and build the new querystring checking all
	' items stored at the moment in the querystring
	For Each QueryStringItem in Request.QueryString
		If Not QueryStringItem = "page" Then 
			linktoQS = linktoQS & "&" & QueryStringItem & "=" & Request.QueryString(QueryStringItem)
		End If
	Next

	
	Response.Write (strAsgTxtPage & "&nbsp;" & detpage & "&nbsp;" & strAsgTxtOf & "&nbsp;" & objAsgRs2.pageCount & "<br />" )
	detMaxpage=10 
	detMidpage=(detMaxpage\2)+1
	detUltimaPagina=objAsgRs2.pageCount
	
	if detMaxpage>detUltimaPagina then
		for detIndice = 1 to detUltimaPagina
			if CInt(detIndice)=CInt(detpage) then
			Response.Write ("[<a href=""" & linkto & "?detpage=" & detIndice & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;" & detIndice & """>" & detIndice & "</a>] " )
			else
			Response.Write ("[<a href=""" & linkto & "?detpage=" & detIndice & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;" & detIndice & """>" & detIndice & "</a>] " )
			end if
		next 
	else
		if CInt(detMidpage)<CInt(detpage) then
		
			detfromindice=detpage-detMidpage+1
			dettoindice=detpage+detMidpage-1
			if dettoindice>detUltimaPagina then 
				dettoindice=detUltimaPagina
				detfromindice=dettoindice-detMaxpage+1
			end if
		else
			detfromindice=1
			dettoindice=detMaxpage
		end if
		
		if detfromindice <> 1 then
			Response.Write ("[<a href=""" & linkto & "?detpage=1" & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;1"">&lt;&lt;</a>] " )
		end if
		for detIndice = detfromindice to dettoindice
			if CInt(detIndice)=CInt(detpage) then
				Response.Write ("[<a href=""" & linkto & "?detpage=" & detIndice & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;" & detIndice & """>" & detIndice & "</a>] " )
			else
				Response.Write ("[<a href=""" & linkto & "?detpage=" & detIndice & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;" & detIndice & """>" & detIndice & "</a>] " )
			end if
		next 
		if detfromindice<detUltimaPagina-detMaxpage then
			Response.Write ("[<a href=""" & linkto & "?detpage=" & detUltimaPagina & linktoQS & """ title=""" & strAsgTxtGoToPage & "&nbsp;" & detUltimaPagina & """>&gt;&gt;</a>] " )
		end if
	end if

end function


'-----------------------------------------------------------------------------------------
' Vai al Mese
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	19.11.2003 | 11.05.2004
' Commenti:	
'-----------------------------------------------------------------------------------------
%><!--#include file="inc_array_month.asp" --><%

function GoToPeriod(ByVal linkto, ByVal linktoQS)

		Dim QueryStringItem

		Response.Write vbCrLf & "<!-- date period panel -->"
		Response.Write vbCrLf & "<table width=""300"" border=""0"" cellspacing=""0"" cellpadding=""0"" height=""30"">"
		Response.Write vbCrLf & "<form action="""& linkto & "?" & linktoQS & """ method=""get"">"
		Response.Write vbCrLf & "  <tr class=""smalltext"" valign=""middle"" align=""left"">"
		Response.Write vbCrLf & "	<td width=""25%"">" & strAsgTxtShow & "</td>"
		Response.Write vbCrLf & "	<td width=""65%"">"

		Response.Write vbCrLf & "	<select name=""periodmm"" class=""smallform"">"
		For intAsgMonthLoop = 1 to Ubound(aryAsgMonth)
		Response.Write vbCrLf & "		<option value=""" & Right("0" & intAsgMonthLoop, 2) & """" 
			If Cint(Left(mese, 2)) = aryAsgMonth(intAsgMonthLoop, 1) Then Response.Write " selected"
		Response.Write " >" & aryAsgMonth(intAsgMonthLoop, 2) & "</option>"
		Next
		Response.Write vbCrLf & "	</select>"

		Response.Write vbCrLf & "	<select name=""periodyy"" class=""smallform"">"
		For looptmp = Year(dtmAsgStartStats) to dtmAsgYear 
		Response.Write vbCrLf & "		<option value=""" & looptmp & """" 
			If CInt(Right(mese, 4)) = CInt(looptmp) Then Response.Write " selected"
		Response.Write " >" & looptmp & "</option>"
		Next
		Response.Write vbCrLf & "	</select>"

		Response.Write vbCrLf & "	</td>"
		Response.Write vbCrLf & "	<td width=""10%"">"
		Response.Write vbCrLf & "	<input type=""submit"" name=""showperiod"" value=""" & strAsgTxtShow & """ class=""smallform"" /></td>"
		Response.Write vbCrLf & "	<input type=""hidden"" name=""elenca"" value=""mese"" class=""smallform"" />"

		For Each QueryStringItem in Request.QueryString
			If Not QueryStringItem = "periodmm" AND Not QueryStringItem = "periodyy" AND Not QueryStringItem = "showperiod" AND Not QueryStringItem = "elenca" Then
			Response.Write vbCrLf & "	<input type=""hidden"" name=""" & QueryStringItem & """ value=""" & Request.QueryString(QueryStringItem) & """ class=""smallform"" />"
			End If
		Next
		
		Response.Write vbCrLf & "	</td>"
		Response.Write vbCrLf & "  </tr>"
		Response.Write vbCrLf & "</form>"
		Response.Write vbCrLf & "</table>"
		Response.Write vbCrLf & "<!-- / date period panel -->"

end function


'-----------------------------------------------------------------------------------------
' Vai a Elenco
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	02.12.2003 | 11.15.2004
' Commenti:	
'-----------------------------------------------------------------------------------------
function GoToGrouping(ByVal linkto, ByVal linktoQS)

		Dim QueryStringItem

		Response.Write vbCrLf & "<!-- grouping panel -->"
		Response.Write vbCrLf & "<table width=""300"" border=""0"" cellspacing=""0"" cellpadding=""0"" height=""30"">"
		Response.Write vbCrLf & "<form action="""& linkto & "?" & linktoQS & """ method=""get"">"
		Response.Write vbCrLf & "  <tr class=""smalltext"" valign=""middle"" align=""left"">"
		Response.Write vbCrLf & "	<td width=""25%"">" & strAsgTxtShow & "</td>"
		Response.Write vbCrLf & "	<td width=""65%"">"
		Response.Write vbCrLf & "	<select name=""elenca"" class=""smallform"">"
		Response.Write vbCrLf & "		<option value=""mese"""
			If elenca = "mese" Then Response.Write "selected"
		Response.Write " >" & strAsgTxtByMonth & "</option>"
		Response.Write vbCrLf & "		<option value=""tutti"""
			If elenca = "tutti" Then Response.Write "selected"
		Response.Write " >" & strAsgTxtAll & "</option>"
		Response.Write vbCrLf & "	</select>"
		Response.Write vbCrLf & "	</td>"
		Response.Write vbCrLf & "	<td width=""10%"">"
		Response.Write vbCrLf & "	<input type=""submit"" name=""showlist"" value=""" & strAsgTxtShow & """ class=""smallform"" /></td>"
		
		For Each QueryStringItem in Request.QueryString
			If Not QueryStringItem = "elenca" AND Not QueryStringItem = "showlist" Then
			Response.Write vbCrLf & "	<input type=""hidden"" name=""" & QueryStringItem & """ value=""" & Request.QueryString(QueryStringItem) & """ class=""smallform"" />"
			End If
		Next
		
		Response.Write vbCrLf & "	</td>"
		Response.Write vbCrLf & "  </tr>"
		Response.Write vbCrLf & "</form>"
		Response.Write vbCrLf & "</table>"
		Response.Write vbCrLf & "<!-- / grouping panel -->"

end function


'-----------------------------------------------------------------------------------------
' Search for Data
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	11.15.2004
' Commenti:	
'-----------------------------------------------------------------------------------------
function SearchForData(ByVal linkto, ByVal linktoQS, ByVal dbField)

		Dim aryAsgDbField		'Holds the array of the database fields to search in
		Dim QueryStringItem		'Holds the querystring collection
		
		'Clean variable
		looptmp = Null
		'Split the variable to an array containing all database fields to search in
		aryAsgDbField = Split(dbField, "|")		

		Response.Write(vbCrLf & "<!-- search engine panel -->")
		Response.Write(vbCrLf & "<table width=""300"" border=""0"" cellspacing=""0"" cellpadding=""0"" height=""30"">")
		Response.Write(vbCrLf & "<form action="""& linkto & "?" & linktoQS & """ method=""get"">")
		Response.Write(vbCrLf & "  <tr class=""smalltext"" valign=""middle"" align=""left"">")
		Response.Write(vbCrLf & "	<td width=""25%"">" & strAsgTxtSearch & "</td>")
		Response.Write(vbCrLf & "	<td width=""65%"">")
		Response.Write(vbCrLf & "	<input type=""text"" name=""searchfor"" value=""" & Request.QueryString("searchfor") & """ class=""smallform"" size=""10"" />")
		Response.Write(vbCrLf & "	<select name=""searchin"" class=""smallform"">")
		For looptmp = 0 to Ubound(aryAsgDbField)
		Response.Write(vbCrLf & "		<option value=""" & aryAsgDbField(looptmp) & """")
			If Request.QueryString("searchin") = aryAsgDbField(looptmp) Then Response.Write("selected")
		Response.Write(" >" & aryAsgDbField(looptmp) & "</option>")
		Next
		Response.Write(vbCrLf & "	</select>")
		Response.Write(vbCrLf & "	</td>")
		Response.Write(vbCrLf & "	<td width=""10%"">")
		Response.Write(vbCrLf & "	<input type=""submit"" name=""showsearch"" value=""" & strAsgTxtShow & """ class=""smallform"" /></td>")
		
		For Each QueryStringItem in Request.QueryString
			If Not QueryStringItem = "searchfor" AND Not QueryStringItem = "searchin" AND Not QueryStringItem = "showsearch" Then
			Response.Write(vbCrLf & "	<input type=""hidden"" name=""" & QueryStringItem & """ value=""" & Request.QueryString(QueryStringItem) & """ class=""smallform"" />")
			End If
		Next
		
		Response.Write(vbCrLf & "	</td>")
		Response.Write(vbCrLf & "  </tr>")
		Response.Write(vbCrLf & "</form>")
		Response.Write(vbCrLf & "</table>")
		Response.Write(vbCrLf & "<!-- / search engine panel -->")

end function


'-----------------------------------------------------------------------------------------
' Check search for Data
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	11.15.2004
' Commenti:	
'-----------------------------------------------------------------------------------------
function CheckSearchForData(ByVal sqlstring, firstcondition)
	
	'Read for keywords to search for
	asgSearchfor = Trim(Request.QueryString("searchfor"))
	
	'Read for a field to search in
	asgSearchin = Trim(Request.QueryString("searchin"))

	'If there are keywords to search for and a field to search in then add SQL search string
	If Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 Then
		
		'If this isn't the first WHERE condition add the string using AND operator
		If firstcondition = False Then
			sqlstring = sqlstring & " AND " & asgSearchin & " LIKE '%" & FilterSQLInput(asgSearchfor) & "%' "
		'If this is the first WHERE condition add the string using WHERE operator
		Else
			sqlstring = sqlstring & " WHERE " & asgSearchin & " LIKE '%" & FilterSQLInput(asgSearchfor) & "%' "
		End If

	'If there are no enough information to query the database then return the normal SQL query	
	Else

		sqlstring = sqlstring

	End If

	'Return function
	CheckSearchForData = sqlstring
	
end function


'-----------------------------------------------------------------------------------------
' Check search for Data
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	11.15.2004
' Commenti:	
'-----------------------------------------------------------------------------------------
function HighlightSearchKey(ByVal contentstring, ByVal databasefield)

	'If some data has been searched and this is the database 
	'where you have searched in then highlight search terms
	If Len(asgSearchfor) > 0 AND Len(asgSearchin) > 0 AND asgSearchin = databasefield Then
	
		contentstring = Replace(contentstring, asgSearchfor, "<span class=""highlighted"">" & asgSearchfor & "</span>", 1, -1, vbTextCompare)

	End If
	
	'Return function
	HighlightSearchKey = contentstring

end function


'-----------------------------------------------------------------------------------------
' Taglia le stringhe lunghe
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	10.03.2004 | 
' Commenti:	
'-----------------------------------------------------------------------------------------
function StripValueTooLong(ByVal stringToStrip, ByVal stringMaxLenght, ByVal stripLeft, ByVal stripRight)

	If Len(stringToStrip) > stringMaxLenght Then 
		stringToStrip = Left(stringToStrip, stripLeft) & "..." & Right(stringToStrip, stripRight)
	Else
		stringToStrip = stringToStrip
	End If

	StripValueTooLong = stringToStrip

end function


'-----------------------------------------------------------------------------------------
' Icona di competenza dominio
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	12.03.2004 | 
' Commenti:			
'-----------------------------------------------------------------------------------------
function ChooseDomainIcon(ByVal outputPage, ByVal prefixType)

	strtmp = Null
	strtmp = outputPage

If prefixType = "classic" Then

	'Taglia tutto il prefisso sito + http:// se non è una pagina sconosciuta
	'Versioni precedenti 1.2
	'If Mid(asgOutputPage, 1, Len("http://" & strAsgSiteURLremote)) = "http://" & strAsgSiteURLremote Then asgOutputPage = Mid(asgOutputPage, Len("http://" & strAsgSiteURLremote) + 1) 
	If Mid(strtmp, 1, Len(strAsgSiteURLremote)) = strAsgSiteURLremote Then strtmp = Mid(strtmp, Len(strAsgSiteURLremote) + 1) 

ElseIf prefixType = "visitors" Then

	'Taglia tutto il prefisso sito + http:// se non è una pagina sconosciuta
	'NB. La formula originale prevedeva
	'	'Taglia tutto il prefisso sito + http:// se non è una pagina sconosciuta
		'If Mid(asgOutputPage, 1, Len("http://" & strAsgSiteURLremote)) = "http://" & strAsgSiteURLremote Then asgOutputPage = Mid(asgOutputPage, Len("http://" & strAsgSiteURLremote) + 1) 
	If Mid(strtmp, 1, Len("http://" & strAsgSiteURLremote)) = "http://" & strAsgSiteURLremote Then strtmp = Mid(strtmp, Len("http://" & strAsgSiteURLremote) + 1) 

End If
	
	'Mostra una icona appropriata in base alla corrispondenza
	If outputPage <> strtmp Then
		ChooseDomainIcon = vbCrLf & "<img src=""images/home.gif"" alt=""" & strAsgSiteURLremote & """ align=""absmiddle"" border=""0"" />"
	Else
		ChooseDomainIcon = vbCrLf & "<img src=""images/arrow_small_dx.gif"" alt=""" & outputPage & """ align=""absmiddle"" border=""0"" />"
	End If
	
	'Ritorna il valore della variabile asgOutputPage
	'per consentire l'uso della successiva funzione di taglio stringa
	asgOutputPage = strtmp

end function


'-----------------------------------------------------------------------------------------
' Manda a capo le stringhe troppo lunghe
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	15.03.2004
' Commenti:	Tratto dal sito di Mems (www.oscarjsweb.com) - forum HTML.it		
'-----------------------------------------------------------------------------------------
function ShareWords(tempTXT, maxlenght)
	
	Dim Limit, arrTxt, looptmp2, tempLenght, start
	
	Limit = maxlenght
	arrTXT = Split(tempTXT)
	
	For looptmp = 0 To UBound(arrTXT)
	
	tempLenght = Len(arrTXT(looptmp))
	If tempLenght > Limit Then
	inttmp = tempLenght / Limit
	
	If inttmp - CInt(inttmp) <> 0 Then
	inttmp = inttmp + 1
	End If
	
	start = 1
	
	For looptmp2 = 1 To inttmp
	Response.Write Mid(arrTXT(looptmp),start,Limit) & " "
	start = start + Limit
	Next
	Else
	Response.Write arrTXT(looptmp) & " "
	End If
	
	Next
	
end function


'-----------------------------------------------------------------------------------------
' Richiamo ultima versione da sito
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	30.03.2004 | 
' Commenti:	Thanks to ToroSeduto
'-----------------------------------------------------------------------------------------
function GetLastVersion(ByVal siteUrl)
	
	Dim objXMLHTTP
	
	Set objXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")
	objXMLHTTP.Open "GET", siteUrl, false
	objXMLHTTP.Send     
	GetLastVersion = CStr(objXMLHTTP.ResponseText)
	Set objXMLHTTP = Nothing 
	
end function


'-----------------------------------------------------------------------------------------
' Controllo nuove versioni
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	30.03.2004 | 
' Commenti:	
'-----------------------------------------------------------------------------------------
function CheckUpdate(ByVal asgVersion, ByVal asgUpdate)
	
	'Dim strAsgLastVersion			'Ultima Versione dal sito
	Dim aryAsgLastVersion			'Array con info ultima versione
	
	strAsgLastVersion = GetLastVersion("http://www.weppos.com/asg/checkversion/check_update.asp?host=" & Server.URLEncode(Request.ServerVariables("HTTP_HOST")))
	aryAsgLastVersion = Split(strAsgLastVersion, "|")

	strAsgLastVersion = aryAsgLastVersion(0)
	dtmAsgLastUpdate = aryAsgLastVersion(1)
	urlAsgLastUpdate = aryAsgLastVersion(2)
	
	If LCase(asgVersion) <> LCase(strAsgLastVersion) then 	
		
		'Versioni differenti
		intAsgLastUpdate = 2
	
	Else
		
		If Clng(asgUpdate) <> Clng(dtmAsgLastUpdate) then 	
			'Versioni uguali ma differente aggiornamento
			intAsgLastUpdate = 3
		Else
			'Tutto combacia
			intAsgLastUpdate = 1
		End If
	
	End If

end function


'-----------------------------------------------------------------------------------------
' Adatta il fuso orario outbound
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	31.03.2004 |
' Commenti:	
'-----------------------------------------------------------------------------------------
function FormatOutTimeZone(ByVal datetimeValue, ByVal datetimeType)
	
	Dim outTimeZone 
	outTimeZone = aryAsgTimeZone(1)
	
	If IsNull(datetimeValue) OR Not Len(datetimeValue) > 0 Then datetimeValue = dtmAsgNow
	
	'Modifica l'orario in base al fuso
	If Left(outTimeZone, 1) = "+" Then
		datetimeValue = DateAdd("h", outTimeZone, datetimeValue)
	ElseIf Left(outTimeZone, 1) = "-" Then
		datetimeValue = DateAdd("h", outTimeZone, datetimeValue)
	Else
		'
	End If
	
	'Verifica il tipo di periodo richiesto
	Select Case datetimeType
		Case "Year"
			dtmtmp = CInt(Year(datetimeValue))
		Case "Month"
			dtmtmp = CInt(Month(datetimeValue))
		Case "Day"
			dtmtmp = CInt(Day(datetimeValue))
		Case "Hour"
			dtmtmp = CInt(Hour(datetimeValue))
		Case "Minute"
			dtmtmp = CInt(Minute(datetimeValue))
		Case "Second"
			dtmtmp = CInt(Second(datetimeValue))
		Case "Time"
			dtmtmp = CDate(TimeSerial(Hour(datetimeValue), Minute(datetimeValue), Second(datetimeValue)))
		Case "Date"
			dtmtmp = CDate(DateSerial(Year(datetimeValue), Month(datetimeValue), Day(datetimeValue)))
		
		'Month value for stats report
		Case "MonthReport"
			dtmtmp = Right("0" & Month(datetimeValue) & "-" & Year(datetimeValue), 7)

	End Select
	
	If Not datetimeType = "Time" AND Not datetimeType = "Date" AND Not datetimeType = "MonthReport" Then
		If datetimeValue < 10 Then datetimeValue = "0" & datetimeValue 
	End If
	
	FormatOutTimeZone = dtmtmp
		
end function


%>