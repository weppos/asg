<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--include virtual="/myasg/config.asp" -->
<!--#include file="w2k3_config.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'// WARNING! Program protection.
'	Changing default values may allow users to access the page.
' -->	Sostituito da Avvertimento On Screen! 
'	Call checkPermission("False", "False", "False", appAsgSecurity)

Dim strAsgSelectedIP	'IP Passato in QueryString
Dim asgOutputPage
Dim strNewFilteredIP	'IP da Filtrare
Dim strCommand			'Comando da Eseguire sull'IP
Dim blnUpdateCompleted	'TRUE se completato

'Richiama informazioni
strAsgSelectedIP = Trim(Request.QueryString("IP"))
strNewFilteredIP = Trim(Request.Form("filterIP"))
blnUpdateCompleted = False

'Verifica per Inserimento IP nel Filtro
If Request.Form("submit") = TXT_Update AND Len(strNewFilteredIP) > 0 AND Session("asgLogin") = "Logged" Then

	strCommand = Request.Form("command")
	
	'Resetta ed Aggiungi
	If strCommand = "reset" Then
		
		'Aggiornamento
		strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "config SET conf_filtered_ips = '" & strNewFilteredIP & "'"
		objAsgConn.Execute(strAsgSQL)
		
		'Imposta a TRUE l'aggiornamento
		blnUpdateCompleted = True
	
	'Aggiungi alla lista
	ElseIf strCommand = "add" Then

		'Richiama le informazioni sull'IP anche se in memoria
		'ma ci sarebbero troppi controlli da fare!
		If ASG_USE_MYSQL then
			strAsgSQL = "SELECT conf_filtered_ips FROM " & ASG_TABLE_PREFIX & "config LIMIT 1"
		else
			strAsgSQL = "SELECT TOP 1 conf_filtered_ips FROM " & ASG_TABLE_PREFIX & "config"
		end if
		'Open Rs
		objAsgRs.Open strAsgSQL, objAsgConn
		
		if not objAsgRs.EOF then
			
			'Rivalorizza Variabile
			appAsgFilteredIPs = Trim(objAsgRs("conf_filtered_ips"))
			'Pulisci spazi
			appAsgFilteredIPs = Replace(appAsgFilteredIPs, " ", "")
			
			'Controlla presenza " , " finale
			If Right(appAsgFilteredIPs, 1) = "," Then
				strNewFilteredIP = appAsgFilteredIPs & strNewFilteredIP
			'In mancanza aggiungi
			Else
				strNewFilteredIP = appAsgFilteredIPs & "," & strNewFilteredIP
			End If
			
		End If
		
		objAsgRs.Close
			
		'Aggiornamento
		strAsgSQL = "UPDATE " & ASG_TABLE_PREFIX & "config SET conf_filtered_ips = '" & strNewFilteredIP & "'"
		objAsgConn.Execute(strAsgSQL)
		
		'Imposta a TRUE l'aggiornamento
		blnUpdateCompleted = True
	
	End If

	
	'Se si utilizzano le variabili Application aggiornale
	If blnApplicationConfig Then
						
		'Aggiorna Variabili Application
		Application(ASG_APPLICATION_PREFIX & "FilteredIPs") = strNewFilteredIP
		'Forza il ricalcolo delle Application
		Application(ASG_APPLICATION_PREFIX & "Config") = False
	
	End If

	
End If


' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing


%>
<%= STR_ASG_PAGE_DOCTYPE %>
<html>
<head>
<title><%= TXT_Filtered_IPs & "&nbsp;" & TXT_For & "&nbsp;" & strAsgSelectedIP %> | ASP Stats Generator <%= ASG_VERSION %></title>
<%= STR_ASG_PAGE_CHARSET %>
<meta name="copyright" content="Copyright (C) 2003-2005 Carletti Simone" />
<!--#include file="includes/meta.inc.asp" -->

<!-- ASP Stats Generator v. <%= ASG_VERSION %> is created and developed by Simone Carletti.
To download your Free copy visit the official site http://www.weppos.com/asg/ -->

</head>

<%

'
Response.Write(vbCrLf & "<body bgcolor=""" & STR_ASG_SKIN_PAGE_BGCOLOUR & """ background=""" & STR_ASG_SKIN_PAGE_BGIMAGE & """>")

' TableBar			
Call buildTableBar(TXT_Filtered_IPs, MENUGROUP_VisitorProfiles)
	
' 
Response.Write(vbCrLf & "<div class=""table_layout"">")

'CONTENUTO
'---------------------------------------------------|
Response.Write(vbCrLf & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"">")

' Show only if the user is logged in
If Session("asgLogin") = "Logged" Then
		
		'CONTENUTO AGGIORNAMENTO
		'---------------------------------------------------|
	
	'Aggiornato
	If blnUpdateCompleted Then
		
		Response.Write(vbCrLf & "<tr class=""table_cont_row"">")
		Response.Write(vbCrLf & "<td align=""center"" colspan=""2""><p>" & TXT_Update_Completed & "</p></td>")
		Response.Write(vbCrLf & "</tr>")
	
	Else	
		
		'CONTENUTO
		'---------------------------------------------------|
	
	'Manca l'IP in QueryString
	If NOT Len(strAsgSelectedIP) > 0 Then 
		
		Response.Write(vbCrLf & "<tr class=""table_cont_row"">")
		Response.Write(vbCrLf & "<td align=""center"" colspan=""2""><p>" & TXT_MissedDataToElab & "</p></td>")
		Response.Write(vbCrLf & "</tr>")
	
	'IP passato correttamente	
	Else
		
		Response.Write(vbCrLf & "<form name=""frmFilterIp"" action=""popup_filter_ip.asp?IP=" & strAsgSelectedIP & """ method=""post"">")
		
		'Form IP
		Response.Write(vbCrLf & "<tr ")
		Response.Write(buildTableContRollover("table_cont_row"))
		Response.Write(">")
		Response.Write(vbCrLf & "<td align=""right"" width=""25%"">" & TXT_IPAddress & "&nbsp;:&nbsp;&nbsp;</td>")
		Response.Write(vbCrLf & "<td align=""left""  width=""75%"">&nbsp;<input type=""text"" size=""25"" maxlenght=""20"" name=""filterIP"" value=""" & strAsgSelectedIP & """ /></td>")
		Response.Write(vbCrLf & "</tr>")
		
		'Info RANGE
		Response.Write(vbCrLf & "<tr class=""table_cont_row"">")
		Response.Write(vbCrLf & "<td align=""left"" colspan=""2""><p>" & TXT_InformationsToExitByIpRange & "</p></td>")
		Response.Write(vbCrLf & "</tr>")
		
		'Azione
		Response.Write(vbCrLf & "<tr ")
		Response.Write(buildTableContRollover("table_cont_row"))
		Response.Write(">")
		Response.Write(vbCrLf & "<td align=""right"" height=""16"">" & TXT_Action & "&nbsp;:&nbsp;&nbsp;</td>")
		Response.Write(vbCrLf & "<td align=""left"">&nbsp;<select name=""command"">")
		Response.Write(vbCrLf & "<option value=""add"">" & TXT_AddToList &"</option>")
		Response.Write(vbCrLf & "<option value=""reset"">" & TXT_ResetAndAddToList &"</option>")
		Response.Write(vbCrLf & "</select>&nbsp;&nbsp;&nbsp;")
		Response.Write(vbCrLf & "<input type=""submit"" name=""submit"" value=""" & TXT_Update & """ />")
		Response.Write(vbCrLf & "</tr>")
		
		Response.Write(vbCrLf & "</form>")
	
	End If

	'Fine condizione Aggiornato
	End If	


' Show an insufficient permission advice
Else
	
	Response.Write(vbCrLf & "<tr class=""table_cont_no_record"">")
	Response.Write(vbCrLf & "<td align=""center""><p>" & TXT_InsufficientPermission & "</p></td>")
	Response.Write(vbCrLf & "</tr>")
	
End If

'CONTENUTO (Chiusura)
'---------------------------------------------------|
Response.Write(vbCrLf & "</table>")



' 
Response.Write(vbCrLf & "</div>")

' Footer
' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write(vbCrLf & "<div class=""table_footerbar"">Powered by <a href=""http://www.weppos.com/asg/"" title=""ASP Stats Generator"">ASP Stats Generator</a> v" & ASG_VERSION & " &copy; 2003-2004 <a href=""http://www.weppos.com/"">weppos</a>")
If ASG_ELABORATION_TIME Then Response.Write(" - " & TXT_ThisPageWasGeneratedIn & "&nbsp;" & FormatNumber(Timer() - startAsgElab, 4) & "&nbsp;" & TXT_seconds)
Response.Write("</div>")
Response.Write(vbCrLf & "<br /><div class=""footer"" align=""center"">Powered by <a href=""http://www.weppos.com/asg/"" title=""ASP Stats Generator"">ASP Stats Generator</a> v" & ASG_VERSION & " <br />Copyright &copy; 2003-2005 <a href=""http://www.weppos.com/"">weppos</a><div>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******

Response.Write(vbCrLf & "</body></html>")

%>