<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'

			

'-----------------------------------------------------------------------------------------
' Icona Lingua	
'-----------------------------------------------------------------------------------------
' Function:	Restituisce una icona in base al nome della lingua
' Date: 	19.11.2003 | 25.02.2004
' Comment:	Da Classe 2.0.4 versione restituita in 
'			[code] - Lingua Estesa
'-----------------------------------------------------------------------------------------
Function ShowIconLanguage(languages)

	'// Classe 2.0.4
	
	'Italiano
	If InStr(1, languages, "Italiano", 1) > 0  Then
		Response.Write "it.png"
	
	'Inglese
	ElseIf InStr(1, languages, "Inglese", 1) > 0  Then
		Response.Write "gb.png"
	
	'Tedesco
	ElseIf InStr(1, languages, "Tedesco", 1) > 0  Then
		Response.Write "de.png"
	
	'Spagnolo
	ElseIf InStr(1, languages, "Spagnolo", 1) > 0  Then
		Response.Write "es.png"
	
	'Irlandese
	ElseIf InStr(1, languages, "Irlandese", 1) > 0  Then
		Response.Write "ir.png"
	
	'Russo
	ElseIf InStr(1, languages, "Russo", 1) > 0  Then
		Response.Write "ru.png"
	
	'Giapponese
	ElseIf InStr(1, languages, "Giapponese", 1) > 0  Then
		Response.Write "jp.png"

	'// Classe 2.0.4c
	
	'Olandese
	ElseIf InStr(1, languages, "Olandese", 1) > 0  Then
		Response.Write "nl.png"
	
	'Francese
	ElseIf InStr(1, languages, "Francese", 1) > 0  Then
		Response.Write "fr.png"

	'// Classe 2.0.4d
	
	'Portoghese
	ElseIf InStr(1, languages, "Portoghese", 1) > 0  Then
		Response.Write "pt.png"
	
	'Coreano
	ElseIf InStr(1, languages, "Coreano", 1) > 0  Then
		Response.Write "kr.png"

	'// Classe 2.1
	
	'Norvegese
	ElseIf InStr(1, languages, "Norvegese", 1) > 0  Then
		Response.Write "no.png"
	
	'Rumeno
	ElseIf InStr(1, languages, "Rumeno", 1) > 0  Then
		Response.Write "ro.png"
	
	'Danese
	ElseIf InStr(1, languages, "Danese", 1) > 0  Then
		Response.Write "dk.png"
	
	'Svedese
	ElseIf InStr(1, languages, "Svedese", 1) > 0  Then
		Response.Write "se.png"
	
	'Cinese
	ElseIf InStr(1, languages, "Cinese", 1) > 0  Then
		Response.Write "cn.png"
	
	'Ebreo
	ElseIf InStr(1, languages, "Ebreo", 1) > 0  Then
		Response.Write "il.png"
	
	'Turco
	ElseIf InStr(1, languages, "Turco", 1) > 0  Then
		Response.Write "tr.png"

	'// Classe 3.x
	
	'Polacco
	ElseIf InStr(1, languages, "Polacco", 1) > 0  Then
		Response.Write "pl.png"
	
	'Sloveno
	ElseIf InStr(1, languages, "Sloveno", 1) > 0  Then
		Response.Write "sk.png"
	
	'Ceco
	ElseIf InStr(1, languages, "Ceco", 1) > 0  Then
		Response.Write "cz.png"
	
	'Finlandese
	ElseIf InStr(1, languages, "Finlandese", 1) > 0  Then
		Response.Write "fi.png"
	
	'Croato
	ElseIf InStr(1, languages, "Croato", 1) > 0  Then
		Response.Write "hr.png"
	
	'Bulgaro
	ElseIf InStr(1, languages, "Bulgaro", 1) > 0  Then
		Response.Write "bg.png"
	
	'Arabo
	ElseIf InStr(1, languages, "Arabo", 1) > 0  Then
		Response.Write "sa.png"
	
	'Indiano
	ElseIf InStr(1, languages, "Indiano", 1) > 0  Then
		Response.Write "in.png"
	
	'Ungherese
	ElseIf InStr(1, languages, "Ungherese", 1) > 0  Then
		Response.Write "hu.png"
	
	'Greco
	ElseIf InStr(1, languages, "Greco", 1) > 0  Then
		Response.Write "gr.png"
	
	'Lituano
	ElseIf InStr(1, languages, "Lituano", 1) > 0  Then
		Response.Write "lt.png"
	'Slovacco
	ElseIf InStr(1, languages, "Slovacco", 1) > 0  Then
		Response.Write "sk.png"
	
	'Lithuanian
	ElseIf InStr(1, languages, "Lettone", 1) > 0  Then
		Response.Write "lt.png"
	'Ukrainian
	ElseIf InStr(1, languages, "Ucraino", 1) > 0  Then
		Response.Write "ua.png"
	'Estonian
	ElseIf InStr(1, languages, "Estone", 1) > 0  Then
		Response.Write "ee.png"
	'Turkish Origin
	ElseIf InStr(1, languages, "Turca", 1) > 0  Then
		Response.Write "tr.png"
	

	'Mostra Sconosciuto
	Else
		Response.Write "unknown.png"
	
	End If

End Function


'-----------------------------------------------------------------------------------------
' Icona Filtro indirizzo
'-----------------------------------------------------------------------------------------
' Function:	
' Date: 	06.04.2004
' Comment:	
'-----------------------------------------------------------------------------------------
Function ShowIconFilterIp(ByVal ipaddress)
					
	'Filter IP
	'// Link PopUp
	Response.Write(vbCrLf & "<a href=""JavaScript:openWin('popup_filter_ip.asp?IP=" & ipaddress & "','Filter','toolbar=0,location=0,status=0,menubar=0,scrollbars=1,resizable=1,width=550,height=250')"" title=""" & TXT_Filtered_IPs & """>")
								
	'// L'IP è escluso
	If InStr(1, appAsgFilteredIPs, ipaddress, 1) > 0 Then
										
		'// Icona esclusione
		Response.Write(vbCrLf & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "locked_icon.gif"" alt=""" &  TXT_Filtered_IPs & """ border=""0"" align=""absmiddle"" />")
									
	'// L'IP è escluso
	Else
									
		'// Icona esclusione
		Response.Write(vbCrLf & "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "unlocked_icon.gif"" alt=""" &  TXT_Filtered_IPs & """ border=""0"" align=""absmiddle"" />")
								
	End If
								
	'// Chiudi Link PopUp
	Response.Write("</a>")
	
End Function


'-----------------------------------------------------------------------------------------
' Shows the icon to browse the details of the selected value.
' If detail matches with the current detail it will show an "open details" icon.
'
' @since 3.0
'-----------------------------------------------------------------------------------------
public function showIconDetails(argDetails, argCurrentDetails, argTitle)

	Dim lvLayout

	if Len(argDetails) > 0 AND argCurrentDetails = argDetails then
		lvLayout = "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/deton.png"" alt=""" & argTitle & """ border=""0"" align=""middle"" />"
	else
		lvLayout = "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/detoff.png"" alt=""" & argTitle & """ border=""0"" align=""middle"" />"
	end if
	
	' return function 
	showIconDetails = lvLayout

end function


'-----------------------------------------------------------------------------------------
'
' @since 3.0
'-----------------------------------------------------------------------------------------
public function showIconTipUnknows(argValue)

	Dim lvLayout

	if argValue = ASG_UNKNOWN then
		lvLayout = "<img src=""" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png"" alt=""" & TXT_Info & """ border=""0"" align=""middle"" onmouseover=""stm(Info[1],Style[1])"" onmouseout=""htm()"" />&nbsp;"
	else
		lvLayout = ""
	end if
	
	' return function 
	showIconTipUnknows = lvLayout

end function


%>