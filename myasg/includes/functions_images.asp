<%

'/**
' * ASP Stats Generator - Powerful and reliable ASP website counter
' *
' * This file is part of the ASP Stats Generator package.
' * (c) 2003-2007 Simone Carletti <weppos@weppos.net>, All Rights Reserved
' *
' * 
' * COPYRIGHT AND LICENSE NOTICE
' *
' * The License allows you to download, install and use one or more free copies of this program 
' * for private, public or commercial use.
' * 
' * You may not sell, repackage, redistribute or modify any part of the code or application, 
' * or represent it as being your own work without written permission from the author.
' * You can however modify source code (at your own risk) to adapt it to your specific needs 
' * or to integrate it into your site. 
' *
' * All links and information about the copyright MUST remain unchanged; 
' * you can modify or remove them only if expressly permitted.
' * In particular the license allows you to change the application logo with a personal one, 
' * but it's absolutly denied to remove copyright information,
' * including, but not limited to, footer credits, inline credits metadata and HTML credits comments.
' *
' * For the full copyright and license information, please view the LICENSE.htm
' * file that was distributed with this source code.
' *
' * Removal or modification of this copyright notice will violate the license contract.
' *
' *
' * @category        ASP Stats Generator
' * @package         ASP Stats Generator
' * @author          Simone Carletti <weppos@weppos.net>
' * @copyright       2003-2007 Simone Carletti, All Rights Reserved
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id$
' */
 
'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */
			

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