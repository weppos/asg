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

			
			'********** 		NOTE DI SVILUPPO 		*********
			' Inizio conversione nomi e variabili in inglese	'
			'****************************************************
			

'-----------------------------------------------------------------------------------------
' Costruisci Riga Tabella Contenuti - Nessun Record
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	10.03.2004 | 
' Commenti:	
'-----------------------------------------------------------------------------------------
function BuildTableContNoRecord(ByVal colspanValue, ByVal message)
	
	'Verifica se è presente un messaggio alternativo.
	'Nel caso non sia definito usa quello
	'standard.
	If message = "standard" Then 
		message = strAsgTxtNoRecordInDatabase
	ElseIf message = "search" Then
		message = strAsgTxtSearchFoundNoResults
	End If 
			
	Response.Write(vbCrLf & "<tr class=""smalltext"" bgcolor=""" & strAsgSknTableContBgColour & """>")
	Response.Write(vbCrLf & "  <td colspan=""" & colspanValue & """ background=""" & strAsgSknPathImage & strAsgSknTableContBgImage & """ align=""center"">" & message & "</td>")
	Response.Write(vbCrLf & "</tr>")

end function
			

'-----------------------------------------------------------------------------------------
' Costruisci Riga Tabella Contenuti - Debug automatico icone non riconosciute
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	14.03.2004 | 
' Commenti:	
'-----------------------------------------------------------------------------------------
function BuildTableContCheckIcon(ByVal colspanValue, ByVal iconType, ByVal pageNum)
	
	Dim strAsgTableContent
	strAsgTableContent = ""
	
	strAsgTableContent = strAsgTableContent & vbCrLf & "<!-- Informazioni icone non riconosciute -->"
	strAsgTableContent = strAsgTableContent & vbCrLf & "<tr class=""smalltext"" align=""center"" valign=""top"">"
	strAsgTableContent = strAsgTableContent & vbCrLf & "  <td width=""100%"" colspan=""" & colspanValue & """><br /><img src=""" & strAsgSknPathImage & iconType & ".asp?icon=checkicon&page=" & pageNum & """ alt="""" /><br /></td>"
	strAsgTableContent = strAsgTableContent & vbCrLf & "</tr>"
	strAsgTableContent = strAsgTableContent & vbCrLf & "<!-- Informazioni icone non riconosciute -->"
			  
			
	If iconType = "browser" AND Session("blnAsgIconBrowser" & pageNum) <> "notified" AND blnAsgCheckIcon Then
	
		Response.Write(strAsgTableContent)
	
	ElseIf iconType = "os" AND Session("blnAsgIconOs" & pageNum) <> "notified" AND blnAsgCheckIcon Then
	
		Response.Write(strAsgTableContent)
	
	ElseIf iconType = "engine" AND Session("blnAsgIconEngine" & pageNum) <> "notified" AND blnAsgCheckIcon Then
	
		Response.Write(strAsgTableContent)
	
	End If
			
end function
			

'-----------------------------------------------------------------------------------------
' Costruisci Riga Tabella Contenuti - Spaziatore finale
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	14.03.2004 | 
' Commenti:	
'-----------------------------------------------------------------------------------------
function BuildTableContEndSpacer(ByVal colspanValue)

	Response.Write(vbCrLf & "<tr class=""smalltext"" bgcolor=""" & strAsgSknTableTitleBgColour & """>")
	Response.Write(vbCrLf & "  <td colspan=""" & colspanValue & """ background=""" & strAsgSknPathImage & strAsgSknTableTitleBgImage & """ height=""2""></td>")
	Response.Write(vbCrLf & "</tr>")

end function
			

'-----------------------------------------------------------------------------------------
' Costruisci Riga Footer - Linea Bordo
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	10.05.2004 | 
' Commenti:	
'-----------------------------------------------------------------------------------------
function BuildFooterBorderLine()

	Response.Write(vbCrLf & "<tr bgcolor=""" & strAsgSknTableLayoutBorderColour & """>")
	Response.Write(vbCrLf & "  <td align=""center"" height=""1""></td>")
	Response.Write(vbCrLf & "</tr>")

end function
			

%>