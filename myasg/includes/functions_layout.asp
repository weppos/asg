<%

' 
' = ASP Stats Generator - Powerful and reliable ASP website counter
' 
' Copyright (c) 2003-2008 Simone Carletti <weppos@weppos.net>
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
' 
' 
' @category        ASP Stats Generator
' @package         ASP Stats Generator
' @author          Simone Carletti <weppos@weppos.net>
' @copyright       2003-2008 Simone Carletti
' @license         http://www.opensource.org/licenses/mit-license.php
' @version         SVN: $Id$
' 


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