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
' FUNZIONI DI ELABORAZIONE	
'-----------------------------------------------------------------------------------------
Dim strtmp
Dim inttmp
Dim dtmtmp
Dim looptmp


'-----------------------------------------------------------------------------------------
' Pulisci Input	
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	25.11.2003 | 11.05.2004
' Commenti:	
'-----------------------------------------------------------------------------------------
function FilterSQLInput(ByVal input)

	'Remove malicious input for SQL execution from data
	input = Replace(input, "&", "&amp;", 1, -1, 1)
	input = Replace(input, "<", "&lt;")
	input = Replace(input, ">", "&gt;")
	input = Replace(input, "[", "&#091;")
	input = Replace(input, "]", "&#093;")
	input = Replace(input, """", "", 1, -1, 1)
	input = Replace(input, "=", "&#061;", 1, -1, 1)
	input = Replace(input, "'", "''", 1, -1, 1)
	input = Replace(input, "select", "sel&#101;ct", 1, -1, 1)
	input = Replace(input, "join", "jo&#105;n", 1, -1, 1)
	input = Replace(input, "union", "un&#105;on", 1, -1, 1)
	input = Replace(input, "where", "wh&#101;re", 1, -1, 1)
	input = Replace(input, "insert", "ins&#101;rt", 1, -1, 1)
	input = Replace(input, "delete", "del&#101;te", 1, -1, 1)
	input = Replace(input, "update", "up&#100;ate", 1, -1, 1)
	input = Replace(input, "like", "lik&#101;", 1, -1, 1)
	input = Replace(input, "drop", "dro&#112;", 1, -1, 1)
	input = Replace(input, "create", "cr&#101;ate", 1, -1, 1)
	input = Replace(input, "modify", "mod&#105;fy", 1, -1, 1)
	input = Replace(input, "rename", "ren&#097;me", 1, -1, 1)
	input = Replace(input, "alter", "alt&#101;r", 1, -1, 1)
	input = Replace(input, "cast", "ca&#115;t", 1, -1, 1)

	FilterSQLInput = input
	
end function


'-----------------------------------------------------------------------------------------
' Purifica Input	
'-----------------------------------------------------------------------------------------
' Funzione: 
' Data: 	25.11.2003 | 25.11.2003
' Commenti:	
'-----------------------------------------------------------------------------------------
function CleanInput(ByVal input)

	'Elimina i valori
	input = Replace(input, "&", "", 1, -1, 1)
	input = Replace(input, "<", "", 1, -1, 1)
	input = Replace(input, ">", "", 1, -1, 1)
	input = Replace(input, "'", "", 1, -1, 1)
	input = Replace(input, """", "", 1, -1, 1)

	CleanInput = input
	
end function


'-----------------------------------------------------------------------------------------
' Permetti Accesso	
'-----------------------------------------------------------------------------------------
' Funzione: 
' Data: 	30.11.2003 | 30.11.2003
' Commenti:	
'-----------------------------------------------------------------------------------------
function AllowEntry(ByVal nessuno, ByVal limitato, ByVal completo, ByVal protezione)
	
	Dim aryAsgPermetti(2)
	
	aryAsgPermetti(0) = CBool(nessuno)
	aryAsgPermetti(1) = CBool(limitato)
	aryAsgPermetti(2) = CBool(completo)
	
	If aryAsgPermetti(protezione) = False Then
	
		If Session("AsgLogin") <> "Logged" Then
			
			'Pulisci
			Set objAsgRs = Nothing
			objAsgConn.Close
			Set objAsgConn = Nothing
			
			'Indirizza
			Response.Redirect("login.asp?backto=" & Server.URLEncode(Request.ServerVariables("URL")))
		
		End If
		
	End If

end function

%>