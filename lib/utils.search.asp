<!--#include file="utils.datetime.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'



'/**
' * Highlight searched keywords.
' * 
' * @param		
' * @param		
' * @return 	string § string with keywords highlighted.
' *
' * @since 		3.0
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function searchTerms(input, databaseField, searchFor, searchIn)

	' If some data has been searched and this is the database 
	' where you have searched in then highlight search terms
	if Len(searchFor) > 0 AND Len(searchIn) > 0 AND searchIn = databaseField then
		input = Replace(input, searchFor, "<span class=""highlighted"">" & searchFor & "</span>", 1, -1, vbTextCompare)
	end If
	
	' Return function
	'searchTerms = Server.HTMLEncode(argString)
	searchTerms = input

end function

%>