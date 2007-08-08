<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


'/**
' * Test a regular expression pattern on a string.
' * The test is not case sensitive.
' * 
' * @param 		(string) pattern	- the regular expression pattern to test.
' * @param	 	(string) text		- the string to search for the test.
' * @return 	(bool) true if the test returns some results,
' *				false otherwise.
' *
' * @since 		3.0
' *
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function regexpTest(pattern, text)

	Dim objRegexp, found
	Set objRegexp = New RegExp
  
	objRegexp.Pattern = pattern
	objRegexp.IgnoreCase = true
	found = objRegexp.Test(text)

	regexpTest = found

end function

'/** 
' * Strip out malicious SQL characters from text.
' * 
' * @param		(string) input 		- the text to be filtered
' * @param		(bool) removeHTML	- set to true to strip out HTML tags
' * @param		(bool) removeApex	- set to true to strip out single quote
' * @return 	(string) the cleaned string.
' * 
' * @author		Simone Carletti <carletti@weppos.net>
' */ 
public function filterSQLinput(input, removeHTML, removeApex)

	'input = Server.HTMLEncode(input)

	if removeHTML then
	input = Replace(input, "<", "&lt;")
	input = Replace(input, ">", "&gt;")
	input = Replace(input, "=", "&#061;", 1, -1, 1)
	input = Replace(input, """", "&quot;", 1, -1, 1)
	end if
	if removeApex then
	input = Replace(input, "'", "''", 1, -1, 1)
	end if
	input = Replace(input, "]", "&#093;")
	input = Replace(input, "[", "&#091;")
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

	filterSQLinput = input

end function


%>