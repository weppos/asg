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
' @version         SVN: $Id: track.asp 125 2008-04-22 20:44:14Z weppos $
' 


' 
' Performs a regular expression test.
' Searches +subject+ for a match to the regular expression given in +pattern+
' and returns true if +subject+ contains at least one match, false otherwise.
'
' @param    string  subject
' @param    string  pattern
' @param    bool    ignorecase
' @param    bool    global
' @return   bool
' 
public function asgRegExpTest(subject, pattern, ignorecase, global)
  Dim objRE
  Dim blnTest
  
  Set objRE = New RegExp
  with RegularExpressionObject
    .pattern = pattern
    .ignoreCase = ignorecase
    .global = global
  end with
  blnTest = objRE.test(subject)
  Set objRE = Nothing
  
  asgRegExpTest = blnTest
end function

' 
' Perform a regular expression search and replace.
'
' @param    string  subject
' @param    string  pattern
' @param    string  replacement
' @param    bool    ignorecase
' @param    bool    global
' 
public function asgRegExpReplace(subject, pattern, replacement, ignorecase, global)
  Dim objRE
  Dim strResult
  
  Set objRE = New RegExp
  with RegularExpressionObject
    .pattern = pattern
    .ignoreCase = ignorecase
    .global = global
  end with
  strResult = objRE.replace(subject, replacement)
  Set objRE = Nothing
  
  asgRegExpReplace = strResult
end function

' 
' Executes a regular expression match.
' Searches +subject+ for a match to the regular expression given in +pattern+
' and returns true if +subject+ contains at least one match, false otherwise.
' 
' The param +matches+ is filled with the results of the search.
' If at least one match exists, +matches+ becomes a 
' Regular Expression matches instance, else is set to null.
'
' If you simply need to perform a regular expression test
' you should use asgRegExpTest that uses the most appropriate #test method.
'
' @param    string  subject
' @param    string  pattern
' @param    string  replacement
' @param    bool    ignorecase
' @param    bool    global
' 
public function asgRegExpExecute(subject, pattern, ByRef matches, ignorecase, global)
  Dim objRE
  Dim blnMatches
  
  Set objRE = New RegExp
  with RegularExpressionObject
    .pattern = pattern
    .ignoreCase = ignorecase
    .global = global
  end with
  Set matches = objRE.execute(subject)
  Set objRE = Nothing
  
  if matches.count > 0 then
    blnMatches = true
  else
    blnMatches = false
    matches = null
  end if
  
  asgRegExpExecute = blnMatches
end function


%>
