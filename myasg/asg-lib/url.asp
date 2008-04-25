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
' Returns strUrl without the query string part.
'
' @param  string  strUrl
'
function asgUrlStripQuery(strUrl)
  Dim intInstr, strValue

  intInstr = instr(strURL, "?")
  if intInstr then strValue = left(strURL, intInstr - 1) else strValue = strURL
  asgUrlStripQuery = strValue

end function


'
' Returns strUrl without the scheme part.
'
' @param  string  strUrl
'
function asgUrlStripScheme(strUrl)
  Dim intInstr, strValue

  intInstr = instr(strURL, "://")
  if intInstr then strValue = right(strURL, len(strURL) - (3 + intInstr - 1)) else strValue = strURL
  asgUrlStripScheme = strValue

end function


'
' Returns the domain part from strUrl.
'
' @param  string  strUrl
' 
' TODO: trim trailing slash and update all dependencies
'
function asgUrlGetDomain(strUrl)
  Dim intInstr, strValue

  strValue = asgUrlStripScheme(strUrl)
  intInstr = instr(strValue, "/")
  if intInstr > 0 then strValue = left(strValue, intInstr) else strValue = strValue & "/"
  asgUrlGetDomain = strValue

end function


'
' Decodes a string encoded with Server.URLEncode().
'
' @param  string  strUrl
' @link   http://www.aspnut.com/reference/encoding.asp
'
function asgUrlDecode(strUrl)
    Dim arySplit
    Dim strOutput
    Dim ii
    
    if isNull(strUrl) then
       asgUrlDecode = ""
       exit function
    end If

    ' convert all pluses to spaces
    strOutput = REPLACE(strUrl, "+", " ")

    ' next convert %hexdigits to the character
    arySplit = Split(strOutput, "%")

    if isArray(arySplit) then
      strOutput = arySplit(0)
      for ii = 0 to UBound(arySplit) - 1
        strOutput = strOutput & _
          Chr("&H" & Left(arySplit(ii + 1), 2)) &_
          Right(arySplit(ii + 1), Len(arySplit(ii + 1)) - 2)
      next
    end if

    asgUrlDecode = strOutput
end function


%>
