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


' 
' Sends an HTTP Request to given url.
' method variable denotes the method used for the HTTP request.
' 
' Returns a string containing the HTTP response.
'
' @param    string  method
' @param    string  url
' @return   string
'
' TODO: validate err.number
' 
public function asgHttpRequest(method, url)
  Dim objXmlHttp
  Dim intStatus, strResponse
  
  Set objXmlHttp = Server.CreateObject("Microsoft.XMLHTTP")
  'on error resume next 
  
  objXmlHttp.open method, url, false
  'objXmlHttp.setRequestHeader "User-Agent", "foo"
  objXmlHttp.send
  
  intStatus = objXmlHttp.status 
  'if err.number <> 0 or intStatus <> 200 then
  if intStatus <> 200 then
    strResponse = intStatus
  else
    strResponse = CStr(objXmlHttp.ResponseText)
  end if  
  
  Set objXmlHttp = Nothing
  asgHttpRequest = strResponse
end function

' 
' Sends an HTTP GET Request to given url.
' method variable denotes the method used for the HTTP request.
' 
' Returns a string containing the HTTP response.
'
' @param    string  url
' @return   string
'
public function asgHttpGetRequest(url)
  asgHttpGetRequest = asgHttpRequest("GET", url)
end function

' 
' Sends an HTTP POST Request to given url.
' method variable denotes the method used for the HTTP request.
' 
' Returns a string containing the HTTP response.
'
' @param    string  url
' @return   string
'
public function asgHttpPostRequest(url)
  asgHttpPostRequest = asgHttpRequest("POST", url)
end function


%>
