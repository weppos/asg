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
' Checks whether a newer release is available
' and returns an array with latest release version, date and url.
'
public function asgVersionCheck(strCurrentVersion)
  
  Dim strHost, strVersion, strUrl, strResponse
  Dim aryLastVersion
  
  strHost = Request.ServerVariables("HTTP_HOST")
  strVersion = strCurrentVersion
  strUrl = "http://www.asp-stats.com/api/v1/version_check?" &_
           "host=" & Server.URLEncode(strHost) & "&" &_
           "version=" & Server.URLEncode(strVersion)
  
  strResponse = asgHttpGetRequest(strUrl)
  if varType(strResponse) = 8 then ' successful response
    aryLastVersion = split(strResponse, "|")
  else
    aryLastVersion = array()
  end if
  
  asgVersionCheck = aryLastVersion

end function


%>