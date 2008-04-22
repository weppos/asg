<%

'/**
' * ASP Stats Generator - Powerful and reliable ASP website counter
' *
' * This file is part of the ASP Stats Generator package.
' * (c) 2003-2008 Simone Carletti <weppos@weppos.net>, All Rights Reserved
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
' * @copyright       2003-2008 Simone Carletti
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id$
' */

'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


' 
' Sends an HTTP Request to given url.
' method variable denotes the method used for the HTTP request.
' 
' Returns a string containing the HTTP response.
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
public function asgHttpGetRequest(url)
  asgHttpGetRequest = asgHttpRequest("GET", url)
end function

' 
' Sends an HTTP POST Request to given url.
' method variable denotes the method used for the HTTP request.
' 
' Returns a string containing the HTTP response.
'
public function asgHttpPostRequest(url)
  asgHttpPostRequest = asgHttpRequest("POST", url)
end function


%>
