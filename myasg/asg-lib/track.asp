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
' Returns absolute path to track file, 
' build from request server variables.
'
function asgTrackFilePath()
  Dim strScheme, strDomain, strPath
  Dim strFilePath

  if Request.ServerVariables("HTTPS") = "on" then 
    strScheme = "https"
  else
    strScheme = "http"
  end if
  strDomain = Request.ServerVariables("SERVER_NAME")
  strPath   = Replace(Request.ServerVariables("SCRIPT_NAME"), "asg-track.js.asp", "asg-track.asp")

  strFilePath = strScheme & "://" & strDomain & strPath
  asgTrackFilePath = strFilePath
end function


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


%>
