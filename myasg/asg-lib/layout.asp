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


' Dependencies:
' - /asg-lib/file.asp.


'
' Returns the elaboration time message since elaboration started.
'
' Note. This function uses the "global" startAsgElab variable.
' Be aware of side effects!
' 
' @return string
'
public function asgElabtime()
  Dim strElabtime
  
  strElabtime = asgComputeElabtime(startAsgElab, Timer())
  asgElabtime = Replace(ASG_TEXT_PAGE_GENERATED_IN, "%{seconds}", strElabtime)
end function 


'
' Returns the web path to the flag image
' corresponding to given strCountryCode.
' This is basically an alias for asgCountryFlagIcon().
'
' @param  string  strPath
' @param  string  strCountryCode
' @return string
'
public function asgCountryFlagIcon(strPath, strCountryCode)
  asgCountryFlagIcon = asgFlagIcon(strPath, strCountryCode)
end function


'
' Returns the web path to the flag image
' corresponding to given strLanguage.
' This function internally uses asgCountryFlagIcon().
'
' TODO: remove as soon as ASG will store country codes
' instead of hard stored language names.
'
' @param  string  strPath
' @param  string  strCountryCode
' @return string
'
public function asgLanguageFlagIcon(strPath, strLanguage)
  Dim cc, ln
  
  ln = lcase(strLanguage)
  if instr(ln, "italiano") then
    cc = "it"
  elseif instr(ln, "inglese") then
    cc = "gb"
  elseif instr(ln, "francese") then
    cc = "fr"
  elseif instr(ln, "tedesco") then
    cc = "de"
  elseif instr(ln, "spagnolo") then
    cc = "es"
  elseif instr(ln, "catalano") then
    cc = "catalonia"
  elseif instr(ln, "portoghese") then
    cc = "pt"
  elseif instr(ln, "slovacco") then
    cc = "sk"
  elseif instr(ln, "bulgaro") then
    cc = "bg"
  elseif instr(ln, "croato") then
    cc = "hr"
  elseif instr(ln, "ceco") then
    cc = "cs"
  elseif instr(ln, "macedone") then
    cc = "mk"
  elseif instr(ln, "albanese") then
    cc = "sq"
  elseif instr(ln, "polacco") then
    cc = "pl"
  elseif instr(ln, "serbo") then
    cc = "sr"
  elseif instr(ln, "svedese") then
    cc = "sv"
  elseif instr(ln, "russo") then
    cc = "ru"
  elseif instr(ln, "ungherese") then
    cc = "hu"
  elseif instr(ln, "estone") then
    cc = "et"
  elseif instr(ln, "lituano") then
    cc = "lt"
  elseif instr(ln, "norvegese") then
    cc = "no"
  elseif instr(ln, "finlandese") then
    cc = "fi"
  elseif instr(ln, "danese") then
    cc = "dk"
  elseif instr(ln, "olandese") then
    cc = "nl"
  elseif instr(ln, "greco") then
    cc = "gr"
  elseif instr(ln, "lettone") then
    cc = "lv"
  elseif instr(ln, "sloveno") then
    cc = "sl"
  elseif instr(ln, "rumeno") then
    cc = "ro"
  elseif instr(ln, "turco") then
    cc = "tk"
  elseif instr(ln, "giapponese") then
    cc = "jp"
  elseif instr(ln, "cinese") then
    cc = "cn"
  elseif instr(ln, "canada") then
    cc = "ca"
  elseif instr(ln, "messico") then
    cc = "mx"
  else
    cc = "xx"
  end if
  
  asgLanguageFlagIcon = asgFlagIcon(strPath, cc)
end function


'
' Returns the web path to requested flag icon, if exists. 
' The flag icon web path is composed by strPath
' and the strCountryCode ISO country code.
' Returns the path to a generic flag icon
' if requested flag doesn't exist.
'
' TODO: this function will probably be removed or changed
' in a future (an more reliable) ASG version.
'
' @param  string  strPath
' @param  string  strCountryCode
' @return string
'
public function asgFlagIcon(strPath, strCountryCode)
  Dim p, f
  
  f = lcase(strCountryCode)
  if f = "xx" then
    p = strPath & "xx.gif"
  elseif asgFileExists(Server.MapPath(strPath & f & ".gif")) then
    p = strPath & f & ".gif"
  else
    p = strPath & "xx.gif"
  end if
  
  asgFlagIcon = p
end function 


%>
