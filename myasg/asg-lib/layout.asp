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
