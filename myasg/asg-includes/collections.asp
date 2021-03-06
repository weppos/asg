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
' @version         SVN: $Id: binary.asp 125 2008-04-22 20:44:14Z weppos $
' 


'
' Returns an array containing a string representation
' for each month of the Year.
' 
' @return array
'
public function asgArrayMonths()
  Dim aryMonths(12)

  aryMonths(1) = ASG_LABEL_JANUARY
  aryMonths(2) = ASG_LABEL_FEBRUARY
  aryMonths(3) = ASG_LABEL_MARCH
  aryMonths(4) = ASG_LABEL_APRIL
  aryMonths(5) = ASG_LABEL_MAY
  aryMonths(6) = ASG_LABEL_JUNE
  aryMonths(7) = ASG_LABEL_JULY
  aryMonths(8) = ASG_LABEL_AUGUST
  aryMonths(9) = ASG_LABEL_SEPTEMBER
  aryMonths(10) = ASG_LABEL_OCTOBER
  aryMonths(11) = ASG_LABEL_NOVEMBER
  aryMonths(12) = ASG_LABEL_DECEMBER
  
  asgArrayMonths = aryMonths
end function 


%>
