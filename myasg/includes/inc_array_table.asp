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


'Ciclo di Elaborazione
Dim intAsgTableLoop
'Dichiara i Risultati
Dim aryAsgTable(10, 2)


aryAsgTable(0,1) = "All"
aryAsgTable(0,2) = strAsgTxtResetAllTables
aryAsgTable(1,1) = "Detail"
aryAsgTable(1,2) = strAsgTxtDetailContent
aryAsgTable(2,1) = "System"
aryAsgTable(2,2) = strAsgTxtSystemContent
aryAsgTable(3,1) = "Daily"
aryAsgTable(3,2) = strAsgTxtDailyContent
aryAsgTable(4,1) = "Hourly"
aryAsgTable(4,2) = strAsgTxtHourlyContent
aryAsgTable(5,1) = "Language"
aryAsgTable(5,2) = strAsgTxtLanguageContent
aryAsgTable(6,1) = "Referer"
aryAsgTable(6,2) = strAsgTxtRefererContent
aryAsgTable(7,1) = "Page"
aryAsgTable(7,2) = strAsgTxtPageContent
aryAsgTable(8,1) = "Query"
aryAsgTable(8,2) = strAsgTxtQueryContent
aryAsgTable(9,1) = "Country"
aryAsgTable(9,2) = strAsgTxtCountryContent
aryAsgTable(10,1) = "IP"
aryAsgTable(10,2) = strAsgTxtIPContent

%>
