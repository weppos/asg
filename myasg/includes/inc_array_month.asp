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
Dim intAsgMonthLoop
'Dichiara i Risultati
Dim aryAsgMonth(12, 2)


aryAsgMonth(0,1) = 0
aryAsgMonth(0,2) = ""
aryAsgMonth(1,1) = 1
aryAsgMonth(1,2) = strAsgTxtJanuary
aryAsgMonth(2,1) = 2
aryAsgMonth(2,2) = strAsgTxtFebruary
aryAsgMonth(3,1) = 3
aryAsgMonth(3,2) = strAsgTxtMarch
aryAsgMonth(4,1) = 4
aryAsgMonth(4,2) = strAsgTxtApril
aryAsgMonth(5,1) = 5
aryAsgMonth(5,2) = strAsgTxtMay
aryAsgMonth(6,1) = 6
aryAsgMonth(6,2) = strAsgTxtJune
aryAsgMonth(7,1) = 7
aryAsgMonth(7,2) = strAsgTxtJuly
aryAsgMonth(8,1) = 8
aryAsgMonth(8,2) = strAsgTxtAugust
aryAsgMonth(9,1) = 9
aryAsgMonth(9,2) = strAsgTxtSeptember
aryAsgMonth(10,1) = 10
aryAsgMonth(10,2) = strAsgTxtOctober
aryAsgMonth(11,1) = 11
aryAsgMonth(11,2) = strAsgTxtNovember
aryAsgMonth(12,1) = 12
aryAsgMonth(12,2) = strAsgTxtDecember

%>
