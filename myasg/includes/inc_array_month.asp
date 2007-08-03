<%

'/**
' * ASP Stats Generator - Powerful and reliable ASP website counter
' *
' * This file is part of the ASP Stats Generator package.
' * (c) 2003-2007 Simone Carletti <weppos@weppos.net>, All Rights Reserved
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
' * @copyright       2003-2007 Simone Carletti, All Rights Reserved
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id$
' */
 
'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


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
