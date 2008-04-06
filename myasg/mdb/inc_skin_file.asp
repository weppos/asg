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
' * @copyright       2003-2008 Simone Carletti
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id$
' */
 
'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


'SPERIMENTALE!
Const strAsgSknPathImage      = "images/" 'Percorso della cartella immagini
'-----------------------------------------------------------------------------------------
' Generali	
'-----------------------------------------------------------------------------------------
Const strAsgSknPageBgColour    = "#F9F9F9" 'Colore di sfondo delle pagine
Const strAsgSknPageBgImage    = "" 'Immagine di sfondo delle pagine
Const strAsgSknPageWidth      = "900" 'Larghezza delle pagine <br />&nbsp;ATTENZIONE! Variare solo se necessario! Una bassa risoluzione causerà problemi nella visualizzazione
'-----------------------------------------------------------------------------------------
' Tabelle di Layout
'-----------------------------------------------------------------------------------------
Const strAsgSknTableLayoutBorderColour    = "#999999" 'Colore del bordo della tabella Layout
Const strAsgSknTableLayoutBgColour    = "#F6F6F6" 'Colore di sfondo della tabella Layout
Const strAsgSknTableLayoutBgImage    = "" 'Immagine di sfondo della tabella Layout
Const strAsgSknTableBarBgColour    = "" 'Colore di sfondo della Barra Header e Footer
Const strAsgSknTableBarBgImage    = "layout/bar_bg.jpg" 'Immagine di sfondo della Barra Header e Footer
'-----------------------------------------------------------------------------------------
' Tabelle Dati
'-----------------------------------------------------------------------------------------
Const strAsgSknTableTitleBgColour    = "" 'Colore di sfondo del titolo della tabella contenuti
Const strAsgSknTableTitleBgImage    = "layout/small_bar_bg.jpg" 'Immagine di sfondo del titolo della tabella contenuti
Const strAsgSknTableContBgColour    = "#ECEAF2" 'Colore di sfondo della tabella contenuti
Const strAsgSknTableContBgImage    = "" 'Immagine di sfondo della tabella contenuti
'-----------------------------------------------------------------------------------------
' Configurazioni opzionali
'-----------------------------------------------------------------------------------------
'Mostra il tempo di elaborazione
Const blnAsgElabTime = True 'Mostra il tempo di elaborazione.<br />&nbsp;Accetta valori <strong>True</strong> o <strong>False</strong>

%>
