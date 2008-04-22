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
