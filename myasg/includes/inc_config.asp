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


'Server MapPath
Dim strAsgMapPath
Dim strAsgMapPathTo
Dim strAsgMapPathIP  



'							========================================
'---------------------------   		Percorsi di collegamento		 -------------------------------------
'							========================================

'-----------------------------------------------------------------------------------------
' strAsgPathFolderDb
'-----------------------------------------------------------------------------------------
' Contiene il percorso alla cartella del server contenente i database
' NB. E' necessaria una cartella speciale con attivi i permessi di scrittura
' ed in genere è prevista da ogni provider che supporti l'uso di Access.
Const strAsgPathFolderDb = "mdb/" 
'
' Ecco alcuni esempi:
'
' // Sottocartella applicazione (myasg) '/mdb' - Include relativi
' Const strAsgPathFolderDb = "mdb/"
'
' // Sottocartella applicazione (myasg) '/mdb' - Include assoluti
' Const strAsgPathFolderDb = "/myasg/mdb/"
'
' // Sottocartella root  '/mdb' - Include relativi
' Const strAsgPathFolderDb = "../mdb/"
'
' // Sottocartella root  '/mdb' - Include assoluti
' Const strAsgPathFolderDb = "/mdb/"
'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------
' strAsgPathFolderWr
'-----------------------------------------------------------------------------------------
' Contiene il percorso alla cartella del server con attivi i permessi
' di scrittura e dove inserire il file inc_skin_file.asp
' E' possibile specificare la stessa del database.
Const strAsgPathFolderWr = "mdb/"




'							========================================
'---------------------------   			Nomi dei database			 -------------------------------------
'							========================================

'-----------------------------------------------------------------------------------------
' Statistiche
'-----------------------------------------------------------------------------------------
'
Const strAsgDatabaseSt = "dbstats" 

'-----------------------------------------------------------------------------------------
' IP
'-----------------------------------------------------------------------------------------
'
Const strAsgDatabaseIp = "ip-to-country" 


' Prefisso dei campi della tabella statistiche
' Utile se si vuole integrare l'applicazione in una tabella esistente per
' evitare conflitti
Const strAsgTablePrefix = "tblst_"


' Prefisso dei cookie
Const strAsgCookiePrefix = "ASG"




'							========================================
'---------------------------   		Collegamento al database		-------------------------------------
'							========================================

'-----------------------------------------------------------------------------------------
' Specifiche MapPath al database
'-----------------------------------------------------------------------------------------
strAsgMapPath = Server.MapPath(strAsgPathFolderDb & strAsgDatabaseSt & ".mdb")
strAsgMapPathTo = Server.MapPath(strAsgPathFolderDb & strAsgDatabaseSt & ".bak")
strAsgMapPathIP = Server.MapPath(strAsgPathFolderDb & strAsgDatabaseIp & ".mdb")




'							========================================
'---------------------------   Stringhe di connessione al database	-------------------------------------
'							========================================

'-----------------------------------------------------------------------------------------
' Microsoft Access 97	
'-----------------------------------------------------------------------------------------

' Driver Specifico
'-----------------
'strAsgConn = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & strAsgMapPath

' Driver OLEDB più veloce ed affidabile
'--------------------------------------
'strAsgConn = "Provider=Microsoft.Jet.OLEDB.3.51; Data Source=" & strAsgMapPath

'-----------------------------------------------------------------------------------------
' Microsoft Access 2000 - 2002 - 2003	
'-----------------------------------------------------------------------------------------

' Driver OLEDB più veloce ed affidabile
'--------------------------------------
strAsgConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strAsgMapPath



'-----------------------------------------------------------------------------------------
' Stampa di Debug stringa
'-----------------------------------------------------------------------------------------
If Request.QueryString("print") = "true" Then Response.Write(strAsgPathFolderDb & "<br />" & strAsgPathFolderWr & "<br />")


%>
