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


' Server MapPath
Dim strAsgMapPath
Dim strAsgMapPathTo
Dim strAsgMapPathIP  



' ----- BEGINNING OF CONFIGURATION SETTINGS -----


' ***** PUBLIC FOLDER *****
' This is the path to database folder.
' Each web host usually provides a special folder for Access databases
' with write permission enabled and browse permission denied.
' The folder is writeable by the web user but the public access is denied
' in order to prevent anyone from downloading your database.
'
' Be sure to move your database into your database folder
' and change the following path.
' You can use both relative and absolute path.
'
' Each path starts from /myasg folder and follows the same rules ad links in web pages.
' To specify a path to a child folder (for example 'myasg/mdb/') simply enter 'mdb/'.
' If your folder is a parent directory then use '../' to go back to parent folder
' or use an absolute path from your web root, for example '/absolute/path/to/folder/'
' means the folder available at http://example.com/absolute/path/to/folder/.
Const strAsgPathFolderDb = "mdb/" 

' ***** WRITE PERMISSION FOLDER *****
' This is the path to a folder with write permission enabled
' Write permission are required by inc_skin_file.asp skin file
' It could be the same as database folder
Const strAsgPathFolderWr = "mdb/"


' ***** DATABASE NAME *****
' Main application database name
Const strAsgDatabaseSt = "dbstats" 

' ***** IP2COUNTRY DATABASE NAME *****
' This is the name of the ip-to-country database
' used to get country information from IP address
Const strAsgDatabaseIp = "ip-to-country" 


' ***** TABLE PREFIX *****
' Prefix that all ASG database tables will have.
' This is useful if you want to run multiple versions or copies on the same database 
' or if you are sharing the database with other applications.
Const strAsgTablePrefix = "tblst_"


' ***** COOKIE PREFIX *****
' Prefix that all ASG cookies will have.
' This is useful if you run multiple copies of the program on the same site so that 
' cookies don't interfer with each other.
Const strAsgCookiePrefix = "asg_"


' ----- END OF CONFIGURATION SETTINGS -----



strAsgMapPath = Server.MapPath(strAsgPathFolderDb & strAsgDatabaseSt & ".mdb")
strAsgMapPathTo = Server.MapPath(strAsgPathFolderDb & strAsgDatabaseSt & ".bak")
strAsgMapPathIP = Server.MapPath(strAsgPathFolderDb & strAsgDatabaseIp & ".mdb")

strAsgConn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strAsgMapPath

%>
