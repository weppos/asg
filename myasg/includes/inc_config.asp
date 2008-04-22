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
