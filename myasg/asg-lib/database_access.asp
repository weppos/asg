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
' Creates and returns an Access Database connection string
' for given strDataSource.
' Microsoft.Jet.OLEDB.4.0 driver is used.
'
' @param  string  strDataSource
' @return string
'
function asgDatabaseAccessConnectionString(strDataSource)
  Dim strConnectionString
  
  strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strDataSource
  asgDatabaseAccessConnectionString = strConnectionString
end function


'
' Compacts an Access MDB database.
'
' Be aware that the compact function doesn't overwrite the original database
' but creates a copy with the compacted one.
' Please use asgFileRename() function to replace the old database with the compacted one.
'
' @param  string  strSourcePath
' @param  string  strTargetPath
'
function asgDatabaseAccessCompact(strOriginalPath, strCompactedPath)
  Dim strSourceConnection, strTargetConnection
  Dim objJro
  
  strSourceConnection = asgDatabaseAccessConnectionString(strOriginalPath)
  strTargetConnection = asgDatabaseAccessConnectionString(strCompactedPath)
  
  ' delete target file, if it already exists,
  ' to prevent compactDatabase to crash
  asgFileDeleteIfExists(strCompactedPath)

  set objJro = Server.CreateObject("jro.JetEngine")
  objJro.compactDatabase strSourceConnection, strTargetConnection
  Set objJro = Nothing
end function


%>
