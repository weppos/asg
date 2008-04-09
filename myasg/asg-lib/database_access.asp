<%

'/**
' * ASP Stats Generator - Powerful and reliable ASP website counter
' *
' * This file is part of the ASP Stats Generator package.
' * (c) 2003-2008 Simone Carletti <weppos@weppos.net>, All Rights Reserved
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
' * @version         SVN: $Id: functions_stats.asp 8 2007-08-03 12:51:40Z weppos $
' */

'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


'
' Creates and returns an Access Database connection string
' for given strDataSource.
' Microsoft.Jet.OLEDB.4.0 driver is used.
'
' @param  string  strDataSource
' @return string
'
function asgDatabaseAccessConnectionstring(strDataSource)
  Dim strConnectionString
  
  strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & strDataSource
  asgDatabaseAccessConnectionstring = strConnectionString
end function


'
' Compacts an Access MDB database.
'
' Be aware that the compact function doesn't overwrite the original database
' but creates a copy with the compacted one.
' Please use fileRename() function to replace the old database with the compacted one.
'
' @param  string  strSourcePath
' @param  string  strTargetPath
'
function asgDatabaseAccessCompact(strOriginalPath, strCompactedPath)
  Dim strSourceConnection, strTargetConnection
  Dim objJro
  
  strSourceConnection = asgDatabaseAccessConnectionstring(strOriginalPath)
  strTargetConnection = asgDatabaseAccessConnectionstring(strCompactedPath)
  
  set objJro = CreateObject("jro.JetEngine") 
  objJro.compactDatabase strSourceConnection, strTargetConnection
  Set objJro = Nothing 
end function


%>
