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
' Copies a file from strSourcePath to strTargetPath.
'
' @param  strSourcePath
' @param  strTargetPath
' @param  blnOverwrite
'
function asgFileCopy(strSourcePath, strTargetPath, blnOverwrite)
  Dim objFso

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  objFso.copyFile strSourcePath, strTargetPath, blnOverwrite
  Set objFso = Nothing
end function


'
' Deletes the file at strSourcePath.
'
' @param  strSourcePath
' @param  strTargetPath
'
function asgFileDelete(strSourcePath)
  Dim objFso

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  objFso.deleteFile strSourcePath
  Set objFso = Nothing
end function


'
' Deletes the file at strSourcePath
' only if the file exists.
'
' The main difference between this method
' and asgFileDelete is that the latter
' doesn't check whether strSourcePath exists
' and crashes on failure.
'
' @param  strSourcePath
' @param  strTargetPath
'
function asgFileDeleteIfExists(strSourcePath)
  Dim objFso

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  if objFso.fileExists(strSourcePath) then
    objFso.deleteFile strSourcePath
  end if
  Set objFso = Nothing
end function


'
' Moves a file from strSourcePath to strTargetPath.
'
' @param  strSourcePath
' @param  strTargetPath
'
function asgFileMove(strSourcePath, strTargetPath)
  Dim objFso

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  objFso.moveFile strSourcePath, strTargetPath
  Set objFso = Nothing
end function


'
' Replace file strTargetPath with strSourcePath.
' If strTargetPath doesn't exist, strSourcePath file is simply renamed.
'
' @param  strSourcePath
' @param  strTargetPath
'
function asgFileReplace(strSourcePath, strTargetPath)
  Dim objFso

  Set objFso = Server.CreateObject("Scripting.FileSystemObject")
  objFso.copyFile strSourcePath, strTargetPath, true
  objFso.deleteFile strSourcePath
  Set objFso = Nothing
end function


%>
