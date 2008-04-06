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
' * @package         
' * @author          Simone Carletti <weppos@weppos.net>
' * @copyright       2003-2007 Simone Carletti, All Rights Reserved
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id: default.asp 14 2007-08-03 13:25:18Z weppos $
' */
 
'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


'/**
' * PHP like associative array
' *
' * This object provides the ability to create an associative array in ASP.
' * Array keys could be strings instead of integer as required by ASP 3.0 language.
' * 
' * Internally, the associative array is managed as a Scripting.Dictionary object.
' *
' *
' *     Dim asyArray
' *     Set asyArray = new AssociativeArray()
' *     
' *     asyArray("name")("first") = "Sujoy"
' *     asyArray("name")("last") = "Roy"
' *
' *
' * @category        ASP Stats Generator
' * @package         
' * @author          
' * @copyright       
' * @license         
' */
Class AssocArray
  Private dicContainer
  
  Private Sub Class_Initialize()
   Set dicContainer=Server.CreateObject("Scripting.Dictionary")
  End Sub
  
  Private Sub Class_Terminate()
   Set dicContainer=Nothing   
  End Sub


  Public Default Property Get Item(sName)
   If Not dicContainer.Exists(sName) Then
    dicContainer.Add sName,New AssocArray
   End If
   
   If IsObject(dicContainer.Item(sName)) Then
    Set Item=dicContainer.Item(sName)
   Else
    Item=dicContainer.Item(sName)
   End If   
  End Property
  
  Public  Property Let Item(sName,vValue)
   If dicContainer.Exists(sName) Then
    If IsObject(vValue) Then
     Set dicContainer.Item(sName)=vValue
    Else
     dicContainer.Item(sName)=vValue
    End If
   Else
    dicContainer.Add sName,vValue    
   End If
  End Property
End Class

%>