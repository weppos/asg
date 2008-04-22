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
' * @version         SVN: $Id$
' */

'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


Class HttpTest

  Public Function TestCaseNames()
    TestCaseNames = Array("test", "test2", "test3")
  End Function

  Public Sub SetUp()
    'Response.Write("SetUp<br>")
  End Sub

  Public Sub TearDown()
    'Response.Write("TearDown<br>")
  End Sub

  Public Sub test(oTestResult)
    'Response.Write("test<br>")
    Err.Raise 5, "hello", "error"
  End Sub

  Public Sub test2(oTestResult)
    'Response.Write("test2<br>")
    oTestResult.Assert False, "Assert False!"

    oTestResult.AssertEquals 4, 4, "4 = 4, Should not fail!"
    oTestResult.AssertEquals 4, 5, "4 != 5, Should fail!"
    oTestResult.AssertNotEquals 5, 5, "AssertNotEquals(5, = 5) should fail!"

        oTestResult.AssertExists new TestResult, "new TestResult Should not fail!"
        oTestResult.AssertExists Nothing, "Nothing: Should not exist!"
        oTestResult.AssertExists 4, "4 Should exist?!"
  End Sub

  Public Sub test3(oTestResult)
    oTestResult.Assert True, "Success"
  End Sub
  
End Class

%>