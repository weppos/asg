<% @LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="w2k3_config.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


' Reset objects
Set objAsgRs = Nothing
objAsgConn.Close
Set objAsgConn = Nothing


' Autologin user
Dim blnAutologin
blnAutologin = checkPermissionCookie()

' Hide menubar
if ASG_MENUBAR_HIDELOGIN then blnAsgShowToolbar = false


' Check login
if Len(Request.Form("password")) > 0 then

	' Check session value
	Call checkSessionID(Request.Form("sessionid"))

	Dim strAsgPassword
	Dim blnAsgAutologin
	Dim blnAsgError
	
	' Common error variable
	blnAsgError = false
	
	' Collect data from the form
	strAsgPassword = Trim(Request.Form("password"))
	strAsgPassword = CleanInput(strAsgPassword)
	blnAsgAutologin = Cbool(Request.Form("autologin"))

	' Compare the strings
	if StrComp(strAsgPassword, appAsgSitePsw) = 0 then

		' Set the session variable
		Session("asgLogin") = "Logged"
		
		' Set the cookie
		if blnAsgAutologin then
			Response.Cookies(ASG_COOKIE_PREFIX & "autologin")("autologin") = "true"
			Response.Cookies(ASG_COOKIE_PREFIX & "autologin")("psw") = strAsgPassword
			Response.Cookies(ASG_COOKIE_PREFIX & "autologin").Expires = dateAdd("yyyy", 1, date)
		end if
	
	else
		blnAsgError = true
		Session.Contents.Remove("asgLogin")
	end If

End If

'Logout
if Request.QueryString("logout") = "true" then 
	' Remove session variable
	Session.Contents.Remove("asgLogin")
	' Remove cookies
	if Request.Cookies(ASG_COOKIE_PREFIX & "autologin")("autologin") = "true" then
		Response.Cookies(ASG_COOKIE_PREFIX & "autologin")("autologin") = ""
		Response.Cookies(ASG_COOKIE_PREFIX & "autologin")("psw") = ""
	end if
end if

%>
<%= STR_ASG_PAGE_DOCTYPE %>
<html>
<head>
<title><%= appAsgSiteName %> | powered by ASP Stats Generator v<%= ASG_VERSION %></title>
<%= STR_ASG_PAGE_CHARSET %>
<meta name="copyright" content="Copyright (C) 2003-2005 Carletti Simone" />
<%	If Session("asgLogin") = "Logged" AND Len(Request.QueryString("backto")) > 0 Then %>
<meta http-equiv="Refresh" content="3;url=<%= Request.QueryString("backto") %>">
<%	ElseIf Session("asgLogin") = "Logged" AND NOT Len(Request.QueryString("backto")) > 0 Then %>
<meta http-equiv="Refresh" content="3;url=main.asp">
<%	End If %>
<!--#include file="includes/meta.inc.asp" -->

<!-- ASP Stats Generator v. <%= ASG_VERSION %> is created and developed by Simone Carletti.
To download your Free copy visit the official site http://www.weppos.com/asg/ -->

</head>

<body>
<!--#include file="includes/header.asp" -->

<div align="center">
	<div id="login">
		<% if Session("asgLogin") <> "Logged" then %>
		<div class="login_form">
      <form action="login.asp?backto=<%= Server.URLEncode(Request.QueryString("backto")) %>" name="frmLogin" method="post">
				<div class="login_box">
					<div><%= TXT_password %><br /><input type="password" name="password" value="" size="15" maxlength="20" /></div>
					<div><br /><input type="checkbox" name="autologin" value="true" /> <%= TXT_autologin %></div>
					<div><br />
						<input type="submit" name="login" value="<%= TXT_login %>" />
						<input type="hidden" name="sessionid" value="<%= Session.SessionID %>" />
					</div>
				</div>
			</form>
  	</div>
		<script language="javascript" type="text/javascript">document.frmLogin.password.focus();</script>
		<div class="login_text">
			<div class="ltcenter"><img src="images/images/keys.png" alt="<%= TXT_password %>" /></div>
      		<p><%= TXT_password_desc %></p>
		</div>
		<div class="ltclear"><br />
			<% if blnAsgError then Response.Write(vbCrLf & "<p class=""errortext"">" & TXT_password_wrong & "</p>") %>
			<p><%= TXT_cookiesMustBeEnabled %></p>
		</div>
		<% else %>
			<div class="ltcenter">
				<p><%= TXT_login_completed & "<br />" & TXT_login_entryAllowed %></p>
				<p><%= TXT_login_redirectPreviousPage  %></p>
				<p><a href="<% If Len(Request.QueryString("backto")) > 0 Then Response.Write(Request.QueryString("backto")) Else Response.Write("main.asp") %>" title="<%= TXT_login_clickAndGo %>"><%= TXT_login_clickAndGo %></a></p>
				<p><a href="login.asp?logout=true" title="<%= TXT_logout %>"><%= TXT_logout_execute %></a></p>
			</div>
		<% end if %>
	</div>
</div>

<br /></div>
<!-- / body -->
<%

' Footer
Response.Write(vbCrLf & "<div id=""footer"">")
' ***** START WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** INIZIO AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  	******
Response.Write(vbCrLf & "<br /><div style=""text-align: center;"">Powered by <a href=""http://www.weppos.com/asg/"" title=""ASP Stats Generator"">ASP Stats Generator</a> v" & ASG_VERSION & " ") 
if ASG_BUILDINFO then Response.Write("build " & ASG_VERSION_BUILD)
Response.Write(vbCrLf & "<br />Copyright &copy; 2003-2005 <a href=""http://www.weppos.com/"">weppos</a></div>")
if ASG_ELABORATION_TIME then Response.Write("<div class=""elabtime"">" & Replace(TXT_elabtime, "$time$", FormatNumber(Timer() - startAsgElab, 4)) & "</div>")
' ***** END WARNING - REMOVAL or MODIFICATION IN PART or ALL OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT	******
' ***** FINE AVVERTENZA - RIMOZIONE o MODIFICA PARZIALE/TOTALE DEL CODICE COMPORTA VIOLAZIONE DELLA LICENZA  ******
Response.Write(vbCrLf & "</div>")

%>
<!--#include file="includes/footer.asp" -->
</body></html>