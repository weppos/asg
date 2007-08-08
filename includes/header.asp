<div id="tip" style="visibility:hidden;position:absolute;z-index:1000;top:-100"></div>
<script language="JavaScript" type="text/javascript" src="3rdparty/tipmessage/tip_style.js"></script>
<script language="JavaScript" type="text/javascript" src="tip_warning.js.asp"></script>
<%

'-----------------------------------------------------------------------------------------
' Check for updates
'-----------------------------------------------------------------------------------------
' Esecuzioni in base a controllo
Select Case intAsgLastUpdate
	
	Case 0
		'Non calcolato
	Case 1
		'Corrisponde
	Case 2
		'Display the alert
		Response.Write("<script language=""JavaScript""><!-- ")
		Response.Write(vbCrLf & vbCrLf & "//Show the popup")
		Response.Write(vbCrLf & "checkUpdate = confirm('Available " & strAsgLastVersion & " version released on " & Right(dtmAsgLastUpdate, 2) & "/" & Mid(dtmAsgLastUpdate, 5, 2) & "/" & Left(dtmAsgLastUpdate, 4) & "! \nDownload the update?')")
		Response.Write(vbCrLf & "if (checkUpdate == true) {")
		If Len(urlAsgLastUpdate) > 0 Then
		Response.Write(vbCrLf & "	window.location='" & urlAsgLastUpdate & "'")
		Else
		Response.Write(vbCrLf & "	window.location='http://www.weppos.com/asg/'")
		End If
		Response.Write(vbCrLf & "}")
		Response.Write(vbCrLf & "// --></script>")
	Case 3
		'Display the alert
		Response.Write("<script language=""JavaScript""><!-- ")
		Response.Write(vbCrLf & vbCrLf & "//Show the popup")
		Response.Write(vbCrLf & "checkUpdate = confirm('Available for your " & strAsgLastVersion & " version an update released on " & Right(dtmAsgLastUpdate, 2) & "/" & Mid(dtmAsgLastUpdate, 5, 2) & "/" & Left(dtmAsgLastUpdate, 4) & ". \nDownload the update?')")
		Response.Write(vbCrLf & "if (checkUpdate == true) {")
		If Len(urlAsgLastUpdate) > 0 Then
		Response.Write(vbCrLf & "	window.location='" & urlAsgLastUpdate & "'")
		Else
		Response.Write(vbCrLf & "	window.location='http://www.weppos.com/asg/'")
		End If
		Response.Write(vbCrLf & "}")
		Response.Write(vbCrLf & "// --></script>")

End Select

%>
<!-- header -->
<a id="top"></a>
<div id="wrapper">
	<div id="header">
		<div id="asglogo"><img src="images/images/logo_asg_3.0.jpg" alt="ASP Stats Generator v<%= ASG_VERSION %>" /></div>
	</div>
</div>
<!-- / header -->

<table width="100%" class="menubar" cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td class="menubar_toolbar"><div id="menubarToolbar">&nbsp;</div>
			<script language="javascript" type="text/javascript">
			<% if blnAsgShowToolbar then %> 
			cmDraw ('menubarToolbar', asgMenu, 'hbr', cmThemeOffice, 'ThemeOffice');
			<% end if %>
			</script>
		</td>
	<% if Session("asgLogin") = "Logged" then %>
	<% if intAsgSetupLock < 1 then %>
		<td class="menubar_toolbar">
		<!-- <div class="errortext" onMouseOver="stm(Warning[0],Style[2])" onMouseOut="htm()">
		<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/warning.png" alt="<%= TXT_Setuplock_off_warning %>" border="0" align="middle" />&nbsp;<%= TXT_Setuplock_off_warning %>
		</div>-->
		</td>
	<% end if %>
		<td class="menubar_toolbar" align="right">
		<a href="login.asp?logout=true" title="<%= TXT_logout_execute %>" class="menubar_toolbar_text" onClick="return confirm('<%= TXT_logout_conf %>');">
		<%= MENUSECTION_Logout %> <strong><%= TXT_administrator %></strong></a> <img src="<%= STR_ASG_SKIN_PATH_IMAGE %>icons/admin.png" alt="<%= TXT_administrator %>" align="middle" />
		</td>
	<% end if %>
	</tr>
</table>

<!-- body -->
<div id="body"><br />