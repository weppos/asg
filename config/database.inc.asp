<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


if ASG_USE_ACCESS then

	Dim strAsgMapPath
	Dim strAsgMapPathTo
	Dim strAsgMapPathIP
	
	strAsgMapPath = Server.MapPath(ASG_ACCESS_PATH & ASG_ACCESS_DATABASE & ".mdb")
	strAsgMapPathTo = Server.MapPath(ASG_ACCESS_PATH & ASG_ACCESS_DATABASE & ".bak")
	strAsgMapPathIP = Server.MapPath(ASG_IP2C_PATH & ASG_IP2C_DATABASE & ".mdb")

	%><!--#include file="access-connstrings.inc.asp" --><%

elseif ASG_USE_MYSQL then

	%><!--#include file="mysql-connstrings.inc.asp" --><%

end if

%>
