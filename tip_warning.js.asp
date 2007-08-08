<!--#include file="w2k3_write_permission.asp" -->
<!--#include file="lang/tip_warning_lang_file.asp" -->

// JavaScript Document
<%

Dim i
for i = 0 to Ubound(TIP_Warning_t)
	Response.Write(vbCrLf & "Warning[""" & i & """]=[""<img src='" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_warning.png' alt='Warning' border='0' align='middle' />&nbsp;&nbsp;" & TIP_Warning_t(i) & """,""" & TIP_Warning_c(i) & """]")
next

%>