<!--#include file="w2k3_write_permission.asp" -->
<!--#include file="lang/tip_info_lang_file.asp" -->

// JavaScript Document
<%

Dim i
for i = 0 to Ubound(TIP_Info_t)
	Response.Write(vbCrLf & "Info[""" & i & """]=[""<img src='" & STR_ASG_SKIN_PATH_IMAGE & "icons/message_info.png' alt='Info' border='0' align='middle' />&nbsp;&nbsp;" & TIP_Info_t(i) & """,""" & TIP_Info_c(i) & """]")
next

%>