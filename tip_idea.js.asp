<!--#include file="w2k3_write_permission.asp" -->
<!--#include file="lang/TIP_Idea_lang_file.asp" -->

// JavaScript Document
<%

Dim i
for i = 0 to Ubound(TIP_Idea_t)
	Response.Write(vbCrLf & "Idea[""" & i & """]=[""<img src='" & STR_ASG_SKIN_PATH_IMAGE & "icons/tip.png' alt='Tip' border='0' align='middle' />&nbsp;&nbsp;" & TIP_Idea_t(i) & """,""" & TIP_Idea_c(i) & """]")
next

%>