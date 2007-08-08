<!--#include file="includes/templates/default/skin.tmp3.asp" -->
<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'

%>


/* -------------------------------	
			Layout
------------------------------	*/

.table_cont_no_record {
	background-color: <%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>;
	background-image: url(<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>);
	text-align: center;
	height: 20px;
	vertical-align: middle;
}

.table_bar {
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 10pt;
	color: #0066CC;
	font-weight: bold;
	font-variant: small-caps;
	text-align: center;
	background-color: <%= STR_ASG_SKIN_TABLE_BAR_BGCOLOUR %>;
	background-image: url(<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_BAR_BGIMAGE %>);
	border: 1px solid <%= STR_ASG_SKIN_TABLE_LAYOUT_BDCOLOUR %>;
	padding: 2px;
/*	height: 20px;*/
}

.table_menubar {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 7pt;
	color: #0066CC;
	font-weight: bold;
	font-variant: small-caps;
/*	text-align: center;*/
	background-color: <%= STR_ASG_SKIN_TABLE_BAR_BGCOLOUR %>;
	background-image: url(<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_BAR_BGIMAGE %>);
	border: 1px solid <%= STR_ASG_SKIN_TABLE_LAYOUT_BDCOLOUR %>;
	padding: 4px;
/*	height: 20px;*/
}

.table_menubar_page {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	background-color: <%= STR_ASG_SKIN_TABLE_LAYOUT_BGCOLOUR %>; /* <%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %> */
	border-bottom: 1px solid <%= STR_ASG_SKIN_TABLE_LAYOUT_BDCOLOUR %>;
	border-left: 1px solid <%= STR_ASG_SKIN_TABLE_LAYOUT_BDCOLOUR %>;
	border-right: 1px solid <%= STR_ASG_SKIN_TABLE_LAYOUT_BDCOLOUR %>;
	padding: 4px;
}

.table_layout {
	text-align: center;
	background-color: <%= STR_ASG_SKIN_TABLE_LAYOUT_BGCOLOUR %>;
	background-image: url(<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_LAYOUT_BGIMAGE %>);
	border-left: 1px solid <%= STR_ASG_SKIN_TABLE_LAYOUT_BDCOLOUR %>;
	border-right: 1px solid <%= STR_ASG_SKIN_TABLE_LAYOUT_BDCOLOUR %>;
	font-size: 7pt;
	color: #000000;
	padding: 15px;
/*	height: 20px;*/
}

.table_cont_row {
	background-color: <%= STR_ASG_SKIN_TABLE_CONT_BGCOLOUR %>;
	background-image: url(<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_CONT_BGIMAGE %>);
}

.table_title_row {
	background-color: <%= STR_ASG_SKIN_TABLE_TITLE_BGCOLOUR %>;
	background-image: url(<%= STR_ASG_SKIN_PATH_IMAGE & STR_ASG_SKIN_TABLE_TITLE_BGIMAGE %>);
	font-weight: bold;
	font-variant: small-caps;
	font-family: Tahoma, Arial, Helvetica, sans-serif;
	font-size: 9pt;
	height: 18px;
}