<link href="css/style.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/javascript" src="includes/js/javascript.js"></script>
<script language="JavaScript" type="text/javascript" src="3rdparty/tipmessage/main15.js"></script>
<script language="JavaScript" type="text/javascript" src="3rdparty/jscookmenu/jscookmenu.js"></script>
<script language="JavaScript" type="text/javascript" src="3rdparty/jscookmenu/jscookmenu_theme.js"></script>
<script language="JavaScript" type="text/javascript">
<!--

var asgMenu =
[
	[null,'Home',null,null,null,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/homesite.png" alt="<%= appAsgSiteURL %>" />','<%= TXT_homepage_website %>','<%= appAsgSiteURL %>',null,'<%= TXT_homepage_website %>'],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/main.png" alt="<%= TXT_homepage_stats %>" />','<%= TXT_homepage_stats %>','main.asp',null,'<%= TXT_homepage_stats %>'],
		_cmSplit,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/home.png" alt="Weppos.com Homepage" />','weppos.com','http://www.weppos.com/','_blank','Weppos.com Homepage']
	],
	_cmSplit,
	[null,'<%= MENUGROUP_Main %>',null,null,null,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/main.png" alt="<%= MENUSECTION_Summary %>" />','<%= MENUSECTION_Summary %>','main.asp',null,'<%= MENUSECTION_Summary %>'],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/activeusers.png" alt="" />','<%= MENUSECTION_ActiveUsers %>','active_users.asp',null,'<%= MENUSECTION_ActiveUsers %>'],
		_cmSplit,
	<% if Session("asgLogin") <> "Logged" then %>
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/login.png" alt="<%= MENUSECTION_Login %>" />','<%= MENUSECTION_Login %>','login.asp',null,'<%= MENUSECTION_Login %>']
	<% else %>
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/logout.png" alt="<%= MENUSECTION_Logout %>" />','<%= MENUSECTION_Logout %>','login.asp?logout=true',null,'<%= MENUSECTION_Logout %>']
	<% end if %>
	],
	_cmSplit,
	[null,'<%= MENUGROUP_Visitors %>',null,null,null,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/visitor.png" alt="<%= MENUSECTION_VisitorDetails %>" />','<%= MENUSECTION_VisitorDetails %>','asg_visitor.asp',null,'<%= MENUSECTION_VisitorDetails %>'],
		[null,'<%= MENUSECTION_VisitorSystems %>',null,null,'<%= MENUSECTION_VisitorSystems %>',
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/system.png" alt="<%= MENUSECTION_Systems %>" />','<%= MENUSECTION_Systems %>','asg_system.asp',null,'<%= MENUSECTION_Systems %>'],
			_cmSplit,
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/os.png" alt="<%= MENUSECTION_OS %>" />','<%= MENUSECTION_OS %>','asg_os.asp',null,'<%= MENUSECTION_OS %>'],
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/browser.png" alt="<%= MENUSECTION_Browsers %>" />','<%= MENUSECTION_Browsers %>','asg_browser.asp',null,'<%= MENUSECTION_Browsers %>'],
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/lang.png" alt="<%= MENUSECTION_BrowsersLang %>" />','<%= MENUSECTION_BrowsersLang %>','browser_lang.asp',null,'<%= MENUSECTION_BrowsersLang %>'],
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/coloreso.png" alt="<%= MENUSECTION_Reso %>" />','<%= MENUSECTION_Reso %>','asg_color.asp',null,'<%= MENUSECTION_Reso %>'],
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/coloreso.png" alt="<%= MENUSECTION_Colors %>" />','<%= MENUSECTION_Colors %>','asg_color.asp',null,'<%= MENUSECTION_Colors %>']
		],
		_cmSplit,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/ip.png" alt="<%= MENUSECTION_OS %>" />','<%= MENUSECTION_IpAddresses %>','asg_ip_address.asp',null,'<%= MENUSECTION_IpAddresses %>'],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/world.png" alt="<%= MENUSECTION_Countries %>" />','<%= MENUSECTION_Countries %>','country.asp',null,'<%= MENUSECTION_Countries %>']
	],
	_cmSplit,
	[null,'<%= MENUGROUP_Navigation %>',null,null,null,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/www.png" alt="<%= MENUSECTION_VisitedPages %>" />','<%= MENUSECTION_VisitedPages %>','page.asp',null,'<%= MENUSECTION_VisitedPages %>']
	],
	_cmSplit,
	[null,'<%= MENUGROUP_Reports %>',null,null,null,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/hourview.png" alt="" />','<%= MENUSECTION_HourlyReports %>','stats_hourly.asp',null,'<%= MENUSECTION_HourlyReports %>'],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/monthview.png" alt="" />','<%= MENUSECTION_DailyReports %>','stats_daily.asp',null,'<%= MENUSECTION_DailyReports %>'],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/yearview.png" alt="" />','<%= MENUSECTION_MonthlyReports %>','stats_monthly.asp',null,'<%= MENUSECTION_MonthlyReports %>'],
		[null,'<%= MENUSECTION_YearlyReports %>','stats_yearly.asp',null,'<%= MENUSECTION_YearlyReports %>'],
		_cmSplit,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/monthcal.png" alt="" />','<%= MENUSECTION_MonthlyCalendar %>','stats_monthly_calendar.asp',null,'<%= MENUSECTION_MonthlyCalendar %>']
	],
	_cmSplit,
	[null,'<%= MENUGROUP_Marketing %>',null,null,null,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/referer.png" alt="<%= MENUSECTION_Referers %>" />','<%= MENUSECTION_Referers %>','referer.asp',null,'<%= MENUSECTION_Referers %>'],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/engine.png" alt="<%= MENUSECTION_SearchEngines %>" />','<%= MENUSECTION_SearchEngines %>','search_engine.asp',null,'<%= MENUSECTION_SearchEngines %>'],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/query.png" alt="<%= MENUSECTION_SearchQueries %>" />','<%= MENUSECTION_SearchQueries %>','search_query.asp',null,'<%= MENUSECTION_SearchQueries %>'],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/serp.png" alt="<%= MENUSECTION_Serp %>" />','<%= MENUSECTION_Serp %>','serp.asp',null,'<%= MENUSECTION_Serp %>']
	],
	_cmSplit,
	<% if Session("asgLogin") <> "Logged" AND ASG_ADMINBAR_NOLOGIN then %>
	[null,'<font class="menubar_toolbar_textdisabled"><%= MENUGROUP_Tools %></font>',null,null,null],
	_cmSplit,
	[null,'<font class="menubar_toolbar_textdisabled"><%= MENUGROUP_Administration %></font>',null,null,null],
	<% else %>
	[null,'<%= MENUGROUP_Tools %>',null,null,null,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/serverinfo.png" alt="<%= MENUSECTION_ServerInfo %>" />','<%= MENUSECTION_ServerInfo %>','sysinfo.asp?servinfo=1',null,null],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/server.png" alt="<%= MENUSECTION_ServerVariables %>" />','<%= MENUSECTION_ServerVariables %>','sysinfo.asp?servars=1',null,null]
	],
	_cmSplit,
	[null,'<%= MENUGROUP_Administration %>',null,null,null,
		[null,'<%= MENUSECTION_General %>',null,null,'<%= MENUSECTION_General %>',
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/config.png" alt="<%= MENUSECTION_Config %>" />','<%= MENUSECTION_Config %>','settings_common.asp',null,'<%= MENUSECTION_Config %>'],
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/security.png" alt="<%= MENUSECTION_Security %>" />','<%= MENUSECTION_Security %>','settings_security.asp',null,'<%= MENUSECTION_Security %>'],
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/count_config.png" alt="<%= MENUSECTION_TrackingExclusion %>" />','<%= MENUSECTION_TrackingExclusion %>','settings_exitcount.asp',null,'<%= MENUSECTION_TrackingExclusion %>']
		],
		[null,'<%= MENUSECTION_Maintenance %>',null,null,'<%= MENUSECTION_Maintenance %>',
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/dboptimize.png" alt="<%= MENUSECTION_CompactDatabase %>" />','<%= MENUSECTION_CompactDatabase %>','compact_database.asp',null,'<%= MENUSECTION_CompactDatabase %>'],
			_cmSplit,
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/remove.png" alt="<%= MENUSECTION_BatchDeleteOldData %>" />','<%= MENUSECTION_BatchDeleteOldData %>','batch_delete_old_data.asp',null,'<%= MENUSECTION_BatchDeleteOldData %>']
		],
		[null,'<%= MENUSECTION_Email %>',null,null,'<%= MENUSECTION_Email %>',
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/email_config.png" alt="<%= MENUSECTION_EmailConfig %>" />','<%= MENUSECTION_EmailConfig %>','settings_email.asp',null,'<%= MENUSECTION_EmailConfig %>']
		],
		_cmSplit,
		[null,'<%= MENUSECTION_SetupAndUpdate %>',null,null,'<%= MENUSECTION_SetupAndUpdate %>',
			<% if ASG_SETUPLOCK then %>
			['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/setuplock.png" alt="<%= MENUSECTION_Setuplock %>" />','<%= MENUSECTION_Setuplock %>','setup_lock.asp',null,'<%= MENUSECTION_Setuplock %>'],
			<% end if %>
			[null,'','',null,'']
		]
	],
	<% end if %>
	_cmSplit,
	[null,'?',null,null,null,
		// ['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/help.png" />','<%= MENUSECTION_HelpContents %>',null,null,null],
		<% if INFO_Langset_Common = "it" then %>
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/faq.png" alt="" />','<%= MENUSECTION_OnlineFaqs %>','http://www.weppos.com/asg/it/docs/faq.asp','_blank',null],
		<% else %>
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/faq.png" alt="" />','<%= MENUSECTION_OnlineFaqs %>','http://www.weppos.com/asg/en/docs/faq.asp','_blank',null],
		<% end if %>
		_cmSplit,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/home.png" alt="" />','ASP Stats Generator Website','http://www.weppos.com/asg/','_blank',null],
		<% if INFO_Langset_Common = "it" then %>
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/donations.png" alt="" />','<%= MENUSECTION_MakeADonation %>','http://www.weppos.com/asg/it/donations.asp','_blank',null],
		<% else %>
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/donations.png" alt="" />','<%= MENUSECTION_MakeADonation %>','http://www.weppos.com/asg/en/donations.asp','_blank',null],
		<% end if %>
		_cmSplit,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/update.png" alt="" />','<%= MENUSECTION_CheckForNewVersion %>',null,null,null],
		<% if INFO_Langset_Common = "it" then %>
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/bug.png" alt="" />','<%= MENUSECTION_ReportBug %>','http://www.weppos.com/forum/forum_topics.asp?FID=2','_blank',null],
		<% else %>
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/bug.png" alt="" />','<%= MENUSECTION_ReportBug %>','http://www.weppos.com/forum/forum_topics.asp?FID=12','_blank',null],
		<% end if %>
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/home.png" alt="" />','<%= MENUSECTION_TechnicalSupportForum %>','http://www.weppos.com/forum/','_blank',null],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/email.png" alt="" />','<%= MENUSECTION_Feedback %>','http://www.weppos.com/asg/','_blank',null],
		_cmSplit,
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/license.png" alt="" />','<%= MENUSECTION_LicenseAgreement %>',null,null,null],
		['<img src="<%= STR_ASG_SKIN_PATH_IMAGE %>menu/about.png" alt="" />','<%= MENUSECTION_About %> ASP Stats Generator',null,null,null]
	]

];

//-->
</script>