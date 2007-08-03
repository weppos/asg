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
' * @package         ASP Stats Generator
' * @author          Simone Carletti <weppos@weppos.net>
' * @copyright       2003-2007 Simone Carletti, All Rights Reserved
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id: login.asp 13 2007-08-03 13:05:34Z weppos $
' */
 
'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


'-----------------------------------------------------------------------------------------
' Archivio Icone OS
'-----------------------------------------------------------------------------------------
' function :	
' Date: 	14.12.2003 | 14.12.2003
' Comment:	
'-----------------------------------------------------------------------------------------
function IconaOS(ByVal os)
	
	' General transparent .gif for icons impossible to be created
	Const strAsgIconTransparent = "47494638396110001000910000000000FFFFFFFFFFFF00000021F90401000002002C000000001000100000020E948FA9CBED0FA39CB4DA8BB33E05003B"


'							========================================
'---------------------------      	Unable to find an icon			-------------------------------------
'							========================================


	' (Unknown)
	If InStr(1, os, "(unknown)", 1) > 0 Then
		strAsgIconaTemp = strAsgIconTransparent


'							========================================
'---------------------------      		Identified icons			-------------------------------------
'							========================================


	'[SISTEMI PRINCIPALI]
	
	'Microsoft Windows
	ElseIf os = "Microsoft Windows 2003" Then
		strAsgIconaTemp = "47494638396110001000F78D008B8DE5D5C5E1AF7C00E74B00F762059BB2FF598100FFD514FF9946B8C0CCB38818E86A2871AC00D3D3FAAE2F0F71AC02F0C40C8EDC05A2F70990E300989ACB4964F12639E3FFCF00CB9F008987D8AC8F3771AE00B98803FF953CC8C8F7D1D0FC89D504C59800BFC3F9FFD3022E3BDBB3B1EC907F92BDBEF39693CB7E9C51C4C4F4BF9000D5D5FCA3FF00D4D4FBD4CFF97A8897787AE1D3A8007187F5BE3202D9B31AC299A9A48C6C4156EA2C41E5526FF47278E4FF8222DCD3EE9388816EA20FC6C6F7688F2A687FF56C70D6D1C3E1F3C3004459EA5F70EB7F9382728CFA9C61815B77F6E27332D6B0BB95E5075D7EFCAF8B3288A1FDBB8E0BDE6327DE64275561E4908DDBBC8C055A64DC666CD8D0D0FC9AEC07E76A27DADAFE6597009EA0D03D55EC4A57E3C58079495FEC9593DD8F91E6ADB2F28F7A356474EE986B986487FF9C9DECA7423AD1D1FBCF94958EA6FF7EC700B2B1EDC5C5F9CA9D00D2D2FADBDAFE8A8AE37E97FC5067F1495BE7A9A7D7D8D6FAB188217B9261976D9C8F8DDBFF9132B1B2F2FFD20FBFC0F2BD962B6A9B0C415BEFBC8B00689C00FFA154E8B7009B98C54660EFFFFFFF00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000021F9040100008D002C00000000100010000008ED001B096C14A060808108110670C3640A151B2F120E0CC0A5030202031C004890402047814D1624BA38800D00035E1E3CD890A2D19E053C48D2F07392C1960870FA34EAB10010498D3010819030A1D0974644C400589A4649CD084483946864E688111C15C22C35C0C0498B1F483EB4C113A5C01D4616009C31B11400050A5D44182A502009982A0004ACE070E510A1418DD03C29F066498E1D7965143970018A9C4663D4CC10A223AD8010350E085274434BA33C588664A940222F0608074648D123300E993F5632C4C83B07F5053E2706BA68D0808E0A143E342818BE888544811ED694B1032010903A020302003B"
	ElseIf os = "Microsoft Windows CE" Then
		strAsgIconaTemp = "47494638396110001000C41A00CC6633669933FFCC99FF9966CCCC999999CC666699FFFFFF9999996699CCFFCC00FFCC33CCCC66CCCCCC99CCCCFF9933CCCC33FFCCCC66996699CC6666CC3399CC33CCFFCCCCCCFFCC9900FFFFCC00000000000000000000000000000000000021F9040100001A002C000000001000100000058DA0268E64491E91C010994912CF03CCD1E11E433C03C86469070B411489C90083498052A9502422980E40902C9B148428C718CC0881ABB3016C341C8E0641C9744A2E1A47419E30A8251466D6569817EA080C0204840C2D170968750604180A0B901070727406090C8E900A0C7C7E800F108F0B0A04360D70070D058D0AA210642E45A1A30C702E440202112421003B"
	ElseIf os = "Microsoft Windows XP" Then
		strAsgIconaTemp = "47494638396110001000C41A00CC6633669933FFCC99FF9966CCCC999999CC666699FFFFFF9999996699CCFFCC00FFCC33CCCC66CCCCCC99CCCCFF9933CCCC33FFCCCC66996699CC6666CC3399CC33CCFFCCCCCCFFCC9900FFFFCC00000000000000000000000000000000000021F9040100001A002C000000001000100000058DA0268E64491E91C010994912CF03CCD1E11E433C03C86469070B411489C90083498052A9502422980E40902C9B148428C718CC0881ABB3016C341C8E0641C9744A2E1A47419E30A8251466D6569817EA080C0204840C2D170968750604180A0B901070727406090C8E900A0C7C7E800F108F0B0A04360D70070D058D0AA210642E45A1A30C702E440202112421003B"
	ElseIf os = "Microsoft Windows 2000" Then
		strAsgIconaTemp = "47494638396110001000B30A00999999000000FF00000000FF00009900FF00990000009900FFFF0099990000000000000000000000000000000000000021F9040100000A002C000000001000100000045950A900A4BD1850CC95501A208ED5057C5A1018862A4A03288CAB20A81B109F7660B785436E4033F40E850225001A3803045C5228A25409CE68802A49A9B0830002E1B2D0A081C478C3A9AA120976C72DEF7443F6767EAF8800003B"
	'Tutte le versioni NT attualmente stessa icona
	ElseIf InStr(1, os, "Microsoft Windows NT", 1) > 0 Then
		strAsgIconaTemp = "47494638396110001000C40000000033333333666666CCCCCC333300663333CCCC669999990099FF66CC3366CCFFFF3300996600993300006666FFCC00336666FF9966663300FF663300000000000000000000000000000000000000000000000000000000000000000000000021F90401000014002C000000001000100000058020258E646992C3A19EA22100702C88EA210801414880E00400D14060BB35160D42220109520C91C104B7400296CD9161101114AAC9E5AB15314C0B0D090110003A53292F10C860AC69059C178170001E0F0C6F03030A387C10048082230A857B7D0006068C14858F01103F003B402384862F306D3134012F5E070328332C2C21003B"
	ElseIf os = "Microsoft Windows ME" Then
		strAsgIconaTemp = "47494638396110001000D53F00CDE8C59ED38A6BBC4BC2E6B8ACDB9C83C96DD3EECBBEE3B28DCD78A2D59273C258B8E1AADEF4DAF9FDF945AD22F6F8F5EEF8EA9BCE8A57B536E6F5E379C45E9AD183F3F6F1E0F0D990CA7A6BC04EDBEFD496CF82FAFBFA71C35541AE1DB1DDA23EAA1BF3FBF1F7FCF57EC36571BF5AA1D78D8DCA777DC3627FC465A5DA94A8D497EAF7E6F8FEF76EC05364B74566BC4768BC4A92D27A80C3697DC964DAF1D25DB73EE2F2DCE6F4DCD1E8C889C87288C975ECF7E889CD748ED2787BC664FFFFFF21F9040100003F002C000000001000100000069FC09F70482C0A450CA3F2A7C9FC6C88C65278C9EC60A083F030315E2804CCA8F7238002468D0B3609F06808D9460891FE001ECA2F42296810070F1B02390F3F0B1E293F250E181C0A21013A0012173F2A0E003F150E0637273619072805521B3586262F3F03113812261F7627244233313F091F387A441D093F2C2D033F180B136C1C762D9A163249043E0D091D112B42031047422125063F3BDA53E3E44341003B"
	'Tutte le versioni 9x attualmente stessa icona
	ElseIf InStr(1, os, "Microsoft Windows 9", 1) > 0 Then
		strAsgIconaTemp = "47494638396110001000C40000000033333333666666CCCCCC333300663333CCCC669999990099FF66CC3366CCFFFF3300996600993300006666FFCC00336666FF9966663300FF663300000000000000000000000000000000000000000000000000000000000000000000000021F90401000014002C000000001000100000058020258E646992C3A19EA22100702C88EA210801414880E00400D14060BB35160D42220109520C91C104B7400296CD9161101114AAC9E5AB15314C0B0D090110003A53292F10C860AC69059C178170001E0F0C6F03030A387C10048082230A857B7D0006068C14858F01103F003B402384862F306D3134012F5E070328332C2C21003B"
	ElseIf os = "Microsoft Windows 3.1" Then
		strAsgIconaTemp = "47494638396110001000C40000000033333333666666CCCCCC333300663333CCCC669999990099FF66CC3366CCFFFF3300996600993300006666FFCC00336666FF9966663300FF663300000000000000000000000000000000000000000000000000000000000000000000000021F90401000014002C000000001000100000058020258E646992C3A19EA22100702C88EA210801414880E00400D14060BB35160D42220109520C91C104B7400296CD9161101114AAC9E5AB15314C0B0D090110003A53292F10C860AC69059C178170001E0F0C6F03030A387C10048082230A857B7D0006068C14858F01103F003B402384862F306D3134012F5E070328332C2C21003B"
	ElseIf Left(os, 17) = "Microsoft Windows" Then
		strAsgIconaTemp = "47494638396110001000C40000000033333333666666CCCCCC333300663333CCCC669999990099FF66CC3366CCFFFF3300996600993300006666FFCC00336666FF9966663300FF663300000000000000000000000000000000000000000000000000000000000000000000000021F90401000014002C000000001000100000058020258E646992C3A19EA22100702C88EA210801414880E00400D14060BB35160D42220109520C91C104B7400296CD9161101114AAC9E5AB15314C0B0D090110003A53292F10C860AC69059C178170001E0F0C6F03030A387C10048082230A857B7D0006068C14858F01103F003B402384862F306D3134012F5E070328332C2C21003B"
	
	'BeOS
	ElseIf os = "BeOS" Then
		strAsgIconaTemp = "47494638396110001000B30000000000005A8C1818183939393973A55A84B5737373849CC694ADCEA5BDD6CED6E7E7EFF7EF5A52F78C84FFCECEFFFFFF2C0000000010001000000454F0C9438FBC1807B9B3EF9BC3341E170416C93825689025C7158C04634490134FC3D49F93F0E7B865289743A2B4002C148087D3E2112C24864334B67D2C065045A9AB3048BA18E4C32A0197040080A1098FD92F11003B"
	
	'Mac OS
	ElseIf os = "Mac OS 10" Then					' Mac 10 = Mac OS X
		strAsgIconaTemp = "47494638396110001000E65100989EA5969CA3979DA4959BA2959BA39399A0EBEDEE9EA4ABA1A7AD9299A09CA1A8EEEFF09EA3AAF2F3F4E7E9EAA4AAAF939AA1A9AEB49CA2A9D5D8DB9AA0A7AAB0B5ABB0B6A4AAB0D1D4D7D6D9DBF8F8F8E3E5E6A5AAB08B929AE1E2E5EAEAEDE4E5E7D4D7D9ACB1B6CED1D4B2B7BCFCFBFBDFE1E3969CA4979EA590979EC8CBCEA9ADB4BEC2C6DCDEE0A7ACB3BDC1C5B7BBC0C9CCD0999FA69BA0A7F6F7F7E9EAECF5F6F7E8E9EBA3A8AF8F969EDBDCDFB3B8BDCFD2D5B0B4BAFAFAFA959CA3B5B9BEB8BCC0F3F4F59298A0ECEDEEF4F5F6D0D2D5EAEBED969DA4EEF0F1ECEDEFA8ADB2F6F6F6979DA5C2C5CA949BA29DA2A9FFFFFF00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000021F90401000051002C000000001000100000079880518283824A162B0B848A821A241D2F8B8B1B144034915136373542063C0E0E1E0D8A26153F02114E2C384F051719832012090028040510270000393D823E3043BAC2C203084482490801C3C30331831F3202CCC43A830650D3D4000323830D4B04DB00042E45822541B3DB4D013B47821301DAD4022918824C2205E3101C89825ACC58C60C4980108A8C2838C0E0C103060714A8181408003B"
	ElseIf InStr(1, os, "Mac OS", 1) > 0 Then		' Mac
		strAsgIconaTemp = "47494638396110001000E67F008D94B4E5EFF664A2DC4888C8DCE6EE126FCB8AA9C5A2DCFD203A9194B3D33B89D2C7E3FD70A9E38ABCEA255CA67A95C72E77C3C7F1FFBBD2ECF1F2F49DC2E58CC4F11A51A9B4C5DB93C3E491BDE4FBFCFD468FD44457976BA3DA4C96D528549D4D96DCA6E3FC559ADA898EB4B8CDE76A728E6D78A597CCFC78AADEADCEEDC7DEF4D5D7E4A9DDFF91C1ED959BB18DC2F894A6C99CC7F18FABC89BD2FF72A5DA98BFE66AA7DD7BB5EA76B1E81B6CC077A4CD77B4F1EEF1F55E9BD65962889EB4CF5899D97BB0DEC2C5D3F3F7FBC4D1E2496AB4A5CFF69DCBF6ACC1D3D7DFEC76ACDE2E6EC03F77AE3977BC7BB0E28084A79DD8F9518ECABBE3EEBFE7FF92CAE4A2C7E981A6C786AEC58BA5CB5C8BB687C0F587ACD88ECAFF257CCF99D1E693C5F692ACD1A5AFC5ADB1C0AEBDD32C83CF2F83D23F84C4A0C3E55CA1E086B9E383B8EB88B6E15597D64592D74793DB8CB5C5B6C0D6BDC6D94663A2A9ADBE4E7FB2729FD0A7DBEBB2B8CC94D0FFC1CEE0C0CDDED8E1ECB4DBFEB1D1F0D4E5F7FFFFFF21F9040100007F002C00000000100010000007FF807F827F427325872562428383132E3E0077717061003E2E138D00263017241212246030264F997F231C0F7D2967671435122A0F1C007F2B080E41413D1D34356B35692845082B00164B024A1D1B1B19411E640A4B0823724D0340292A401E02361E5D0505161C1F105D364E6D1B0A6C55025D6565391F0E516E22226E20200C7E6A6C10E024870326664434688103C78D850D04A0B111858E9501226E54D0F26241002331D4DCC8904006812C3D76BC38C165C1103E27BE1CC9605203923A3B66B03870600A0B3C336224B8A041038F04342A84881021C4012815B68CE1A12140123D3274509112C08E170C3AB0E4491260C81E22177E18B8F2C6800C193F062E10D9332410003B"
	ElseIf InStr(1, os, "Macintosh", 1) > 0 Then	' Common
		strAsgIconaTemp = "47494638396110001000B30000FFFF00339933FFCC99669966993366FF0000FF66330000FF6633CC99CC99FFCCCC9999FF9966CCFF999900000000000021F9040100000E002C000000001000100000044ED0C949AB9D29846B47189C9488E3D8694192789A397DF0E0C9E004DC786EE53C602982A0506869188E48A4A02228389FD0E54441A85AAF8DCA02C1ED72193FC4614C3E2C2E8A8579C13087DE8E08003B"
	
	'Linux
	ElseIf os = "Linux" Then
		strAsgIconaTemp = "47494638396110001000B30000FFCC33FFFFFF000000999966666666999999CC9933333333666633FFCC99CCCCCC99660000000000000000000000000021F9040100000C002C000000001000100000045D90C949AB65439C726910A0107452801C99424A08606C2B43180542AC83920485A05E0A446098D08C2C04C57098E1540287E5B077B02808529E0051490000D8E50B50F9BA064385E19BA09801830241606E4F020AC0225E282C0C312411003B"
	
	'[CENTRALI]
	
	' 12.11.2004 Solaris							' SUN OS is called Solaris
	ElseIf InStr(1, os, "Sun OS", 1) > 0 Then
		strAsgIconaTemp = "47494638396110001000E66300FFFEFEFFFCFDFFFCFCF2F3F7FEF7F8FFFDFDFEF9FAFDFDFEFBE6E9FEFBFC919ABEC8CCDEFBFBFDFEFEFFFEFEFEE56673F1E0E6F9F9FBE97786FEF8F9ED96A0F7CDD4EE8D98F5C5CCFBEBEEF0F1F6FFFEFF96A4C6CDD2E2F7F6F9374788F4F5F9A2AAC83E5491F9CDD3FADFE3F3B6BFFEFFFFE0495A25387DE14757EE4F5DDC263BE3E5EEE8717DF8F9FBFEFAFBED909CFCFCFD1F32791E3079DADDE9EE8C99FBEFF2D5D8E6EE8D99F8F8FB616EA1FBE5E9E46071F1F2F6F1A7B2FAE1E5808BB4FCEFF2E25766FCE8EBF4CBD1F7F8FAE9EBF26471A354659CF7E8ECC6CADD959EC0F2A6B0D2BECF9FA7C6C1C6DBD6D9E6FADCE2F7F7FAACB2CEA5ADCBFDF4F6E23D4E7C86B1D0D4E3A0ADCB5F6DA0C3C6DBDD364A7C87B1BABED5EC8391EFA2ABED939FF9DCE1FFFFFFFFFFFF00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000021F90401000063002C000000001000100000079480638283848586095001631A058663008A133D066340618E051554003B086224048E6308263E5B2F5F0F63023A2E8262040206412837342A14235E2200AF42122C6017090E354B554386181629820D254810070C0C300ECE1D4C1C204A0A1B58535D5A494F2B0A464721590B573303191F512D3811563231271E3662A263034E4D7EE4E0226541111E440EF45BC8B0A143878100003B"
	' Sun OS
'	ElseIf InStr(1, os, "Sun OS", 1) > 0 Then
'		strAsgIconaTemp = "47494638396110001000910300FFFFFEA6CBF35D5DBEFFFFFF21F90401000003002C000000001000100000023A9C8FA926129B0408139E2085BDD9862751CFA05DC00434A268504D7A4E6A433A18754AF058C3FA8CF8E8500C8BC3B161DC928E4AB2557AF21005003B"

	'OS/2
	ElseIf os = "OS/2" Then
		strAsgIconaTemp = "47494638396110001000E6FF00E7B5B7E1ADB2FBF5F69A6891733E79583F8DD9D5E847419B8185C4C8CDEAF7F9FE4571D6DFE7F9F3F6FD0448D2084BD30C4ED31051D41957D61D5AD6366DDB4779DE507FE05885E15C88E2608BE2618BE3648EE36891E46C94E5749AE67A9EE782A4E981A3E886A7E997B3EC9BB6ED9FB9EEA6BEEFA7BFEFC0D1F4C9D8F5DAE4F8DEE7F9E6EDFBFBFCFEEAF0FBFFEEEAFF5E3EFFAF9FFF2C03FF2F07FF360FFF4420F64527FF6142FF6547FF8068FF8B75FFB2A3FFB5A7FFC3B8FFCDC4D4463FCF4C4AC0C0C000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000021F90401000041002C0000000010001000400778804182838485850A2B0D823932000A85230E2241260E1F160E1A250C8682329F9F389D2C20191F9C2D271C1B2484280E152A3FA03606140E299D413C34302FBAC0C1C2412E13111812112A1EB0171284210E2541090703010B0F0E112C8420100E0408024032050F1DC13E3337C23A3235313B3DC3F5C281003B"
	
	'RiscOS
	ElseIf os = "RiscOS" Then
		strAsgIconaTemp = "47494638396110001000D5000000860183B68416BA1748944800AC0159B85A1A951B91E09158D458D1E3D10094011AA81B58A4587BD87CD1F7D11A9E1B2CC32D1AB21B63A464EFF4EF0D8F0E2C932D59DE5A009E0192C49224BA250EA80F7BBA7C83E6840E950F4EB04E78E1792EAE2EEFFDEF168A178EBA8E4DA04D24942562D863DFE9DF0099012EA52E62A863DFFCDF30993132CA3218B61977AC7832903263DE6430C0317EC97F8EEA8E77E67818901900000069676A69676A69676A69676A69676A69676A69676A69676A21F90401000037002C0000000010001000400696C09B70482CDE6428140C564424510021ED1289C432A8C735591886B0494D8B90341812C49AC37133A118B790C512227E9283D36D524049244535491E6C2B4E7F452F13136D6F7B232346431C041B923753555B5A605D4221104F3624224925258C371C081914147A9F020B0B080D37182A201D1D68420E2EA72A019F602862642866BD37330AA72F2917252F150A0A2F97072CC34641003B"
	
	'Amiga
	ElseIf InStr(1, os, "Amiga", 1) > 0 Then			'(Mancante specifica IRIX 64 [Segnalata] per problemi dimensione icona)
		strAsgIconaTemp = "47494638396110001000C4FF00C0C0C0FFFFCCFFCCCCFF9999FF6699FF6666FF3333FF0000CCFFCCCCCCCCCC99CCCC9999CC6699CC6666CC3366CC3333CC00009999CC99999999666699333399000066666666333366330066000033333333330033000000000000000000000021F90401000000002C000000001000100040059720208A42B234CB384ED252BD8B3451569D711A9018CF1300034783F29A5844974B4D69B9DC3450CE8851280C1E8604E201814C1280D323015E742992B4D39298BD2A32E70DA7918812510E5DCF51A8740902051004090D135F233B3D3F41434547020D07030322065D69710A130B0B3D0D0D2F6869181C1631136F9A19AD191AA700494B4B4F517719351AB6741B117F00125050762A21003B009966FF9966CC9966999966669966339966009933FF9933CC9933999933669933339933009900FF9900CC99009999006699003399000066FFFF66FFCC66FF9966FF6666FF3366FF0066CCFF66CCCC66CC9966CC6666CC3366CC006699FF6699CC6699996699666699336699006666FF6666CC6666996666666666336666006633FF6633CC6633996633666633336633006600FF6600CC66009966006666003366000033FFFF33FFCC33FF9933FF6633FF3333FF0033CCFF33CCCC33CC9933CC6633CC3333CC003399FF3399CC3399993399663399333399003366FF3366CC3366993366663366333366003333FF3333CC3333993333663333333333003300FF3300CC33009933006633003333000000FFFF00FFCC00FF9900FF6600FF3300FF0000CCFF00CCCC00CC9900CC6600CC3300CC000099FF0099CC0099990099660099330099000066FF0066CC0066990066660066330066000033FF0033CC0033990033660033330033000000FF0000CC000099000066000033000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21FE0E4D61646520776974682047494D50002C000000000E000E000008A4000108047060858C1C324A0C1C5862C5810A4728ACC8D165C5421C151C00E1B0A20490231505AEE000240000073F729059D32510C11C231C08E470E488152B5D10C53808C4A28C9A646E227A14A88B0C194072E458B326A89544B38A5A91C154064E3281023D9AC5CA6A17A637AD3C1ACB2AEA8A2E2BD7E01CBA95152B2B0010210A34976EDBB20700AC20CAEA2ED7565506AE283B8B2BE1180B055A71EBD68A42810101003B"
	
	'[UNIX]
	
	' FreeBSD
	ElseIf InStr(1, os, "FreeBSD", 1) Then
		strAsgIconaTemp = "47494638396110001000F700002A262E423A3E725E628E8A8CAEAAAEBEBEBE962E34D1CED16A2E32D8D6D6743C48E9E2E55B2A33B65E62EEEAF09E3A46CA2E2E786E704E2A2E8C2E3BB29296F6F6F73626366A527E925E6AA232329A6E727E3E4A5652527F2E369292968878A0AE7276D2AEAEFAFEFC60464C462E52BEBABACABEC64A262E8C606C7A4E525C3256762E32725A5EA2565AB0322E4E364EDAC2BEBE96969A8E9AA69AA2E8D0D6C6AEAEB2828AAA7E82823A4AD29EA2422A3A867E864E3E6EE7DEE18A525E5A2A426A666A7A6E8AB6AEB4923A3E3E323E6E4652AE424A3A2A3ED0C2CC662C2E4A3442A6928E4E2A326C2E458B3643524A4EB2A2BADED0D2FEFAFA7A324EAA6262D13232C27E7EA69696DACACE4E363A6C5E8ABA9EA22E2E2E62363ECAC6C6F4F2F3362E3AA68A8AB2767E9836409A7E8A843246D2B6BE7A6A7E5E565A7832464E42468E868B92445632262AA63E46FAFAFADACECEAC9CA23B262EBA323AB2323E4E2E3AE0D6DE4A2E467E4A52B6AAAEBAB6B6962E3AD4C6CE722A3692565EAA869A864E669A92AEBEBAD2AA2E36BEAEB6DAC6D22C2A2EEFE2E662566E92303A863232FEFEFABE767C5A4E4E5632423A2A37C53233AAA2A6722E42C5B5B9A2464EE2DEDEBE868E5F2E436A3A5AE2BEBEA4343E9E2E30C2A2AAAE9CAE562E48927A9666365E662E3C4236664E3662A2869EBE929A927E86BEA2B2A696AA762E3A5C2E3592363E66525E9082964242466A323AA6A2A68A4252825A62A26A6AE2C6C2665A5E943E4E8A323B422E46EEDADAD2CAD3864A4EEAE6E67E6E9A4E3E56FAF2F69E32422E262DC5BAC3552B3FBEAAB266324E362A33AE3636B4363D8A3E4AB63E429A4252FEFEFE6E2A3A3E2E4A322933F2EEF13A2A2EC6362E5B2E3BC6BEC2562E2E664656A28E96764A563E2E36AE8A86DADADA7A3236FAF6F7D6CAD3DBC2C7FEFAFE6A32428A6276B8A8B8862E3C864C58C6363AD2D2D2442E3CD8D2D56C2E3B3E2A367E323EA59EA3B9A2A599323D722E36AA3232BDB2B9EEE6EAB032367A2E3C443E44963236A636364A2A303C2E3E863646CEBAC4A6424EE2DADAFFFFFF2C00000000100010004008B100990914B825C1407F0313326BE029C92645D144284C284255B5490FFA0028315120B8371892BC7A612852B78E8B5C8831522F9486555C14825B9307051D3A1DD0CDE898F0569D6000783203F78B8F965E811C087CB34B5EA3848BACA0DB42C60EB31B0750A228D5C356872416F44C1496231E3B462E92045B3411D50E6B1D0E354BA5862DCF450732D01B6648A840116E800D03607022343E50CC9C2A338544B72FF26AB1E0E59799B730714ABC611610003B"
	
	' 12.11.2004 NetBSD
	ElseIf InStr(1, os, "BSD", 1) Then
		strAsgIconaTemp = "47494638396120002000E70000FFFFFFFEFEFEFAFEFCE2DEDEFEFAFEBA9EA260464CFAFAFAE2DADAC27E7E92363E442E3CBDB2B9B2828A9A7E8AA59EA3E0D6DEAA7E82562E2EFEFEFAEEE6EAB63E427F2E368A323B7A3236552B3F4E2E3A3E2E36FAF6F7AE8A865B2A33B0322EB2767EA2464EAE424A7A2E3C6A2E32662E3C9A6E72AA62624A262EEEEAF0F2EEF18A3E4ADED0D2A2565AC2A2AA862E3C962E349E2E30662C2E32262AEFE2E6825A6266525E8A525E927E86D0C2CC9A4252AA2E36722E363E2A362E262D3229338E8A8CBE96965C2E354A3442C5B5B99E3A46722A362A262E443E44BEBABAFEFAFAC532336C2E3BAA869AB032368E868BC5BAC3AA3232BA323A7E323E743C48923A3E99323DF4F2F3CAC6C69A8E9AA63E468C2E3B962E3A863646A4343ED132322E2E2ED8D2D5B6AAAEFAF2F6B8A8B83E2E4AE2C6C2DBC2C7B9A2A58432466E2A3A6C2E45562E483C2E3EA696AA7A6A7E867E86CEBAC4D2CAD3764A56762E3A4E4246908296A6424EA232329632368632323A2A2ED1CED16E46522C2A2EA28E96722E42783246422A3A3A2A37362636BE767CCA2E2E5A2A425A4E4EE9E2E5EEDADA925E6A924456823A4A6A323A66324E4E364E423A3EAAA2A6D8D6D6E7DEE18C606C92565E864C588A42526A3A5A4E3E563E323E786E70EAE6E6BEAEB6D4C6CEA2869ED2B6BEDACACEBEA2B2864E667A324E5C3256462E52665A5E725A5EA6928EDACECED6CAD3AE9CAE8878A06A527E4E36624236664E3E6E62363EF6F6F7C6AEAEB29296AC9CA2B6AEB4E3E3E3AAAAAA727272C7C7C78E8E8E1D1D1D555555393939FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF21FE15437265617465642077697468205468652047494D50002C00000000200020004008FE0001080C40B02041010206185CC850A04300050FA838B024060C26439A38D9C1E3C81301131806783890E1843F8002091A44E80801910549425C4880008731B40E2CEC544B811E217BF61094C910C1870B2042881841A28489130A50A4684892A09B3770E238B12087C61C35743CD409097364559A07142D62D4C8D123487E22499A2452E62F00777FE9F50540972E81BA7A01E8A56BD72E5DBC020FFEFB90E181040A163068E0E0010482112490354874218D1A366EE0C8A183630F1F3F80D43D7B050B8B2C5AB670E9E2E5CB121819C08429DBB920873C6CCAC888B145CF963D02CA9A6D2C3281210519041D39B2793573159E3EC901154AD42852694AFE993AC541C095494F500D49A56AA8CCF77D191F860F80D77C87BD7EF1FA2578177F81FEEDE7DF5FBDF8974B5FBFF8771643041460C042072000536F0551508105165C8041061A6CC0C1011D78609D430CA9B042032CB4E082022F8C00430C32CCA0D342140630411042A020C1104414F18111472091841223961480125040B14414524C4145151F5821830F4054C7D959628C414619669C81C60A69A8B1061B6DBCC4DB59351D7080156CD841C41D7844B98072CB91B8D0157CF48187147464D0831F74D658502117AC31880FD305BAA04108CD01C8218210420822892457E44C0150224718955872092398504147269AF001C7263F70F2402766D659921257ABB042460E71B4E2CA2BA6C0128B2C2A8024C001295C81C62CEE3147A741ACD8A2CA2DB82421074107E2E78B5EF5E9251830C0DCA5CBB4DAEAF5D75E8C39941700F3F1125F7DBAE4C54B2F89E99757B8F4C52B2F000101003B"
	
	' Irix
	ElseIf InStr(1, os, "IRIX", 1) > 0 Then			'(Mancante specifica IRIX 64 [Segnalata] per problemi dimensione icona)
		strAsgIconaTemp = "47494638396110001000B30000567390B2C0CB8B98A5E3E7EA6B859FCDD3DAA0ADBAFAFBFAC0C0C000000000000000000000000000000000000000000021F90401000008002C0000000010001000000475F04859480820CC5D0920454758D2201086711885795CE3D08587D00A4738EA6051DFB9D96060189D4202010060181A431865266458A2A4C594A00AD87E6E4A4B61996A7D92A142B29AF6A60697935770813703A9267D4CCC0FF16F7B007D6F032317350186411E031201238763911A1B8F53199611003B"
	
	'HP-UX
	ElseIf os = "HP-UX" Then
		strAsgIconaTemp = "47494638396110001000C40000F8F6FCEFEAFBBEB5E2CAC3EB9286D4ADA5D8E0DCF63521B84131A94D3FA65A4DAC665BB08379C79F96D4746ABCD6D2F20000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000002C000000001000100000059960218AC3F008638AAC49522802B3ACEB811C87F33CC5620081970287533C00000701892C258846A4920970120B008301E1280409A2062EF11329168B97033D782E020134434170C8188CF841111828140D3009040902020143090F0D720B0D090E7C010F4F070D2525026B0C8701044409034C6D0248869507A1A385A602A87B86008F70027F44342B0D4A81B82B86BF0F030603BF02C40221003B"
	
	'AIX
	ElseIf InStr(1, os, "AIX", 1) > 0 Then
		strAsgIconaTemp = "47494638396110001000B30000E00514C5C5D400097A040F7E00117A0417730C7C40FEB809C0C0C000000000000000000000000000000000000000000021F90401000008002C000000001000100040043F10C949EB0C208373882102115800619AA3759DA97A1DAD8511DB61DCA27C7A5EECFE32E0AAE3A368381F832F40A371788342AC64DA2473BA5DB112E80A8511003B"
	
	'Symbian
	ElseIf os = "Symbian" Then
		strAsgIconaTemp = "47494638396110001000C41D00FFB23EFDFEFEFF9900C9DFE91D77A51472A1B6D4E4FFF5E7CBE1ECFFC36AE9F2F67CB1CCF2F7FA217AA7FFFEFCFFA213FBFDFD4992B6418EB4FFF0D9E2EEF4FFD79B89B9D0B8D6E53284AD03689AFAFCFDDAE9F1006699FFFFFF00000000000021F9040100001D002C000000001000100000053B60278E64698A55E29CE52108005B02CF2497C14D169CA677010EC7F2A308233F8410F3330809BF8BB0F1DB08253F8870F1EB64388A2E63D0BD8500003B"
	
	'[FINALI]
	
	'Unix
	ElseIf os = "UNIX" Then
		strAsgIconaTemp = "47494638396110001000F700000000000808082929293131313939399C9C9CA5A5A5ADADADBDBDBDC6C6C6CECECED6D6D6EFEFEFF7F7F7FFFFFFFFFFFF1010101111111212121313131414141515151616161717171818181919191A1A1A1B1B1B1C1C1C1D1D1D1E1E1E1F1F1F2020202121212222222323232424242525252626262727272828282929292A2A2A2B2B2B2C2C2C2D2D2D2E2E2E2F2F2F3030303131313232323333333434343535353636363737373838383939393A3A3A3B3B3B3C3C3C3D3D3D3E3E3E3F3F3F4040404141414242424343434444444545454646464747474848484949494A4A4A4B4B4B4C4C4C4D4D4D4E4E4E4F4F4F5050505151515252525353535454545555555656565757575858585959595A5A5A5B5B5B5C5C5C5D5D5D5E5E5E5F5F5F6060606161616262626363636464646565656666666767676868686969696A6A6A6B6B6B6C6C6C6D6D6D6E6E6E6F6F6F7070707171717272727373737474747575757676767777777878787979797A7A7A7B7B7B7C7C7C7D7D7D7E7E7E7F7F7F8080808181818282828383838484848585858686868787878888888989898A8A8A8B8B8B8C8C8C8D8D8D8E8E8E8F8F8F9090909191919292929393939494949595959696969797979898989999999A9A9A9B9B9B9C9C9C9D9D9D9E9E9E9F9F9FA0A0A0A1A1A1A2A2A2A3A3A3A4A4A4A5A5A5A6A6A6A7A7A7A8A8A8A9A9A9AAAAAAABABABACACACADADADAEAEAEAFAFAFB0B0B0B1B1B1B2B2B2B3B3B3B4B4B4B5B5B5B6B6B6B7B7B7B8B8B8B9B9B9BABABABBBBBBBCBCBCBDBDBDBEBEBEBFBFBFC0C0C0C1C1C1C2C2C2C3C3C3C4C4C4C5C5C5C6C6C6C7C7C7C8C8C8C9C9C9CACACACBCBCBCCCCCCCDCDCDCECECECFCFCFD0D0D0D1D1D1D2D2D2D3D3D3D4D4D4D5D5D5D6D6D6D7D7D7D8D8D8D9D9D9DADADADBDBDBDCDCDCDDDDDDDEDEDEDFDFDFE0E0E0E1E1E1E2E2E2E3E3E3E4E4E4E5E5E5E6E6E6E7E7E7E8E8E8E9E9E9EAEAEAEBEBEBECECECEDEDEDEEEEEEEFEFEFF0F0F0F1F1F1F2F2F2F3F3F3F4F4F4F5F5F5F6F6F6F7F7F7F8F8F8F9F9F9FAFAFAFBFBFBFCFCFCFDFDFDFEFEFEFFFFFF21F9040100000F002C0000000010001000000863001F081C48B0A0400708132A547870A1438407014894887022C5860A002850B84023C40709013814F931E4C8840D1D9054B832E54A93252B9E8CE9200003850C02A00489500002850608EC4C58C0A70306050218184A5480C4014B993E5C68B02AC180003B"

	' Robots
	ElseIf InStr(1, os, "Robot", 1) > 0 Then
		strAsgIconaTemp = "47494638396110001000E65E00C1BCBB969A96EFEDECE3E7E69DA09D9FA09DCFD1CD92928F8A8A89DEDEDD9D9E9E443D3E8B8C8AF3F2F28C80805656559C9C9C463837E9E9E9EBEAE9BDC1C09EA59EF6F5F4575350BABFBC3F3B31ADA2A2E6E5E46F5D53EEEFF1999998A4A2A1E6E5E6C8C6C5A3A1A1C0C2C0B3AEB0EAEBEAA59C9A9587854C423FF0F1F06C6C6DF0F0EFDDD6D6D4D0D09A9B96B7B0ADC9CACB9697967A7172C0BCBBDEDFDE4F4C46F1F1F1504A3684746F91948F332321ECEEED2B3131ADAFAFEFEDEDA7A7A7FBFCFCD7D7D761534F74605F756D74DAD9D7E7E9E5C6C6C54F4A4CE0E1E0E4E6E38984826C5556F6F7F7E5E4E5605A57C4C6C3EFEBED5A533B694F45B2B2B3534136DBD6D8BEBCB9B7BABA9B9695797371D9D9D998A296E7E7E7FFFFFF00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000021F9040100005E002C0000000010001000000782805E82838485868788898A8B82161A1B8B0D5E46332F3D82090285541F14002743325E182A0C5683035840040B11551C0E4248074B30823412051E592852373A534C211D442C5E5B0A47823117193C35382E2D3E5745362B8320014F155C0F244E824A864D0850065A518B263B5E3F298B39495E10258B23415E22138C825D858100003B"
	
	'[CELLULARI]
		
	' 22.05.2004 SonyEricsson Mobile
	ElseIf InStr(1, os, "SonyEricsson Mobile", 1) > 0 Then
		strAsgIconaTemp = "47494638396110001000E67F006779776BCB7B04AE569C9AA500652577798C00653104451800330300A642005921007939F7F7F7EFEFF7F7F7FFFFF7FFE7E7E7005929EFF7F7009E52D6D7DEB5BEBDE7E7EF21B66BD6D7D6004921A59EA597B3A57B868CCED7D6007131009642F7EFF7DEDFDE42A67352516363617300A64A008E315A61734A8663DEEFCE63716B21B65A296942009E3942966B318663318E5ABDB6BD9CBEADDEE7E7BDC7C6C6CFCE73BE9C6BA6845A9A77ADC7BD00964A008639849E94B5D7BD98D794004908429E735A966B63BE8CEFE7EFE7EFE7E7EFEFADB6B5CEE7BD8CA69C73717B9CBEA5BDE7D639AE7331BE63BDBEC639BE7B085929008642F7FFF7DECFDE73A673185129B5B6BD428E526BCF8CBDBEBD9CCFA5E7DFE76BBE6394CF9494D78C186139738684A1A6A97B86847B8E8CC6C7C6CECFCEB5D7C6CEDFDED6E7D635926B6369736B697BCEC7D6A5DFAD08592184CFADA5CFB5848E948C8E947BBE848C8E9C009E425A716B73AE940886398C9694CECFD6BDD7BDEFEFEF005121FFFFFFFFFFFF21F9040100007F002C00000000100010000007E6807F8283837C0C848816323741390E7C7C0F880D1B2838080855120E0D45878303070A707309427E7E0D7C105282645F786D473E58120FA81216107F203C130129685A440F0C664A1020187C65714C015D3D437E4B2204082C671D18461C1A1B15535B6F2526067D3F387A4E316922481B361709750B0A19043056156C69093A251224F8B020429F3E1E8084C9C20780801D0BA22C3060D0E0843B0342FCD1F0A20994037D321CECF3E1899C3082F890A0E2E58A1B050A3CB4B83026894641144680D9E383CB0A012ED6A8A981688698112A00D8394182030544833AE42950800E0D07840201003B"
		
	' 22.05.2004 Nokia Mobile
	ElseIf InStr(1, os, "Nokia Mobile", 1) > 0 Then
		strAsgIconaTemp = "47494638396110001000E67F000297130000C50198119ADAAB1B9C25059B1A99A9EB0000BB91D7A20000C2666BDC008D06008E023443D30F27CD0003C6FFFFFFD2EFDC7BCE8F73C7825F64D8D1EFD99BB3E70817CAA1A6E8C1E7CCA4B1F5E8F7EBAFBBED0A20CCA1DEB073CA8449BC625C64D5BBC6F1209F3000900622A12A162BCE29A9385158D6EBF6EF93ADE524A3382CA739B2E1BD7586E299B2E39CB5E685D198EDF8F206981648B6579FAEE84D58D4E1F4E60D22CC70BC78DDE5F673C988CDEAD390A3EBD6DCF54CBD64B2E7B0D8F1E093A2F0BFE8C80F9A1C8693E3839CE68B96E58A9BEFCBD8F22648D1E0E7F896D9A75672DCD6F1DF3DB2513548D2F8FFEE8FD69FEAEDFE39AC44EBFCE829A5366075DBB1E1B8B4E0BD9EDDAEF1F7FB7C8DE7787ADF9BB5E588D49B83D49883CB8A87D59B777EE0F3FAFDF1FCF8BFC9F551B55AC6E9CE000EC83849D3CCF4C64454D5E8F7ECE4F7E5C3F1BF98A7EB0D26CA38AB493DAE4EDCE1F69BAAE975CD8B78C986CEECD66167D96066DB6264D98EA0EAD2EEDBD4DBF6FFFFFF21F9040100007F002C00000000100010000007CD807F82838485868784035F12197F7D3B316D7F4E08604C1B7F24733959154472232C7F080B1F0467290034617F130510780C681E027627777F023F6232770065430C3C5A052002117F0B03107F41252B044F7F52337F56547F2D378351405882556B7F6E6F847E463D1A8366484B53425C833D0901012E820A07166AF58370014A3A3CA8F3674C8017381AF818642041911A012ED039F20086031B84E024E8F2E74A02284DD27871C086CC203E07F6FCD9D220C183002A4C3C88334844080C827428D0932709070A2810210A04003B"


'							========================================
'---------------------------      Debug immagini non verificate		-------------------------------------
'							========================================


	'Controllo icone non in definizione
	ElseIf os = "checkicon" Then
	Const strAsgCheckIconLink = "http://www.weppos.com/asg/checkversion/check_icon.asp?"
	
	'Passa valori e pulisci stringa
	Dim strAsgIconOs
	strAsgIconOs = Session("strAsgIconOs")
	Session("strAsgIconOs") = ""

		If Len(strAsgIconOs) > 0 Then
			'Passa i valori per gestire una cache di elaborazione che non consenta doppioni
			Session("blnAsgIconOs" & Request.QueryString("page")) = "notified"
			'o il server si mette in ferie!
			Response.Redirect(strAsgCheckIconLink & "iconos=" & strAsgIconOs & "&host=" & Server.URLEncode(Request.ServerVariables("HTTP_HOST")))
		Else
			strAsgIconaTemp = "4749463839610D000E00B30000FFFFFFF2F2F2EBEBEBE5E5E5DFDFDFD8D8D8D2D2D2CBCBCBC5C5C5BFBFBFB2B2B2A5A5A59F9F9F99999900000000000021F904041400FF002C000000000D000E0000043710C849AB9DC1DC2958DB40E17D57D19C45B86892791E00712AEDD924D2712EA18DE736848FA2B39D7E13991139192C3789136B23088022003B"
		End If


'							========================================
'---------------------------				GENERICO				-------------------------------------
'							========================================


	Else
		strAsgIconaTemp = "47494638396110001000C40000003FB4A0ADD69999993755AB0080DAFFFFFF0848FF7896ED4374FD004CDA0099FF3366FF2860FFC0C5D60077FE6680CA005DDD0066FF29479B8BA6F4AABEFD6A8FF94A62AC0076F04263D21849D7C9D3F40089FE003BC3A4B8F5008EF2FFFFFF21F9040100001F002C0000000010001000000581E0278E42298C28291C156394A9285054706059129DE8AC69818290A3E37D04138AA6216C72203BD9A4C3144A9A89CB4980A84C8256ECC5515A740F8787A559805C3625838141CF38DD1BB820229727866E0E795B11850609000081833285857F6D6F1B0A46020E0E111042041E0A942996179A059C9F312517049C30312AAC2821003B"
		
		'Inserisci i dati relativi alle icone sconosciute
		Session("strAsgIconOs") = Session("strAsgIconOs") & os & "|"
	
	End If

	IconaOS = strAsgIconaTemp

end function '							========================================
'---------------------------    NON MODIFICARE IL CODICE SEGUENTE	-------------------------------------
'							========================================


'-----------------------------------------------------------------------------------------
' Stampa Icona
'-----------------------------------------------------------------------------------------
' Funzione: 
' Date: 	13.12.2003 | 13.12.2003
' Comment:	
'-----------------------------------------------------------------------------------------

function StampaIcona(ByVal tempIcona)
	
	Response.ContentType = "Image/gif"
	For Index = 1 To Len(tempIcona) step 2
		Response.BinaryWrite(ChrB("&h" & Mid(tempIcona,Index,2)))
	Next
	
end function 'Richiama le funzioni
strAsgIconaTemp = IconaOS(Request.QueryString("icon"))
StampaIcona(strAsgIconaTemp)

%>