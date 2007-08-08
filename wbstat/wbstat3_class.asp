<%
'####################################################################################################
'### WBSTAT statistic class ###
'#  imente.free.software (http://www.imente.it)
'#  Classe: WBstat v3.1
'#  Autore: Simone Cingano (http://www.imente.it)
'#  Email:  freesoftware@imente.it
'#  Data:   lunedì 03/06/2004
'#  Collaboratori: Baol74 (http://www.aspxnet.it) e Weppos (http://www.weppos.com)
'####################################################################################################

'#  INFORMAZIONI E SITO DI RIFERIMENTO:
'#  visitate il sito di WBstat per informazioni e aggiornamenti
'#  http://www.imente.it/wbstat
'####################################################################################################

'#  COPYLEFT:
'#  tutto il codice è completamente FREE. questo significa che è utilizzabile e modificabile
'#  senza il consenso diretto dell'autore. in ogni caso una menzione sarebbe gradita.
'#  in caso di modifiche di miglioria sarebbe interessante aprire un contatto per una possibile
'#  collaborazione, o per la semplice visione delle modifiche e per una possibile ufficializzazione
'####################################################################################################

'#  FILE WBSTAT:
'#  
'# + WBSTAT3
'#	wbstat3_class.asp			questo file. il motore del sistema wbstat
'#	+ WBSTAT3_SAMPLES/
'#		wbstat3_multi.asp		esempio di multielaborazione
'#		wbstat3_multi_list.txt		file:  riferimento dal quale opera wbstat3_multi.asp
'#		wbstat3_simple.asp		esempio basilare di utilizzo della classe	
'#	+ WBSTAT3_SPEC/
'#		wbstat3_browser_internet.xml	file:  specifiche XML (Browser internet)
'#		wbstat3_browser_robot.xml	file:  specifiche XML (Robot e Spider)
'#		wbstat3_browser_multimedia.xml	file:  specifiche XML (Lettori Multimediali)
'#		wbstat3_browser_wap.xml		file:  specifiche XML (WAP)
'#		wbstat3_language.xml		file:  specifiche XML (lingue)
'#		wbstat3_os.xml			file:  specifiche XML (sistemi operativi)
'#		wbstat3_os_windows.xml		file:  specifiche XML (sistemi windows)
'#		wbstat3_createfunc.asp		prog: scrive la funzione di richiamo per wbstat
'#
'#		[NELLA VERSIONE DEVELOPMENT sono compresi anche i seguenti file]
'#		wbstat3_spec.mdb		file Access con tutti i dati
'#		wbstat3_updatexml.asp		il programma che crea i file XML dal database
'#		wbstat3_listdatabase.asp	il programma elenca tutti i dati presenti nel DB
'#		wbstat3_listdatabase.xsl	[foglio di stile per il risultato XML del file asp]
'####################################################################################################

'# USARE LA CLASSE:
'#
'# esempio minimo di codice
'#		
'#		< !--#include file="wbstat3/wbstat3_class.asp"-->
'#		dim oBrowser
'#		Set oBrowser = CreateWBstat("wbstat3/wbstat3_spec/",True,"Sconosciuto",False)
'#		oBrowser.Debug "Key",false
'#
'#	PER GENERARE QUESTA STRINGA SI CONSIGLIA L'UTILIZZO DEL GENERATORE AUTOMATICO
'#	che trovate nella cartella WBSTAT3_SPEC con il nome di wbstat3_createfunc.asp
'#
'#	dopo l'inclusione della classe all'interno di un file ASP
'#	la funzione CREATEWBSTAT imposta su una variabile ancora non definita un oggetto wbstatclass
'#	con impostati tutti gli oggetti relativi (le sottoclassi OPTION e tutti gli oggetti di classe)
'#
'#	*** VEDIAMO I PARAMETRI DI FUNZIONE *********************************************************
'#	********* !!!!!  INDICANDO "" VERRANNO UTILIZZATI I PARAMETRI DI DEFAULT !!!!! **************
'#
'#	PATH	[str]	la cartella dove si trovano i file XML, ovvero le specifiche
'#			rispetto al file dove è inclusa la classe
'#	CACHE	[bol]	ogni volta che viene aperta una sessione vengono salvati i dati in una
'#			variabile di sessione. se vengono richiesti nuovamente, e la sessione è
'#			ancora aperta, i dati non vengono ricalcolati ma presi dalla variabile
'#	FILL	[str]	indica il testo da utilizzare quando un item è vuoto (i.e. "Sconosciuto")
'#	BRLEV	[int]	livello di versione per browser
'#			indica il livello da indicare nel riepilogo (item "BROWSER")
'#			MAJOR = 0| MAJOR+MINOR = 1 | ALL = 2
'#	OSLEV	[int]	livello di versione per os
'#			indica il livello da indicare nel riepilogo (item "OS")
'#			MAJOR = 0| MAJOR+MINOR = 1 | ALL = 2
'#	BRMINSIMPLE [bol] indica se nel riepilogo deve indicare solo il primo numero della minorversion
'#			serve soprattutto con Internet Explorer (la 5.21 diventerà 5.2)
'#	UAINCLUDE[bol]	inclusione negli items della USER AGENT sotto il nome "UserAgent"
'#	DEBUG	[bol]	indicare debug true per poter passare parametri fittizzi alla classe
'#	GUESS	[bol]	indicare true per attivare le funzioni di guessing
'#			se non viene identificato alcun browser (tramite i confronti con i file XML
'#			di specifiche) la classe tenta di identificare ugualmente tramite un
'#			algoritmo di guessing che estrapola dati dalla stringa UA. se si lascia
'#			false rimarrà "Sconosciuto" in questi casi
'#	ROBOT	[bol]	se true esegue il confronto browser anche per i robot, altrimenti li ignora
'#			e li considera sconosciuti
'#	ROBOTPREC [bol]	se true in riepilogo (item "BROWSER") indica "Robot: Google", altrimenti
'#			"Google", nel caso in cui sia stato identificato un robot
'#	WORKBR	[bol]	 determina l'attivazione delle funzioni di identificazione Browser
'#	   WORKBRVERSION versione del browser
'#	   WORKBRDETAILS dettagli del browser (sottotipo, engine...)
'#	   WORKBRACTCAP	 specifiche del browser (supporto)
'#	  WORKBRLANGUAGE lingua del browser
'#	WORKOS		 determina l'attivazione delle funzioni di identificazione OS
'#	  WORKOSVERSION	 versione dell'os
'#	  WORKOSDETAILS	 dettagli sull'os (architettura, sottotipo...)
'#	WORKSPECIAL	 determina l'attivazione delle funzioni: Framework, Mozilla e Gecko Version
'#
'#	*********************************************************************************************	
'#
'#	response.write oBrowser("OS.Version.Minor")
'#
'#	ad esempio in questo modo si recupera il dato versione minore del sistema operativo calcolato
'#	dalla classe
'#	tutti gli oggetti della classe sono recuperabili tramite la funzione DEBUG (indicando come
'#	parametri ORDINAMENTO (KEY o VALUE) e la CHIUSURA (visualizzazione con div nascondibili per
'#	liste di grandi dimensioni...)
'#
'#	in ogni caso mostriamo qui una lista di tutti gli oggetti della classe
'#
'#		Browser			  nome risolutivo (NOME + VERSIONE)
'#		Browser.Name		  nome del browser  
'#		Browser.Version		  versione (completa)
'#		Browser.Version.Major	  versione (solo major)
'#		Browser.Version.Minor	  versione (solo major)
'#		Browser.Version.Rest	  versione (solo build+rest)
'#		Browser.Type		  tipo di browser (numerico)
'#		Browser.Type.Description  definizione tipo di browser (testuale)
'#		Browser.SubType		  sottotipo di browser (numerico)
'#		Browser.SubType.Description  definizione sottotipo di browser (testuale)
'#		Browser.Language	  NOME lingua del browser 	[ITALIANO, INGLESE...]
'#		Browser.Language.Code	  CODICE lingua del browser 	[it, en-us...]
'#		Browser.Security  
'#		Browser.Url
'#		Browser.Engine  	motore di rendering (IE, Mozilla...)
'#		Browser.Guessed		true se il browser è stato "indovinato" poichè sconosciuto
'#
'#		### GLI ACT SONO DATI CERTI (di supporto del browser)
'#		Browser.Act.Cookie	supporto cookie
'#		Browser.Act.Css		versione css supportata (0 per non supportata)
'#		Browser.Act.Frames	supporto frames
'#		Browser.Act.iFrames 	supporto iframes
'#		Browser.Act.Tables 	supporto tabelle
'#		Browser.Act.Wap		browser di tipo WAP (vedi anche browser.type)
'#
'#		### I CAP SON DATI POSSIBILI (poichè compatibile) MA NON CERTI (poichè disattivabili)
'#		Browser.Cap.Activex	supporto ocx
'#		Browser.Cap.Applett	supporto java applett
'#		Browser.Cap.Js		supporto javascript
'#		Browser.Cap.Sound	supporto sonoro
'#		Browser.Cap.Vb		supporto vbscript
'#
'#		Framework.Version 	versione del framework.net 	(0 se assente)
'#		Gecko.Version  		versione del motore gecko 	(0 se assente)
'#		Mozilla.Version 	versione di mozilla		(0 se assente)
'#
'#		OS			nome risolutivo (NOME + VERSIONE)
'#		OS.Name			nome del sistema operativo
'#		OS.Subtype		tipologia (ppc, NT, 9x...)
'#		OS.Type			tipo di OS (numerico)
'#		OS.Type.Description	definizione tipo di OS (testuale)					
'#		OS.Arch			architettura del sistema (Mac ppc, mac 68k, i386, sparc...)
'#		OS.Version		versione (completa)
'#		OS.Version.Major	versione (solo major)
'#		OS.Version.Minor	versione (solo major)
'#		OS.Version.Rest		versione (solo build+rest)
'#
'#		WBstat.cachedata	TRUE = dati in cache | FALSE = dati elaborati sul momento
'#		WBstat.Version		versione di WBstat
'#		
'####################################################################################################

'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

' ********************************* COSTANTI ********************************************************
'### NON MODIFICARE #################################################################################
'# - costanti di identificazione
'#   -- caratteri divisori	>>> qualcosa; qualcosa | (qualcosa,qualcosa) | ...
	Const CST_CHAR_VERSION = 	"(?:[ ]|[/]|[-]|[=]|[_])" 
'#   -- caratteri di versione	>>> SunOS 4 | Netscape/7.2 | linux-2.4.20 | qtver=5.0 ...					
	Const CST_CHAR_IDVERSION = 	"((\d+)(?:\.)?(\d+)?(?:\.)?((?:[-]|[.]|\w)+)?)"
'#   -- caratteri di spaziatura
	Const CST_CHAR_SPACE = 		" ;,/()"
'# - file XML di specifiche
	Const CST_XML_BROWSER =		"wbstat3_browser.xml"	 'browser
	Const CST_XML_LANGUAGE = 	"wbstat3_language.xml"	 'lingue
	Const CST_XML_OS = 		"wbstat3_os.xml"	 'sistemi operativi
	Const CST_XML_OS_WINDOWS =	"wbstat3_os_windows.xml" 'sistemi windows
'# - INFO
	Const CST_WBSTAT_INFO = _
	"WBSTAT/3.1 AUTH:Simone Cingano VER:Original WEB:http://www.imente.it/wbstat"
	Const CST_WBSTAT_VERSION = "3.1"
'####################################################################################################
' ***************************************************************************************************

'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

Class wbstatclass_Options
	Public FillResult
	Public BrVersionPrecision
	Public BrVersionMinorSimple
	Public OsVersionPrecision
	Public SessionID
	Public Cache
	Public IncludeUserAgent
	Public Debugging
	Public Guessing
	Public Robotknown, Robotprecision
	Public workBrowser,workBrowserversion,workBrowserdetails,workBrowserlanguage
	Public workBrowserACTCAP,workOS,workOSversion,workOSdetails,workSpecial
	Private pPath
	

	Private Sub Class_Initialize()
		'impostazioni di base
		SessionID		= "__wbstat3_bros_info__"
		Cache			= False
		FillResult		= ""
		Path			= "wbstat3/wbstat3_spec/"
		BrVersionPrecision	= 1
		BrVersionMinorSimple	= False
		OsVersionPrecision	= 0
		IncludeUserAgent	= True
		Debugging		= False
		Guessing		= True
		Robotknown		= True
		Robotprecision		= False
		'***
		workBrowser = True
		   workBrowserversion = True
		   workBrowserdetails = True
		   workBrowserlanguage = True
		   workBrowserACTCAP = True
		workOS = True
		   workOSversion = True
		   workOSdetails = True
		workSpecial = True
		'***
	End Sub

	Public Property Let Path(value)
		on error resume next
			pPath = Server.MapPath(value) & "\"
			if err.number<>0 then pPath = value & "\"
		on error goto 0
	End Property

	Public Property Get Path()
		Path = pPath
	End Property
End Class

'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§

Class wbstatclass
	Public Items
	Public Options

	Private Reg
	Private Data

	Private pUserAgent
	Private pBrowserPattern
	Private pOSPattern
	Private pSupport
	Private pGuessed
	Private pAutoOS
	Private pAutoUrl

	Private Sub Class_Initialize()
		Set Reg = New RegExp 	'inizializzo la regexp
		Reg.Global = True
		Reg.Ignorecase = True
		'***
		Set Items = CreateObject("Scripting.Dictionary")
		Items.CompareMode = 1
		'***
		Set Options = new wbstatclass_Options
		'***
		Set Data = Server.CreateObject("ADODB.Recordset")
		Data.LockType = 4
	End Sub

	public default property get Item(key)
		Item = Items(key)
	end property

	Public Property Get Version()
		Version = CST_WBSTAT_VERSION
	End Property
	
' ***************************************************************************************************
' ******************************* FUNZIONI GENERALI *************************************************


	Private Function IIF(Cond,CondTrue,CondFalse)	'funzione IIF
		If Cond=True then IIF =	CondTrue : else : IIF=CondFalse : end if
	end function

	Private Function ToBool(value)			'funzione TRUE/FALSE
		ToBool = IIF(value=1,True,False)
	End Function

	Public Sub SetUserAgent(Value)
		pUSerAgent = Value
	End Sub

	Public Sub SetPath(Path)
		Options.Path = Path
	End Sub

	Private Function Fill(Value,IsString,Doit)
		Fill = Trim(Value)
		If Value="" and Doit then Fill = IIF(IsString,trim(options.FillResult),"0")
	End Function

	Private Sub  SetItem(Item,Value,IsString,Doit)
		Items(Item) = Fill(Value,IsString,Doit)
	End Sub
	
	'estrae da una stringa URL solo il nome di dominio
	' E.G. >>> http://www.imente.it  >>> "imente"
	'      >>> http://ciao.com	 >>> "ciao"
	Private function domainname(domain)
	dim tmp,Matches
		tmp = ""
		Reg.Pattern = "http\:\/\/(?:www.)?([^.]+)\."
		if Reg.Test(domain) then
		    set Matches = reg.Execute(domain)
		    tmp = Matches(0).SubMatches(0)
		end if
	
	domainname = tmp	
	end function
	
' ******************************* FUNZIONI BROWSER/OS ***********************************************

	Private Function DecodeType(number)
	Dim pDecodeType
		if isnumeric(number) = false then number = 0
		pDecodeType = Array("Internet Browser","Robot")
		DecodeType	= pDecodeType(number)
	End Function

	Private Function DecodeSubtype(number)
	Dim pDecodeType
		if isnumeric(number) = false then number = 8
		select case number
			case 0
			pDecodeType = "Spider"
			case 1
			pDecodeType = "Validator"
			case 2
			pDecodeType = "Offline Browser"
			case 3
			pDecodeType = "Downloader"
			case 9
			pDecodeType = "Undefined Robot"
			case 10
			pDecodeType = "Internet Browser"
			case 15
			pDecodeType = "Multimedia"
			case 20
			pDecodeType = "WAP Browser"
			case else '8
			pDecodeType = "Other"
		end select
	DecodeSubtype = pDecodeType
	End Function

' ***************************************************************************************************

	Public Sub Eval()	'sub principale
	
	If Options.Cache And Session(Options.SessionID) <> "" then
		LoadCache()
	else
		Items.RemoveAll()
		pSupport = "" : pOSPattern = "" : pBrowserPattern ="" : pAutoOS = ""
		
		if pUserAgent = "" and Options.Debugging then pUserAgent = request.querystring("UA")
		If pUserAgent = "" then pUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
		If Options.IncludeUserAgent then
			SetItem "UserAgent" ,pUserAgent ,True, True
		end if
		if options.workSpecial then
			SetItem "Mozilla.Version"	,Mozilla_Version()	,False,	True
			SetItem "Gecko.Version"		,Gecko_Version()	,False,	True
			SetItem "Framework.Version"	,Frameworkdotnet()	,False,	True
		end if
		if options.workOS then
			SetItem "OS.Name"		,OS_Name()		,True,	True
		end if
		if options.workBrowser then
			SetItem "Browser.Name"		,Browser_Name()		,True,	True
		end if
		if options.workBrowser then
			SetItem "Browser.Url"		,Browser_url()		,True,	False
		end if
		if options.workOS then
			if (Items("OS.Name") = "" or Items("OS.Name") = trim(options.FillResult)) _
			and pAutoOS <> "" then Items("OS.Name") = pAutoOS
			SetItem "OS.Subtype"		,OS_SubType()		,True,	False
		end if
		if options.workBrowser and options.workBrowserVersion then
			SetItem "Browser.Version"	,Browser_Version()	,False,	False
			if Items("Browser.Guessed") then Guess_Version()
		end if
		if options.workOS and options.workOSdetails then
			SetItem "OS.Arch"		,OS_arch()		,True,	False
		end if
		if options.workOS and options.workOSversion then
			SetItem "OS.Version"		,OS_version()		,False,	False
		end if	
		if options.workBrowser and options.workBrowserLanguage then
			SetItem "Browser.Language"	,Browser_language()	,True,	True
		end if
		if options.workBrowser then
			SetItem "Browser"		,Browser_Resolve()	,True,	True
		end if
		if options.workOS then
			SetItem "OS"			,OS_Resolve()		,True,	True
		end if
		if options.workBrowser and options.workBrowserDetails then
			SetItem "Browser.Security"	,Browser_Security()	,True,	False
			Items("Browser.Engine") = _
		      IIF(Items("Browser.Engine") = "",Items("Browser.Name"),Items("Browser.Engine"))
		end if
		if options.workBrowser and options.workBrowserActCap then
			Items("Browser.Act.Wap") = IIF(Items("Browser.Subtype")=20,true,false)
			Browser_Act_Cookie()
			Browser_ActAndCap()
		end if
		Items("WBstat.info")	=	CST_WBSTAT_INFO
		If Options.Cache then WriteCache()
		Items("WBstat.cachedata") = False
	End If
	End Sub
	
' ***************************************************************************************************

	Public Sub DeleteCache()
		Session.Contents.Remove(Options.SessionID)
	End Sub

	Private Sub LoadCache()
	Dim Values
		Values = Session(Options.SessionID)
		values = Replace(Values,"=1|","=cbool(1)|")
		values = Replace(Values,"=0|","=cbool(0)|")
		Values = Replace(Values,"@","Items(""")
		Values = Replace(Values,"=",""")=")
		Values = Replace(Values,"|",VbCrLF)
		'response.write replace(Values,vbcrlf,"</br>")
		Execute Values
		Items("WBstat.cachedata") = True
	End Sub

	Private Sub WriteCache()
	Dim Res,Elm,thisval
		Res = ""
		For Each Elm in Items.Keys
			if  Items(Elm) <> "" then
				thisval = ""
				if items(elm) = true and thisval = "" then thisval = 1
				if items(elm) = false and thisval = "" then thisval = 0
				if thisval = "" then thisval = """" & Items(Elm) & """"
				Res = Res & "@" &  Elm & "=" & thisval & "|"
			end if
		Next
		Session(Options.SessionID) = Res
	End Sub
	
' ***************************************************************************************************

	Public Sub Debug(Sort,OpenCloseID)
	Dim Elm,Value,bgcolor,Data,Attrib,Id,Title,isimportant
		Attrib = ""
		if pUserAgent = "" then pUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
		Title = pUserAgent
		If OpenCloseID or OpenCloseID > 0 then
			If OpenCloseID=True then OpenCloseID=0
			id = "TableInfo_" & OpenCloseID
			Attrib = "ID='" & Id & "' Style='display:none'"
			Title = "<span style=""cursor:hand"" onclick=""javascript:var obj=" & id & _
			".style;obj.display=(obj.display==''?'none':'')""> + " & Title & "</span>"
		end if
		Set Data = Server.CreateObject ("ADODB.Recordset")
		Data.CursorLocation = 3
		Data.Fields.Append "Key",200,255
		Data.Fields.Append "Value",200,255
		Data.Open()
		For Each Elm in Items.Keys
			Data.AddNew
			Data("Key") = Elm
			Value = Items(Elm)
			Data("Value") = IIF(IsNull(Value),"",Value)
			Data.Update()
		Next
		If Sort <> "" then Data.Sort = Sort
		Data.MoveFirst
		Response.Write("<table width=""100%""cellspacing=""0"" cellpadding=""0"">")
		Response.Write("<tr><td style=""text-align=center"">")
		Response.Write("<table width=""70%"" cellspacing=""0"" cellpadding=""0""><tr><td>")
		Response.Write("<table width=""100%"" Border=0")
		Response.Write(" style=""font-size:11px;font-family:verdana,arial;"" ")
		Response.Write("bgColor=""silver"" cellspacing=""1"" cellpadding=""2"">")
		Response.Write("<tr style=""background-color:#DEE2F3;color:Navy"">")
		Response.Write("<td colspan=""2"" ><b>" & Title & "</b></td></tr>")
		Response.Write("</table>")
		Response.Write("<table " & Attrib & " width=""100%"" Border=0")
		Response.Write(" style=""font-size:11px;font-family:verdana,arial;""")
		Response.Write(" bgColor=""silver"" cellspacing=""1"" cellpadding=""2"">")
		Response.Write("<tr style=""background-color:#F3F0DE;color:Navy"">")
		Response.Write("<Td width=""15%"">Key</td><td>Value</td></tr>")
		While Not Data.EOF
			If Value="" or IsNull(Value) then value ="&nbsp;"
			select case Data("Key")
				case "OS", "Browser","UserAgent"
				bgcolor = "#E9DECD"
				isimportant = true
				case "WBstat.info","WBstat.cachedata"
				bgcolor = "#E6EAF9"
				isimportant = true
				case else
				bgcolor = IIF(bgcolor="#F4F5E9","#F4F5E4","#F4F5E9")
				isimportant = false
			end select
			Response.Write("<tr style=""background-color:" & bgcolor & """><td>")
			Response.Write(IIF(isimportant,"<span style=""font-weight:bold"">" & _
				Data("Key") & "</span>",Data("Key")) & "</td><td>")
			Response.Write(IIF(isimportant,"<span style=""font-weight:bold"">" & _
				Data("Value") & "</span>",Data("Value")) & "</td></tr>")
			Data.MoveNext()
		Wend
		Response.Write("</table>")
		Response.Write("</td></tr></table>")
		Response.Write("</td></tr></table>")
		Data.Close()
		Set Data = Nothing
	End Sub
	
' ***************************************************************************************************

	'************************************ Browser	*********************************************

	Private Function GetVersion(Item)
	Dim VersionPrecision,Res,Temp
		VersionPrecision = _
			IIF(Item = "Browser",Options.BrVersionPrecision,Options.OsVersionPrecision)
		Res = Items(Item & ".Version.Major")
		If Res="" then Res = "0"
		If VersionPrecision >= 1 then
		   Res =  Res & "." & _
			   IIF(Items(Item & ".Version.Minor")<>"",_
				IIF(Item = "Browser" and Options.BrVersionMinorSimple, _
				left(Items(Item & ".Version.Minor"),1),_
			   Items(Item & ".Version.Minor")),"0")
		   If VersionPrecision = 2 then
		 	Res =  Res & "." & _
			IIF(Items(Item & ".Version.Rest")<>"",Items(Item & ".Version.Rest"),"0")
		   End if
		end if
		If Res="0" or Res="0.0" or Res="0.0.0" then Res=""
		GetVersion = Res
	End Function

	private function Browser_Resolve()
	dim tmp,pos
		if Items("Browser.Type")="1" then
			tmp = IIF(Options.RobotPrecision and Items("Browser.Name") <> "Robot" and Items("Browser.Name") <> "",_
				"Robot: " & Items("Browser.Name"),_
				Items("Browser.Name"))
		else
			tmp = Items("Browser.Name")
			if  tmp = "Internet Explorer" then  Tmp = "Microsoft " & Tmp
			tmp = tmp & " " & GetVersion("Browser")
		end if
		
		if trim(tmp) = trim(options.FillResult) then
			if Items("Browser.Url") <> "" then tmp = domainname(Items("Browser.Url"))
		end if
		
	Browser_Resolve = tmp
	end function

	Private function Browser_Name()
	Dim aFiles,aTypes,BrowserSubtype,BrowserType,Find,Count,Name
		pAutoURL=""
		Find = False
		Name = ""
		Data.Open Options.Path & CST_XML_BROWSER
		Do While Not Data.EOF
		   if Options.Robotknown or Data("Type") > 9 then 'fa i robot
		      pBrowserPattern = Data("Pattern")
		      Reg.Pattern = pBrowserPattern
		      if Reg.Test(pUserAgent) then
		         Name = Data("Name")
		         select case data("Type")
		            case 0,1,2,3,8,9
		               BrowserType = 1
		            case 10,15,20
		               BrowserType = 0
		         end select
		         BrowserSubtype = data("Type")
		         pAutoOS = ""
		         pSupport = Data("Support")
		         pAutoOS = Data("Os") : items("Browser.Engine") = Data("Engine")
		         pAutoURL = Data("Url")
		         Find = True
		         exit Do
		      end if
		   end if
		   Data.MoveNext()
		Loop
		Data.Close
		Reg.Pattern = "mozilla"
		if Reg.Test(pUserAgent) and Name = "" then
		   BrowserType = 0 : BrowserSubtype = 10
		   Reg.Pattern = "(rv:(\d+(\.\d)*))|(;[ ](\d+(\.\d)*))|dreamkey|viking"
		   if Reg.Test(pUserAgent) then
		      Name = "Mozilla"
		      pAutoURL = "http://www.mozilla.org"
		   else
		      Name = "Netscape"
		      pAutoURL = ""
		   end if
		end if
		
		if Items("OS.Type") = 1 then Name = Items("OS.Name")
		
		Items("Browser.Guessed") = False
		if Name = "" then
			BrowserType = 0 : BrowserSubtype = 10
			Items("Browser.Guessed") = True
			Name = Guess_name()
		end if
		if Name = "Robot" then Name = Guess_name()
		
		Browser_Name = Name
		Items("Browser.Type") = BrowserType
		Items("Browser.Type.Description") = DecodeType(BrowserType)
		Items("Browser.Subtype") = BrowserSubtype
		Items("Browser.Subtype.Description") = DecodeSubType(BrowserSubtype)
	End Function

	Private Function Browser_Version()
	dim Res,Major,Minor,Rest,Matches
	if right(Items("OS.Name"),4) = "Palm" or right(Items("OS.Name"),6) = "Mobile" or _
	right(Items("OS.Name"),7) = "Console" then exit function
		select case lcase(Items("Browser.Name"))
		   case ""
		      Res = ""
		   case "mozilla":
		      Reg.Pattern = "()(?:rv:)" & CST_CHAR_IDVERSION
		      if Reg.Test(pUserAgent) then
		         set Matches = reg.Execute(pUserAgent)
		         Major = Matches(0).SubMatches(2)
			 Minor = Matches(0).SubMatches(3)
			 Rest = Matches(0).SubMatches(4)
		         Res = Matches(0).SubMatches(1)
		      else
		         Reg.Pattern = "^(mozilla)" & CST_CHAR_VERSION & "?" & CST_CHAR_IDVERSION
		         if Reg.Test(pUserAgent) then
		            set Matches = reg.Execute(pUserAgent)
		            Major = Matches(0).SubMatches(2)
			    Minor = Matches(0).SubMatches(3)
			    Rest = Matches(0).SubMatches(4)
		            Res = Matches(0).SubMatches(1)
		         end if
		      end if
		   case "netscape":
		   
		      Reg.Pattern = "(mozilla)" & CST_CHAR_VERSION & "?" & CST_CHAR_IDVERSION
		      if Reg.Test(pUserAgent) then
		         set Matches = reg.Execute(pUserAgent)
		         Major = Matches(0).SubMatches(2)
			 Minor = Matches(0).SubMatches(3)
			 Rest = Matches(0).SubMatches(4)
		         Res = Matches(0).SubMatches(1)
		      end if

		     Reg.Pattern = "(mozilla)" & CST_CHAR_VERSION & "?" & CST_CHAR_IDVERSION & "gold"
		      if Reg.Test(pUserAgent) then
		         set Matches = reg.Execute(pUserAgent)
		         Major = Matches(0).SubMatches(2)
			 Minor = Matches(0).SubMatches(3)
			 Rest = Matches(0).SubMatches(4) & "gold"
		         Res = Matches(0).SubMatches(1)
		      end if

		      if pBrowserPattern <> "" then
		         Reg.Pattern = pBrowserPattern & CST_CHAR_VERSION & "?" & CST_CHAR_IDVERSION
		         if Reg.Test(pUserAgent) then
		            set Matches = reg.Execute(pUserAgent)
		            Major = Matches(0).SubMatches(2)
			    Minor = Matches(0).SubMatches(3)
			    Rest = Matches(0).SubMatches(4)
		            Res = Matches(0).SubMatches(1)
		         end if
		      end if
		   case Else
		      if pBrowserPattern <> "" then
		         Reg.Pattern = pBrowserPattern & CST_CHAR_VERSION & "?" &  CST_CHAR_IDVERSION
		         if Reg.Test(pUserAgent) then
		            set Matches = reg.Execute(pUserAgent)
		            Major = Matches(0).SubMatches(2)
			    Minor = Matches(0).SubMatches(3)
			    Rest = Matches(0).SubMatches(4)
		            Res = Matches(0).SubMatches(1)
		         end if
		      end if
		end select
		
		Items("Browser.Version.Major") = Fill( Major, False , True)
		Items("Browser.Version.Minor") = Fill( Minor, False , True)
		Items("Browser.Version.Rest") =  Fill( Rest,  False , True)
		Browser_Version = Res

	end function
	
' ***************************************************************************************************

	Private function Browser_Act_Cookie()
		If Request.ServerVariables("HTTP_COOKIE") = "" Then
			Items("Browser.Act.Cookie") = False
		else
			Items("Browser.Act.Cookie") = True
		End if
	end function

	Private Sub Browser_AC_Insert(Css,Frames,iFrames,Tables,Sound,Vb,Js,Applet,ActiveX)
	
		Items("Browser.Act.Css")	= IIF(Css <> -1,	Css,		"")
		Items("Browser.Act.Frames")	= IIF(Frames <> -1,	ToBool(Frames),	"")
		Items("Browser.Act.iFrames")	= IIF(iFrames <> -1,	ToBool(iFrames),"")
		Items("Browser.Act.Tables")	= IIF(Tables <> -1,	ToBool(Tables),	"")
		Items("Browser.Cap.Sound")	= IIF(Sound <> -1,	ToBool(Sound),	"")
		Items("Browser.Cap.Vb")		= IIF(Vb <> -1,		ToBool(Vb),	"")
		Items("Browser.Cap.Js")		= IIF(Js <> -1,		ToBool(Js),	"")
		Items("Browser.Cap.Applett")	= IIF(Applet <> -1,	ToBool(Applet),	"")
		Items("Browser.Cap.Activex")	= IIF(ActiveX <> -1,	ToBool(ActiveX),"")
		
	end sub

	Private Sub Browser_ActAndCap()
	Dim Elm,Ar,Values
	
		If Items("Browser.Type") = 0 then
		   select case Items("Browser.Name")
			
		      case "Netscape":
			   select case Items("Browser.Version.Major")
			      case 6,7: 	Browser_AC_Insert 2,1,1,1,0,0,1,1,0
			      case 4	: 	Browser_AC_Insert 1,1,0,1,0,0,1,1,0
			      case 3	: 	Browser_AC_Insert 0,1,0,1,0,0,1,1,0
			      case 2	: 	Browser_AC_Insert 0,1,0,1,0,0,1,1,0
			      case else : 	Browser_AC_Insert -1,-1,-1,1,-1,-1,-1,-1,-1
			   end select
			      
			case "Mozilla":	Browser_AC_Insert 2,1,1,1,0,0,1,1,0
			   
			case else:
			   If pSupport<>"" then
			      Ar = Split(pSupport,VbCrLf)
			      For Each Elm in Ar
			         Elm = Trim(Elm)
			         Elm = Mid(Elm,2,InStrRev(Elm,")")-2)
			         Values = Split(Elm,")(")
			         If Values(0) = Items("Browser.Version.Major") or Values(0)="x" then
			         If Values(1) = Items("Browser.Version.Minor") or Values(1)="x" then
			            Execute "Browser_AC_Insert " & Values(2)
			            Exit For
				 end if
			         End if
			      Next
		            end if
		   end select
		end if
		
	end sub
	
' ***************************************************************************************************

	Private function Browser_Url()
	dim Res,Matches
	
		Res= ""
		if pAutoURL = "" then
			Reg.Pattern = "((http\:\/\/)?(www)?([^\@\.\[\]" & CST_CHAR_SPACE & _
			"]{3,}\.){1,2}[^\@\.\[\]\d" & CST_CHAR_SPACE & "]{2,3})(?:[\@\.\[\]" & _
			CST_CHAR_SPACE & "].*)?$"
			if Reg.Test(pUserAgent) then
				set Matches = reg.Execute(pUserAgent)
				if Matches(0).SubMatches(1) = "" then Res = "http://"
				Res = Res & Matches(0).SubMatches(0)
			end if
		else
			Res = pAutoURL
		end if
		
	Browser_Url=lcase(Res)
	end function
	
' ***************************************************************************************************

	Private Function Browser_Security()
	dim Res,Matches
		Reg.Pattern = "(?:\[|[ ]|,|;)([uin])(?:\]|[ ]|,|;|\))"
		if Reg.Test(pUserAgent) then
			Set Matches = reg.Execute(pUserAgent)
			Res = Matches(0).SubMatches(0)
		end if
		Res = uCase (Res)
		Browser_Security = Res
	end function
	
' ***************************************************************************************************

	Private function Browser_language()
		Dim Language,LanguageCode ,Matches,Patterns(2),Elm
		
		Patterns(0)="(?:[ ]|,|;)([a-z]{2,5})(?:,|;|\))"
		Patterns(1)="(?:\[|[-])([a-z]{2,5})(?:\]|[-])"
		Patterns(2)="(?:\[|[ ]|,|;)([a-z]{2,5}-[a-z]{2,5})(?:\]|,|;|\))"
		for Each Elm in Patterns
			Reg.Pattern =  Elm
			if Reg.Test(pUserAgent) then
				set Matches = reg.Execute(pUserAgent)
				LanguageCode = Matches(0).SubMatches(0)
				Language = DecodeLanguage(LanguageCode )
			end if
		next
		if Language = "" then
			LanguageCode	= lcase(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE"))
			Language		= DecodeLanguage(LanguageCode)
		End if
		Browser_Language = Language
		Items("Browser.Language.Code") = Fill(LanguageCode,True,False)
	end function

	Private Function DecodeLanguage(Value)
		Data.Open Options.Path & CST_XML_LANGUAGE
		
		Do While Not Data.EOF
			Reg.Pattern = "^" & Data("Pattern")
			if Reg.Test(Value) then
				DecodeLanguage = Data("Name")
				exit Do
			end if
			Data.MoveNext()
		Loop
		Data.Close
	End Function


' ***************************************************************************************************
	'************************************ Operating System **************************************

	Private function OS_Name()
	Dim Res,ArDesc
		Res = ""
		arDesc = Array("Operating System","Console, Mobile, Palm")
		Items("OS.Type") = 0
		Data.Open Options.Path & CST_XML_OS
		Do While Not Data.EOF
			pOSPattern = Data("Pattern")
			Reg.Pattern = pOSPattern
			if Reg.Test(pUserAgent) then
				Res = Data("Name")
				Items("OS.Type") = Data("Type")
				Items("OS.Type.Description") = arDesc(Data("Type"))
				exit Do
			end if
			Data.MoveNext()
		Loop
		Data.Close
	OS_name = Res
	end function

	Private Function OS_Subtype()
	dim Res,Matches,BrowserType,BrowserSubtype
		Select Case lcase(Items("OS.Name"))
		   case "windows":
		      Reg.Pattern = "win(?:dows)?(?: )?(9x|95|98)"
		      if Reg.Test(pUserAgent) then Res = "9x"
		      Reg.Pattern = "win(?:dows)?(?: )?(?:nt)?(?: )?(?:nt|5\.\d|xp|200\d)"
		      if Reg.Test(pUserAgent) then Res = "NT"
		   case "macintosh": 'sistema vecchiotto MAC
		      Res = ""
		   case "linux": 'distribuzione
		      Reg.Pattern = "(debian|mdk|slack|redhat|gentoo|suse)"
		      if Reg.Test(pUserAgent) then
		         set Matches = Reg.Execute(pUserAgent)
		         Res = lcase(Matches(0).SubMatches(0))
		         Res = replace(Res, "debian", "Debian")
		         Res = replace(Res, "mdk", "Mandrake")
		         Res = replace(Res, "slack", "Slackwave")
		         Res = replace(Res, "redhat", "RedHat")
		         Res = replace(Res, "gentoo", "Gentoo")
		         Res = replace(Res, "suse", "Suse")
		      end if
		   case "mac": 'sistema moderno MAC OS
		      Res = "OS"
		   case lcase("SonyEricsson Mobile"), lcase("Nokia Mobile"):
		      BrowserType = 0
		      BrowserSubtype = 20
		      Items("Browser.Type") = BrowserType
		      Items("Browser.Type.Description") = DecodeType(BrowserType)
		      Items("Browser.Subtype") = BrowserSubtype
		      Items("Browser.Subtype.Description") = DecodeSubType(BrowserSubtype)
		      Reg.Pattern = "(?:(nokia|sony)(?: |-)?(?:ericsson)?)(?:[ ]|\/)?(\w*)"
		      if Reg.Test(pUserAgent) then
		         set Matches = reg.Execute(pUserAgent)
		         if Matches(0).SubMatches(0) = "nokia" then
		            Res = lcase(Matches(0).SubMatches(1))
		         else
		            Res = ucase(Matches(0).SubMatches(1))
		         end if
		      end if
		End Select
	OS_Subtype = Res
	end function

	Private Function OS_version()
	Dim Res,Major,Minor,Rest,Matches
	if Items("OS.Type") = 1 then exit function
		select case lcase(Items("OS.Name"))
		   case "":
		      Res = ""
		   case "macintosh":
		      Res = ""
		   case "mac":
		      Reg.Pattern = "(?:mac(?:intosh)?(?: )?(?:os)?(?: )?)(x|\d+)"
		      if Reg.Test(pUserAgent) then
		         set Matches = reg.Execute(pUserAgent)
		         Major = ucase(Matches(0).SubMatches(0))
		         if Major = "X" then Major = "10"
		         Res = Major
		      end if
		   case "linux":
		      Reg.Pattern = "linux" & CST_CHAR_VERSION & "?" &  CST_CHAR_IDVERSION
		      if Reg.Test(pUserAgent) then
		         set Matches = reg.Execute(pUserAgent)
		         Major = Matches(0).SubMatches(1)
			 Minor = Matches(0).SubMatches(2)
			 Rest = Matches(0).SubMatches(3)
		         Res = Matches(0).SubMatches(0)
		      end if
		   case "windows":
		      Reg.Pattern = "(?:win(?:dows)?(?: )?(?:9x)?(?:nt)?(?: )?)" & CST_CHAR_IDVERSION
		      if Reg.Test(pUserAgent) then
		         set Matches = reg.Execute(pUserAgent)
		         Major = Matches(0).SubMatches(1)
			 Minor = Matches(0).SubMatches(2)
			 Rest = Matches(0).SubMatches(3)
		         Res = Matches(0).SubMatches(0)
		      end if
		   case "sun os":
		   	Res = ""
		   case Else
		      Reg.Pattern = lcase(pOSPattern) & CST_CHAR_VERSION & "?" &  CST_CHAR_IDVERSION
		      if Reg.Test(pUserAgent) then
		         set Matches = reg.Execute(pUserAgent)
		         Major = Matches(0).SubMatches(2)
			 Minor = Matches(0).SubMatches(3)
			 Rest = Matches(0).SubMatches(4)
		         Res = Matches(0).SubMatches(1)
		      end if
		end select

		if Items("OS.Name") = "Windows" then
			select case lcase(Major)
			   case "2000"
			      Major = "5" : Minor = "0" : Rest = ""
			   case "xp"
			      Major = "5" : Minor = "1" : Rest = ""
			   case "2003"
			      Major = "5" : Minor = "2" : Rest = ""
			end select
		end if

	OS_version = Res
	Items("OS.Version.Major") = Fill(Major,False,True)
	Items("OS.Version.Minor") = Fill(Minor,False,True)
	Items("OS.Version.Rest") = Fill(Rest,False,True)
	end function

' ***************************************************************************************************

	Private function OS_arch()
	Dim Res,Matches
	if Items("OS.Type") = 1 then OS_arch = Items("OS.Name") : exit function
		select case lcase(Items("OS.Name"))
		   case "linux","unix","bsd","irix","hp-ux","sco","aix","reliant","dec","sinix",_
		        "vms","unixware","mpras","sun os":
		      Reg.Pattern = "(i\d86|x86_64|(?:strong)?(?:[ ]|-)?arm|ia64|m68k|ppc(64)?|mips(64)?|" & _
		      		    "cris|parisc|alpha|s390(x)?|sh|sparc(64)?)"
		      if Reg.Test(pUserAgent) then
		         set Matches = reg.Execute(pUserAgent)
		         Res = Matches(0).SubMatches(0)
		         if right(Res,2)="86" and Res <> "i386" then Res = "i386"
		      end if
		   case "mac","macintosh":
		      Res = "Mac"
		      Reg.Pattern = "ppc|power(?:[ ]|_|-)?pc"
		      if Reg.Test(pUserAgent) then
		         Res = "PowerPC"
		      end if
		      Reg.Pattern = "68k|68000"
		      if Reg.Test(pUserAgent) then
		         Res = "68k"
		      end if
		   case "windows":
		      Res = "i386"
		   case Else
		      Reg.Pattern = "(commodore|atari|amiga)"
		      if Reg.Test(pUserAgent) then
		         set Matches = reg.Execute(pUserAgent)
		         Res = Items("OS.Name")
		      end if
		end select
	OS_arch = Res
	end function
	
' ***************************************************************************************************

	private function OS_resolve()
	dim Res
		if Items("OS.Name")="Windows" then
			Data.Open Options.Path & CST_XML_OS_WINDOWS
			Do While Not Data.EOF
				Reg.Pattern = Data("Pattern")
				if Reg.Test(pUserAgent) then
					OS_resolve= Data("Name")
					Items("OS.Version.Major") = Fill(Data("Major"),False, True)
					Items("OS.Version.Minor") = Fill(Data("Minor"),False, True)
					Items("OS.Version.Rest") = Fill(Data("Rest"),  False, True)
					Items("OS.SubType") = Data("SubType")
					exit Do
				end if
				Data.MoveNext()
			Loop
			Data.Close
		elseif Items("OS.Name")="Linux" then
			OS_resolve=trim(Items("OS.Name"))
		elseif Items("OS.Name") = "SonyEricsson Mobile" or _
		       Items("OS.Name") = "Nokia Mobile" then
			OS_resolve=trim(Items("OS.Name"))
		else
			Res = Items("OS.Name")
			If Items("OS.SubType") <> "" and Items("OS.SubType") <> "undefined" then _
				Res = Res & " " & Items("OS.SubType")
			Res = Res & " " & GetVersion("OS")
			Os_Resolve = Res
		end if
	end function

	'************************************ Gecko, Mozilla, Framework *****************************

	Private function Gecko_Version()
	Dim Res,Matches
		Reg.Pattern = "gecko" & CST_CHAR_VERSION & "?(\d{1,8})"
		if Reg.Test(pUserAgent) then
			set Matches = reg.Execute(pUserAgent)
			Res = Matches(0).SubMatches(0)
			Set Matches = nothing
		end if
		Gecko_Version = Res
	end function

	Private function Mozilla_Version()
	Dim Res,Matches
		Reg.Pattern = "^(mozilla)" & CST_CHAR_VERSION & "?" & CST_CHAR_IDVERSION
		if Reg.Test(pUserAgent) then
			set Matches = reg.Execute(pUserAgent)
			if Matches(0).SubMatches(2) & "." & Matches(0).SubMatches(3) = "." then
			  Res = 0
			elseif Matches(0).SubMatches(2) & "." & Matches(0).SubMatches(3) = Matches(0).SubMatches(2) & "." then
			  Res = Matches(0).SubMatches(2)
			else
			  Res = Matches(0).SubMatches(2) & "." & Matches(0).SubMatches(3)
			end if
			Set Matches = nothing
		end if
		Mozilla_version = Res
	end function

	Private Function Frameworkdotnet()
	dim Res,Matches,Elm,last,lastreal
	    Res = ""
	    Reg.Pattern = "\.net clr (\d+(\.\d+)*)"
	    if Reg.Test(pUserAgent) then
		set Matches = reg.Execute(pUserAgent)
		last = 0
		lastreal = 0
		for Each Elm In Matches
			if last < replace(Elm.SubMatches(0),".","") then
			last = replace(Elm.SubMatches(0),".","")
			lastreal = Elm.SubMatches(0)
			end if
		next
		Res = lastreal
	    end if
	Frameworkdotnet = Res
	end function

	'************************************ Guess *************************************************

	Private function Guess_Name()
	dim Res,Matches,Patterns(4),Elm
	
		if Items("Browser.Name") = "" then
	
		    Patterns(0)="^(?:.*[" & CST_CHAR_SPACE & ":.])?([^" & CST_CHAR_SPACE & _
		    ":.@]*(spider|crawl|(?:[^r]|^)(?:[^o]|^)bot)[^" & CST_CHAR_SPACE & ":.@]*).*"
			
		    Patterns(1)="^([^" & CST_CHAR_SPACE & ":.]+)([" & CST_CHAR_SPACE & "])"
			
		    Patterns(2)="(?:[" & CST_CHAR_SPACE & "])([^" & CST_CHAR_SPACE & ":.]+)" & _
		    "(?:[" & CST_CHAR_SPACE & "])"
			
		    Patterns(3)="$(?:[" & CST_CHAR_SPACE & "])([^" & CST_CHAR_SPACE & ":.]+)"
			
		    Patterns(4)="(?:[" & CST_CHAR_SPACE & "])?([^" & CST_CHAR_SPACE & "]+)" & _
		    "(?:[" & CST_CHAR_SPACE & "])?"
			
		    for Each Elm in Patterns
		        Reg.Pattern = Elm
		        if reg.test(pUserAgent) and Res = "" then
		            set Matches = reg.Execute(pUserAgent)
		            Res = Matches(0).SubMatches(0)
		            exit for
		        end if
		    next
		    
		    if Res = "" or lcase(Res) = "mozilla" then Res = pUserAgent
		    if Res = "http:" then Res = items("Browser.Url")
			
		    Guess_Name = Res
			
		else
		
		    Guess_Name = Items("Browser.Name")
			
		end if
		
	end function

	private sub Guess_Version()
	dim Matches,Patterns(1),Elm
	   if Items("Browser.Version") = "" or Items("Browser.Version") = "0" then
	      Patterns(0) = Items("Browser.Name") & CST_CHAR_VERSION & "(?:v)?" & CST_CHAR_IDVERSION
	      Patterns(1) = "" & CST_CHAR_VERSION & "(?:v)?" & CST_CHAR_IDVERSION
	      for Each Elm In Patterns
	         Reg.Pattern = Elm
	         if Reg.Test(pUserAgent) then
	            set Matches = reg.Execute(pUserAgent)
	            Items("Browser.Version.Major") = Fill(Matches(0).SubMatches(1),False,True)
	            Items("Browser.Version.Minor") = Fill(Matches(0).SubMatches(2),False,True)
	            Items("Browser.Version.Rest") = Fill(Matches(0).SubMatches(3),False,True)
	            Items("Browser.Version") = Matches(0).SubMatches(0)
	            exit for
	         end if
	      next
	   end if
	end sub

	Private Sub Class_Terminate()
		Items.RemoveAll
		Set Data = Nothing
		Set Options = nothing
		Set Items = Nothing
		Set Reg = Nothing
	end sub

End Class


Function CreateWBstat(pamPath,pamCache,pamFill,pamBrlev,pamOSlev,pamBrMinSimple,pamUA,_
		      pamDebug,pamGuess,pamRobot,pamRobotPrec,pamworkBr,pamworkBrver,_
		      pamworkBrdet,pamworkBrlang,pamworkBrACTCAP,pamworkOS,pamworkOSver,_
		      pamworkOSdet,pamworkSpecial)
Dim obj
	Set obj = new wbstatclass
	
	If not(pamPath = "") then	  obj.Options.Path = 		     pamPath
	If not(pamFill = "") then	  obj.Options.Fillresult =           pamFill
	If not(pamCache = "") then	  obj.Options.Cache =		     pamCache
	If not(pamDebug = "") then	  obj.Options.Debugging = 	     pamDebug
	If not(pamBrLev = "") then	  obj.Options.BrVersionPrecision =   pamBrLev
	If not(pamOSLev = "") then	  obj.Options.OSVersionPrecision =   pamOSLev
	If not(pamBrMinSimple = "") then  obj.Options.BrVersionMinorSimple = pamBrMinSimple
	If not(pamBrMinSimple = "") then  obj.Options.BrVersionMinorSimple = pamBrMinSimple
	If not(pamUA = "") then 	  obj.Options.IncludeUserAgent =     pamUA
	If not(pamGuess = "") then	  obj.Options.Guessing =	     pamGuess
	If not(pamRobot = "") then	  obj.Options.Robotknown =	     pamRobot
	If not(pamRobotPrec = "") then	  obj.Options.Robotprecision =	     pamRobotPrec
	
	If not(pamworkBr = "") then 	  obj.Options.workBrowser = 	     pamworkBr
	If not(pamworkBrver = "") then 	  obj.Options.workBrowserversion =   pamworkBrver
	If not(pamworkBrdet = "") then	  obj.Options.workBrowserdetails =   pamworkBrdet
	If not(pamworkBrlang = "") then	  obj.Options.workBrowserlanguage =  pamworkBrlang
	If not(pamworkBrACTCAP = "") then obj.Options.workBrowserACTCAP =    pamworkBrACTCAP
	If not(pamworkOS = "") then	  obj.Options.workOS = 	     	     pamworkOS
	If not(pamworkOSver = "") then	  obj.Options.workOSversion = 	     pamworkOSver
	If not(pamworkOSdet = "") then	  obj.Options.workOSdetails = 	     pamworkOSdet
	If not(pamworkSpecial = "") then  obj.Options.workSpecial = 	     pamworkSpecial

	obj.eval()
	Set CreateWBstat = obj
	Set Obj = nothing
End Function

Function CreateWBstatSimple(pamPath,pamCache,pamFill,pamDebug)
Dim obj
	Set obj = new wbstatclass
	If not(pamPath = "") then 		obj.Options.Path = 			pamPath
	If not(pamFill = "") then 		obj.Options.Fillresult = 		pamFill
	If not(pamCache = "") then 		obj.Options.Cache = 			pamCache
	If not(pamDebug = "") then 		obj.Options.Debugging = 		pamDebug
	obj.eval()
	Set CreateWBstatSimple = obj
	Set Obj = nothing
End Function

'####################################################################################################
'#  WBStat statistic class - Copyleft 2002-2004 Simone Cingano										#
'####################################################################################################
%>