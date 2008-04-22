<%

' 
' = ASP Stats Generator - Powerful and reliable ASP website counter
' 
' Copyright (c) 2003-2008 Simone Carletti <weppos@weppos.net>
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
' 
' 
' @category        ASP Stats Generator
' @package         ASP Stats Generator
' @author          Simone Carletti <weppos@weppos.net>
' @copyright       2003-2008 Simone Carletti
' @license         http://www.opensource.org/licenses/mit-license.php
' @version         SVN: $Id$
' 


'-----------------------------------------------------------------------------------------
' FUNZIONI DI ELABORAZIONE	
'-----------------------------------------------------------------------------------------
Dim strtmp
Dim inttmp
Dim dtmtmp
Dim looptmp


'-----------------------------------------------------------------------------------------
' Decodifica URLEncode()	
'-----------------------------------------------------------------------------------------
' Funzione:	Decodifica in caratteri classici la codifica URLEncode()
' Data: 	16.11.2003 | 07.04.2004
' Commenti:	Tratto dal sito di Mems (www.oscarjsweb.com) - forum HTML.it	
'			07.04.2004 Aggiunto filtro di adattamento potenziato
'-----------------------------------------------------------------------------------------
function DecodeURL(url, filterplus)
	
	If filterplus = True Then
		url = Replace(url, "+", " ")
		url = Replace(url, "%20", " ")
	End If

	For looptmp = 1 to 255
		url = Replace(url, Server.URLEncode(chr(looptmp)), chr(looptmp))
	Next
	
	DecodeURL = url

end function


'-----------------------------------------------------------------------------------------
' Pulisci Input	
'-----------------------------------------------------------------------------------------
' Funzione:	
' Data: 	25.11.2003 | 11.05.2004
' Commenti:	
'-----------------------------------------------------------------------------------------
function FilterSQLInput(ByVal input)

	'Remove malicious input for SQL execution from data
	input = Replace(input, "&", "&amp;", 1, -1, 1)
	input = Replace(input, "<", "&lt;")
	input = Replace(input, ">", "&gt;")
	input = Replace(input, "[", "&#091;")
	input = Replace(input, "]", "&#093;")
	input = Replace(input, """", "", 1, -1, 1)
	input = Replace(input, "=", "&#061;", 1, -1, 1)
	input = Replace(input, "'", "''", 1, -1, 1)
	input = Replace(input, "select", "sel&#101;ct", 1, -1, 1)
	input = Replace(input, "join", "jo&#105;n", 1, -1, 1)
	input = Replace(input, "union", "un&#105;on", 1, -1, 1)
	input = Replace(input, "where", "wh&#101;re", 1, -1, 1)
	input = Replace(input, "insert", "ins&#101;rt", 1, -1, 1)
	input = Replace(input, "delete", "del&#101;te", 1, -1, 1)
	input = Replace(input, "update", "up&#100;ate", 1, -1, 1)
	input = Replace(input, "like", "lik&#101;", 1, -1, 1)
	input = Replace(input, "drop", "dro&#112;", 1, -1, 1)
	input = Replace(input, "create", "cr&#101;ate", 1, -1, 1)
	input = Replace(input, "modify", "mod&#105;fy", 1, -1, 1)
	input = Replace(input, "rename", "ren&#097;me", 1, -1, 1)
	input = Replace(input, "alter", "alt&#101;r", 1, -1, 1)
	input = Replace(input, "cast", "ca&#115;t", 1, -1, 1)

	FilterSQLInput = input
	
end function


'-----------------------------------------------------------------------------------------
' Purifica Input	
'-----------------------------------------------------------------------------------------
' Funzione: 
' Data: 	25.11.2003 | 25.11.2003
' Commenti:	
'-----------------------------------------------------------------------------------------
function CleanInput(ByVal input)

	'Elimina i valori
	input = Replace(input, "&", "", 1, -1, 1)
	input = Replace(input, "<", "", 1, -1, 1)
	input = Replace(input, ">", "", 1, -1, 1)
	input = Replace(input, "'", "", 1, -1, 1)
	input = Replace(input, """", "", 1, -1, 1)

	CleanInput = input
	
end function


'-----------------------------------------------------------------------------------------
' Permetti Accesso	
'-----------------------------------------------------------------------------------------
' Funzione: 
' Data: 	30.11.2003 | 30.11.2003
' Commenti:	
'-----------------------------------------------------------------------------------------
function AllowEntry(ByVal nessuno, ByVal limitato, ByVal completo, ByVal protezione)
	
	Dim aryAsgPermetti(2)
	
	aryAsgPermetti(0) = CBool(nessuno)
	aryAsgPermetti(1) = CBool(limitato)
	aryAsgPermetti(2) = CBool(completo)
	
	If aryAsgPermetti(protezione) = False Then
	
		If Session("AsgLogin") <> "Logged" Then
			
			'Pulisci
			Set objAsgRs = Nothing
			objAsgConn.Close
			Set objAsgConn = Nothing
			
			'Indirizza
			Response.Redirect("login.asp?backto=" & Server.URLEncode(Request.ServerVariables("URL")))
		
		End If
		
	End If

end function


'-----------------------------------------------------------------------------------------
' 
'-----------------------------------------------------------------------------------------
' Funzione: 
' Data: 	
' Commenti:	
'-----------------------------------------------------------------------------------------
function GetContinent(ByVal country)
	
	Select Case country
		Case "AFGHANISTAN"
			strtmp = strAsgTxtAsia
		Case "ALBANIA"
			strtmp = strAsgTxtEurope
		Case "ALGERIA"
			strtmp = strAsgTxtAfrica
		Case "AMERICAN SAMOA"
			strtmp = strAsgTxtAmerica
		Case "ANDORRA"
			strtmp = strAsgTxtEurope
		Case "ANGOLA"
			strtmp = strAsgTxtAfrica
		Case "ANTIGUA AND BARBUDA"
			strtmp = strAsgTxtAmerica
		Case "ARGENTINA"
			strtmp = strAsgTxtAmerica
		Case "ARMENIA"
			strtmp = strAsgTxtAsia
		Case "AUSTRALIA"
			strtmp = strAsgTxtOceania
		Case "AUSTRIA"
			strtmp = strAsgTxtEurope
		Case "AZERBAIJAN"
			strtmp = strAsgTxtAsia
		Case "BAHAMAS"
			strtmp = strAsgTxtAmerica
		Case "BAHRAIN"
			strtmp = strAsgTxtAsia
		Case "BANGLADESH"
			strtmp = strAsgTxtAsia
		Case "BARBADOS"
			strtmp = strAsgTxtAmerica
		Case "BELARUS"
'			strtmp = ""
		Case "BELGIUM"
			strtmp = strAsgTxtEurope
		Case "BELIZE"
			strtmp = strAsgTxtAmerica
		Case "BENIN"
			strtmp = strAsgTxtAfrica
		Case "BERMUDA"
			strtmp = strAsgTxtAmerica
		Case "BHUTAN"
			strtmp = strAsgTxtAsia
		Case "BOLIVIA"
			strtmp = strAsgTxtAmerica
		Case "BOSNIA AND HERZEGOVINA"
			strtmp = strAsgTxtEurope
		Case "BOTSWANA"
			strtmp = strAsgTxtAfrica
		Case "BRAZIL"
			strtmp = strAsgTxtAmerica
		Case "BRITISH INDIAN OCEAN TERRITORY"
'			strtmp = ""
		Case "BRUNEI DARUSSALAM"
			strtmp = strAsgTxtAsia
		Case "BULGARIA"
			strtmp = strAsgTxtEurope
		Case "BURKINA FASO"
			strtmp = strAsgTxtAfrica
		Case "BURUNDI"
			strtmp = strAsgTxtAfrica
		Case "CAMBODIA"
			strtmp = strAsgTxtAsia
		Case "CAMEROON"
			strtmp = strAsgTxtAfrica
		Case "CANADA"
			strtmp = strAsgTxtAmerica
		Case "CAPE VERDE"
			strtmp = strAsgTxtAfrica
		Case "CAYMAN ISLANDS"
			strtmp = strAsgTxtAmerica
		Case "CENTRAL AFRICAN REPUBLIC"
			strtmp = strAsgTxtAfrica
		Case "CHAD"
			strtmp = strAsgTxtAfrica
		Case "CHILE"
			strtmp = strAsgTxtAmerica
		Case "CHINA"
			strtmp = strAsgTxtAsia
		Case "COLOMBIA"
			strtmp = strAsgTxtAmerica
		Case "COMOROS"
			strtmp = strAsgTxtAfrica
		Case "CONGO"
			strtmp = strAsgTxtAsia
		Case "COOK ISLANDS"
			strtmp = strAsgTxtOceania
		Case "COSTA RICA"
			strtmp = strAsgTxtAmerica
		Case "COTE D'IVOIRE"
			strtmp = strAsgTxtAfrica
		Case "CROATIA"
			strtmp = strAsgTxtEurope
		Case "CUBA"
			strtmp = strAsgTxtAmerica
		Case "CYPRUS"
			strtmp = strAsgTxtEurope
		Case "CZECH REPUBLIC"
			strtmp = strAsgTxtEurope
		Case "DENMARK"
			strtmp = strAsgTxtEurope
		Case "DJIBOUTI"
			strtmp = strAsgTxtAfrica
		Case "DOMINICAN REPUBLIC"
			strtmp = strAsgTxtAmerica
		Case "EAST TIMOR"
'			strtmp = ""
		Case "ECUADOR"
			strtmp = strAsgTxtAmerica
		Case "EGYPT"
			strtmp = strAsgTxtAfrica
		Case "EL SALVADOR"
			strtmp = strAsgTxtAmerica
		Case "EQUATORIAL GUINEA"
			strtmp = strAsgTxtAfrica
		Case "ERITREA"
			strtmp = strAsgTxtAfrica
		Case "ESTONIA"
			strtmp = strAsgTxtEurope
		Case "ETHIOPIA"
			strtmp = strAsgTxtAfrica
		Case "FALKLAND ISLANDS (MALVINAS)"
			strtmp = strAsgTxtAmerica
		Case "FAROE ISLANDS"
'			strtmp = ""
		Case "FIJI"
			strtmp = strAsgTxtOceania
		Case "FINLAND"
			strtmp = strAsgTxtEurope
		Case "FRANCE"
			strtmp = strAsgTxtEurope
		Case "FRENCH POLYNESIA"
'			strtmp = ""
		Case "GABON"
			strtmp = strAsgTxtAfrica
		Case "GAMBIA"
			strtmp = strAsgTxtAfrica
		Case "GEORGIA"
'			strtmp = ""
		Case "GERMANY"
			strtmp = strAsgTxtEurope
		Case "GHANA"
			strtmp = strAsgTxtAfrica
		Case "GIBRALTAR"
'			strtmp = ""
		Case "GREECE"
			strtmp = strAsgTxtEurope
		Case "GREENLAND"
'			strtmp = ""
		Case "GRENADA"
			strtmp = strAsgTxtAmerica
		Case "GUADELOUPE"
			strtmp = strAsgTxtAmerica
		Case "GUAM"
			strtmp = strAsgTxtAmerica
		Case "GUATEMALA"
			strtmp = strAsgTxtAmerica
		Case "GUINEA"
			strtmp = strAsgTxtAfrica
		Case "GUINEA-BISSAU"
			strtmp = strAsgTxtAfrica
		Case "HAITI"
			strtmp = strAsgTxtAmerica
		Case "HOLY SEE (VATICAN CITY STATE)"
			strtmp = strAsgTxtEurope
		Case "HONDURAS"
			strtmp = strAsgTxtAmerica
		Case "HONG KONG"
			strtmp = strAsgTxtAsia
		Case "HUNGARY"
			strtmp = strAsgTxtEurope
		Case "ICELAND"
			strtmp = strAsgTxtEurope
		Case "INDIA"
			strtmp = strAsgTxtAsia
		Case "INDONESIA"
			strtmp = strAsgTxtAsia
		Case "IRAQ"
			strtmp = strAsgTxtAsia
		Case "IRELAND"
			strtmp = strAsgTxtEurope
		Case "ISLAMIC REPUBLIC OF IRAN"
			strtmp = strAsgTxtAsia
		Case "ISRAEL"
			strtmp = strAsgTxtAsia
		Case "ITALY"
			strtmp = strAsgTxtEurope
		Case "JAMAICA"
			strtmp = strAsgTxtAmerica
		Case "JAPAN"
			strtmp = strAsgTxtAsia
		Case "JORDAN"
			strtmp = strAsgTxtAsia
		Case "KAZAKHSTAN"
			strtmp = strAsgTxtAsia
		Case "KENYA"
			strtmp = strAsgTxtAfrica
		Case "KIRIBATI"
'			strtmp = ""
		Case "KUWAIT"
			strtmp = strAsgTxtAsia
		Case "KYRGYZSTAN"
'			strtmp = ""
		Case "LAO PEOPLE'S DEMOCRATIC REPUBL"
			strtmp = strAsgTxtAsia
		Case "LATVIA"
'			strtmp = ""
		Case "LEBANON"
			strtmp = strAsgTxtAsia
		Case "LESOTHO"
			strtmp = strAsgTxtAfrica
		Case "LIBERIA"
			strtmp = strAsgTxtAfrica
		Case "LIBYAN ARAB JAMAHIRIYA"
			strtmp = strAsgTxtAfrica
		Case "LIECHTENSTEIN"
			strtmp = strAsgTxtEurope
		Case "LITHUANIA"
			strtmp = strAsgTxtEurope
		Case "LUXEMBOURG"
			strtmp = strAsgTxtEurope
		Case "MACAO"
			strtmp = strAsgTxtAsia
		Case "MADAGASCAR"
			strtmp = strAsgTxtAfrica
		Case "MALAWI"
			strtmp = strAsgTxtAfrica
		Case "MALAYSIA"
			strtmp = strAsgTxtAsia
		Case "MALDIVES"
			strtmp = strAsgTxtAsia
		Case "MALI"
			strtmp = strAsgTxtAfrica
		Case "MALTA"
			strtmp = strAsgTxtEurope
		Case "MARTINIQUE"
			strtmp = strAsgTxtAmerica
		Case "MAURITANIA"
			strtmp = strAsgTxtAfrica
		Case "MAURITIUS"
'			strtmp = ""
		Case "MEXICO"
			strtmp = strAsgTxtAmerica
		Case "MONACO"
			strtmp = strAsgTxtEurope
		Case "MONGOLIA"
			strtmp = strAsgTxtAsia
		Case "MOROCCO"
			strtmp = strAsgTxtAfrica
		Case "MOZAMBIQUE"
			strtmp = strAsgTxtAfrica
		Case "MYANMAR"
'			strtmp = ""
		Case "NAMIBIA"
			strtmp = strAsgTxtAfrica
		Case "NAURU"
			strtmp = strAsgTxtAmerica
		Case "NEPAL"
			strtmp = strAsgTxtAsia
		Case "NETHERLANDS"
			strtmp = strAsgTxtEurope
		Case "NETHERLANDS ANTILLES"
			strtmp = strAsgTxtAmerica
		Case "NEW CALEDONIA"
			strtmp = strAsgTxtOceania
		Case "NEW ZEALAND"
			strtmp = strAsgTxtOceania
		Case "NICARAGUA"
			strtmp = strAsgTxtAmerica
		Case "NIGER"
			strtmp = strAsgTxtAfrica
		Case "NIGERIA"
			strtmp = strAsgTxtAfrica
		Case "NORTHERN MARIANA ISLANDS"
'			strtmp = ""
		Case "NORWAY"
			strtmp = strAsgTxtEurope
		Case "OMAN"
			strtmp = strAsgTxtAsia
		Case "PAKISTAN"
			strtmp = strAsgTxtAsia
		Case "PALAU"
			strtmp = strAsgTxtAsia
		Case "PALESTINIAN TERRITORY, OCCUPIE"
			strtmp = strAsgTxtAsia
		Case "PANAMA"
			strtmp = strAsgTxtAmerica
		Case "PAPUA NEW GUINEA"
			strtmp = strAsgTxtAsia
		Case "PARAGUAY"
			strtmp = strAsgTxtAmerica
		Case "PERU"
			strtmp = strAsgTxtAmerica
		Case "PHILIPPINES"
			strtmp = strAsgTxtAsia
		Case "POLAND"
			strtmp = strAsgTxtEurope
		Case "PORTUGAL"
			strtmp = strAsgTxtEurope
		Case "PUERTO RICO"
			strtmp = strAsgTxtAmerica
		Case "QATAR"
			strtmp = strAsgTxtAsia
		Case "REPUBLIC OF KOREA"
			strtmp = strAsgTxtAsia
		Case "REPUBLIC OF MOLDOVA"
'			strtmp = ""
		Case "REUNION"
			strtmp = strAsgTxtAfrica
		Case "ROMANIA"
			strtmp = strAsgTxtEurope
		Case "RUSSIAN FEDERATION"
			strtmp = strAsgTxtAsia
		Case "RWANDA"
			strtmp = strAsgTxtAfrica
		Case "SAMOA"
			strtmp = strAsgTxtOceania
		Case "SAN MARINO"
			strtmp = strAsgTxtEurope
		Case "SAO TOME AND PRINCIPE"
			strtmp = strAsgTxtAfrica
		Case "SAUDI ARABIA"
			strtmp = strAsgTxtAsia
		Case "SENEGAL"
			strtmp = strAsgTxtAfrica
		Case "SERBIA AND MONTENEGRO"
			strtmp = strAsgTxtEurope
		Case "SEYCHELLES"
			strtmp = strAsgTxtAfrica
		Case "SIERRA LEONE"
			strtmp = strAsgTxtAfrica
		Case "SINGAPORE"
			strtmp = strAsgTxtAsia
		Case "SLOVAKIA"
			strtmp = strAsgTxtEurope
		Case "SLOVENIA"
			strtmp = strAsgTxtEurope
		Case "SOLOMON ISLANDS"
			strtmp = strAsgTxtOceania
		Case "SOMALIA"
			strtmp = strAsgTxtAfrica
		Case "SOUTH AFRICA"
			strtmp = strAsgTxtAfrica
		Case "SPAIN"
			strtmp = strAsgTxtEurope
		Case "SRI LANKA"
			strtmp = strAsgTxtAsia
		Case "SUDAN"
			strtmp = strAsgTxtAfrica
		Case "SURINAME"
			strtmp = strAsgTxtAmerica
		Case "SWAZILAND"
			strtmp = strAsgTxtAfrica
		Case "SWEDEN"
			strtmp = strAsgTxtEurope
		Case "SWITZERLAND"
			strtmp = strAsgTxtEurope
		Case "SYRIAN ARAB REPUBLIC"
			strtmp = strAsgTxtAsia
		Case "TAIWAN"
			strtmp = strAsgTxtAsia
		Case "TAJIKISTAN"
			strtmp = strAsgTxtAsia
		Case "THAILAND"
			strtmp = strAsgTxtAsia
		Case "THE DEMOCRATIC REPUBLIC OF THE"
'			strtmp = ""
		Case "THE FORMER YUGOSLAV REPUBLIC O"
'			strtmp = ""
		Case "TOGO"
			strtmp = strAsgTxtAfrica
		Case "TOKELAU"
			strtmp = strAsgTxtOceania
		Case "TONGA"
			strtmp = strAsgTxtOceania
		Case "TRINIDAD AND TOBAGO"
			strtmp = strAsgTxtAmerica
		Case "TUNISIA"
			strtmp = strAsgTxtAfrica
		Case "TURKEY"
			strtmp = strAsgTxtAsia
		Case "TURKMENISTAN"
'			strtmp = ""
		Case "TUVALU"
			strtmp = strAsgTxtOceania
		Case "UGANDA"
			strtmp = strAsgTxtAfrica
		Case "UKRAINE"
			strtmp = strAsgTxtEurope
		Case "UNITED ARAB EMIRATES"
			strtmp = strAsgTxtAsia
		Case "UNITED KINGDOM"
			strtmp = strAsgTxtAmerica
		Case "UNITED REPUBLIC OF TANZANIA"
			strtmp = strAsgTxtAfrica
		Case "UNITED STATES"
			strtmp = strAsgTxtAmerica
		Case "URUGUAY"
			strtmp = strAsgTxtAmerica
		Case "UZBEKISTAN"
'			strtmp = ""
		Case "VANUATU"
			strtmp = strAsgTxtOceania
		Case "VENEZUELA"
			strtmp = strAsgTxtAmerica
		Case "VIET NAM"
			strtmp = strAsgTxtAsia
		Case "VIRGIN ISLANDS, BRITISH"
'			strtmp = ""
		Case "WESTERN SAHARA"
			strtmp = strAsgTxtAsia
		Case "YEMEN"
			strtmp = strAsgTxtAsia
		Case "ZAMBIA"
			strtmp = strAsgTxtAfrica
		Case "ZIMBABWE"
			strtmp = strAsgTxtAfrica
	End Select
	
	GetContinent = strtmp
	
end function

%>