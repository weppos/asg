<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'
			


'-----------------------------------------------------------------------------------------
' FUNZIONI DI ELABORAZIONE	
'-----------------------------------------------------------------------------------------


'-----------------------------------------------------------------------------------------
' Decode values encoded withe the Server.URLEncode() function
'-----------------------------------------------------------------------------------------
Function URLDecode(url, powered)
	
	Dim i
	if powered then
		url = Replace(url, "+", " ")
		url = Replace(url, "%20", " ")
	end If

	for i = 1 to 255
		url = Replace(url, Server.URLEncode(chr(i)), chr(i))
	next
	' Return function
	URLDecode = url

End Function


'-----------------------------------------------------------------------------------------
' Purifica Input	
'-----------------------------------------------------------------------------------------
' Funzione: 
' Date: 	25.11.2003 | 25.11.2003
' Comment:	
'-----------------------------------------------------------------------------------------
Function CleanInput(input)

	' Remove HTML characters
	input = Replace(input, "&", "", 1, -1, 1)
	input = Replace(input, "<", "", 1, -1, 1)
	input = Replace(input, ">", "", 1, -1, 1)
	input = Replace(input, "'", "", 1, -1, 1)
	input = Replace(input, """", "", 1, -1, 1)
	' Return function
	CleanInput = input
	
End Function



'-----------------------------------------------------------------------------------------
' 
'-----------------------------------------------------------------------------------------
' Funzione: 
' Date: 	
' Comment:	
'-----------------------------------------------------------------------------------------
Function getContinent(country)
	
	Dim strTmp
	
	Select Case country
		Case "AFGHANISTAN"
			strTmp = TXT_Asia
		Case "ALBANIA"
			strTmp = TXT_Europe
		Case "ALGERIA"
			strTmp = TXT_Africa
		Case "AMERICAN SAMOA"
			strTmp = TXT_America
		Case "ANDORRA"
			strTmp = TXT_Europe
		Case "ANGOLA"
			strTmp = TXT_Africa
		Case "ANTIGUA AND BARBUDA"
			strTmp = TXT_America
		Case "ARGENTINA"
			strTmp = TXT_America
		Case "ARMENIA"
			strTmp = TXT_Asia
		Case "AUSTRALIA"
			strTmp = TXT_Oceania
		Case "AUSTRIA"
			strTmp = TXT_Europe
		Case "AZERBAIJAN"
			strTmp = TXT_Asia
		Case "BAHAMAS"
			strTmp = TXT_America
		Case "BAHRAIN"
			strTmp = TXT_Asia
		Case "BANGLADESH"
			strTmp = TXT_Asia
		Case "BARBADOS"
			strTmp = TXT_America
		Case "BELARUS"
'			strTmp = ""
		Case "BELGIUM"
			strTmp = TXT_Europe
		Case "BELIZE"
			strTmp = TXT_America
		Case "BENIN"
			strTmp = TXT_Africa
		Case "BERMUDA"
			strTmp = TXT_America
		Case "BHUTAN"
			strTmp = TXT_Asia
		Case "BOLIVIA"
			strTmp = TXT_America
		Case "BOSNIA AND HERZEGOVINA"
			strTmp = TXT_Europe
		Case "BOTSWANA"
			strTmp = TXT_Africa
		Case "BRAZIL"
			strTmp = TXT_America
		Case "BRITISH INDIAN OCEAN TERRITORY"
'			strTmp = ""
		Case "BRUNEI DARUSSALAM"
			strTmp = TXT_Asia
		Case "BULGARIA"
			strTmp = TXT_Europe
		Case "BURKINA FASO"
			strTmp = TXT_Africa
		Case "BURUNDI"
			strTmp = TXT_Africa
		Case "CAMBODIA"
			strTmp = TXT_Asia
		Case "CAMEROON"
			strTmp = TXT_Africa
		Case "CANADA"
			strTmp = TXT_America
		Case "CAPE VERDE"
			strTmp = TXT_Africa
		Case "CAYMAN ISLANDS"
			strTmp = TXT_America
		Case "CENTRAL AFRICAN REPUBLIC"
			strTmp = TXT_Africa
		Case "CHAD"
			strTmp = TXT_Africa
		Case "CHILE"
			strTmp = TXT_America
		Case "CHINA"
			strTmp = TXT_Asia
		Case "COLOMBIA"
			strTmp = TXT_America
		Case "COMOROS"
			strTmp = TXT_Africa
		Case "CONGO"
			strTmp = TXT_Asia
		Case "COOK ISLANDS"
			strTmp = TXT_Oceania
		Case "COSTA RICA"
			strTmp = TXT_America
		Case "COTE D'IVOIRE"
			strTmp = TXT_Africa
		Case "CROATIA"
			strTmp = TXT_Europe
		Case "CUBA"
			strTmp = TXT_America
		Case "CYPRUS"
			strTmp = TXT_Europe
		Case "CZECH REPUBLIC"
			strTmp = TXT_Europe
		Case "DENMARK"
			strTmp = TXT_Europe
		Case "DJIBOUTI"
			strTmp = TXT_Africa
		Case "DOMINICAN REPUBLIC"
			strTmp = TXT_America
		Case "EAST TIMOR"
'			strTmp = ""
		Case "ECUADOR"
			strTmp = TXT_America
		Case "EGYPT"
			strTmp = TXT_Africa
		Case "EL SALVADOR"
			strTmp = TXT_America
		Case "EQUATORIAL GUINEA"
			strTmp = TXT_Africa
		Case "ERITREA"
			strTmp = TXT_Africa
		Case "ESTONIA"
			strTmp = TXT_Europe
		Case "ETHIOPIA"
			strTmp = TXT_Africa
		Case "FALKLAND ISLANDS (MALVINAS)"
			strTmp = TXT_America
		Case "FAROE ISLANDS"
'			strTmp = ""
		Case "FIJI"
			strTmp = TXT_Oceania
		Case "FINLAND"
			strTmp = TXT_Europe
		Case "FRANCE"
			strTmp = TXT_Europe
		Case "FRENCH POLYNESIA"
'			strTmp = ""
		Case "GABON"
			strTmp = TXT_Africa
		Case "GAMBIA"
			strTmp = TXT_Africa
		Case "GEORGIA"
'			strTmp = ""
		Case "GERMANY"
			strTmp = TXT_Europe
		Case "GHANA"
			strTmp = TXT_Africa
		Case "GIBRALTAR"
'			strTmp = ""
		Case "GREECE"
			strTmp = TXT_Europe
		Case "GREENLAND"
'			strTmp = ""
		Case "GRENADA"
			strTmp = TXT_America
		Case "GUADELOUPE"
			strTmp = TXT_America
		Case "GUAM"
			strTmp = TXT_America
		Case "GUATEMALA"
			strTmp = TXT_America
		Case "GUINEA"
			strTmp = TXT_Africa
		Case "GUINEA-BISSAU"
			strTmp = TXT_Africa
		Case "HAITI"
			strTmp = TXT_America
		Case "HOLY SEE (VATICAN CITY STATE)"
			strTmp = TXT_Europe
		Case "HONDURAS"
			strTmp = TXT_America
		Case "HONG KONG"
			strTmp = TXT_Asia
		Case "HUNGARY"
			strTmp = TXT_Europe
		Case "ICELAND"
			strTmp = TXT_Europe
		Case "INDIA"
			strTmp = TXT_Asia
		Case "INDONESIA"
			strTmp = TXT_Asia
		Case "IRAQ"
			strTmp = TXT_Asia
		Case "IRELAND"
			strTmp = TXT_Europe
		Case "ISLAMIC REPUBLIC OF IRAN"
			strTmp = TXT_Asia
		Case "ISRAEL"
			strTmp = TXT_Asia
		Case "ITALY"
			strTmp = TXT_Europe
		Case "JAMAICA"
			strTmp = TXT_America
		Case "JAPAN"
			strTmp = TXT_Asia
		Case "JORDAN"
			strTmp = TXT_Asia
		Case "KAZAKHSTAN"
			strTmp = TXT_Asia
		Case "KENYA"
			strTmp = TXT_Africa
		Case "KIRIBATI"
'			strTmp = ""
		Case "KUWAIT"
			strTmp = TXT_Asia
		Case "KYRGYZSTAN"
'			strTmp = ""
		Case "LAO PEOPLE'S DEMOCRATIC REPUBL"
			strTmp = TXT_Asia
		Case "LATVIA"
'			strTmp = ""
		Case "LEBANON"
			strTmp = TXT_Asia
		Case "LESOTHO"
			strTmp = TXT_Africa
		Case "LIBERIA"
			strTmp = TXT_Africa
		Case "LIBYAN ARAB JAMAHIRIYA"
			strTmp = TXT_Africa
		Case "LIECHTENSTEIN"
			strTmp = TXT_Europe
		Case "LITHUANIA"
			strTmp = TXT_Europe
		Case "LUXEMBOURG"
			strTmp = TXT_Europe
		Case "MACAO"
			strTmp = TXT_Asia
		Case "MADAGASCAR"
			strTmp = TXT_Africa
		Case "MALAWI"
			strTmp = TXT_Africa
		Case "MALAYSIA"
			strTmp = TXT_Asia
		Case "MALDIVES"
			strTmp = TXT_Asia
		Case "MALI"
			strTmp = TXT_Africa
		Case "MALTA"
			strTmp = TXT_Europe
		Case "MARTINIQUE"
			strTmp = TXT_America
		Case "MAURITANIA"
			strTmp = TXT_Africa
		Case "MAURITIUS"
'			strTmp = ""
		Case "MEXICO"
			strTmp = TXT_America
		Case "MONACO"
			strTmp = TXT_Europe
		Case "MONGOLIA"
			strTmp = TXT_Asia
		Case "MOROCCO"
			strTmp = TXT_Africa
		Case "MOZAMBIQUE"
			strTmp = TXT_Africa
		Case "MYANMAR"
'			strTmp = ""
		Case "NAMIBIA"
			strTmp = TXT_Africa
		Case "NAURU"
			strTmp = TXT_America
		Case "NEPAL"
			strTmp = TXT_Asia
		Case "NETHERLANDS"
			strTmp = TXT_Europe
		Case "NETHERLANDS ANTILLES"
			strTmp = TXT_America
		Case "NEW CALEDONIA"
			strTmp = TXT_Oceania
		Case "NEW ZEALAND"
			strTmp = TXT_Oceania
		Case "NICARAGUA"
			strTmp = TXT_America
		Case "NIGER"
			strTmp = TXT_Africa
		Case "NIGERIA"
			strTmp = TXT_Africa
		Case "NORTHERN MARIANA ISLANDS"
'			strTmp = ""
		Case "NORWAY"
			strTmp = TXT_Europe
		Case "OMAN"
			strTmp = TXT_Asia
		Case "PAKISTAN"
			strTmp = TXT_Asia
		Case "PALAU"
			strTmp = TXT_Asia
		Case "PALESTINIAN TERRITORY, OCCUPIE"
			strTmp = TXT_Asia
		Case "PANAMA"
			strTmp = TXT_America
		Case "PAPUA NEW GUINEA"
			strTmp = TXT_Asia
		Case "PARAGUAY"
			strTmp = TXT_America
		Case "PERU"
			strTmp = TXT_America
		Case "PHILIPPINES"
			strTmp = TXT_Asia
		Case "POLAND"
			strTmp = TXT_Europe
		Case "PORTUGAL"
			strTmp = TXT_Europe
		Case "PUERTO RICO"
			strTmp = TXT_America
		Case "QATAR"
			strTmp = TXT_Asia
		Case "REPUBLIC OF KOREA"
			strTmp = TXT_Asia
		Case "REPUBLIC OF MOLDOVA"
'			strTmp = ""
		Case "REUNION"
			strTmp = TXT_Africa
		Case "ROMANIA"
			strTmp = TXT_Europe
		Case "RUSSIAN FEDERATION"
			strTmp = TXT_Asia
		Case "RWANDA"
			strTmp = TXT_Africa
		Case "SAMOA"
			strTmp = TXT_Oceania
		Case "SAN MARINO"
			strTmp = TXT_Europe
		Case "SAO TOME AND PRINCIPE"
			strTmp = TXT_Africa
		Case "SAUDI ARABIA"
			strTmp = TXT_Asia
		Case "SENEGAL"
			strTmp = TXT_Africa
		Case "SERBIA AND MONTENEGRO"
			strTmp = TXT_Europe
		Case "SEYCHELLES"
			strTmp = TXT_Africa
		Case "SIERRA LEONE"
			strTmp = TXT_Africa
		Case "SINGAPORE"
			strTmp = TXT_Asia
		Case "SLOVAKIA"
			strTmp = TXT_Europe
		Case "SLOVENIA"
			strTmp = TXT_Europe
		Case "SOLOMON ISLANDS"
			strTmp = TXT_Oceania
		Case "SOMALIA"
			strTmp = TXT_Africa
		Case "SOUTH AFRICA"
			strTmp = TXT_Africa
		Case "SPAIN"
			strTmp = TXT_Europe
		Case "SRI LANKA"
			strTmp = TXT_Asia
		Case "SUDAN"
			strTmp = TXT_Africa
		Case "SURINAME"
			strTmp = TXT_America
		Case "SWAZILAND"
			strTmp = TXT_Africa
		Case "SWEDEN"
			strTmp = TXT_Europe
		Case "SWITZERLAND"
			strTmp = TXT_Europe
		Case "SYRIAN ARAB REPUBLIC"
			strTmp = TXT_Asia
		Case "TAIWAN"
			strTmp = TXT_Asia
		Case "TAJIKISTAN"
			strTmp = TXT_Asia
		Case "THAILAND"
			strTmp = TXT_Asia
		Case "THE DEMOCRATIC REPUBLIC OF THE"
'			strTmp = ""
		Case "THE FORMER YUGOSLAV REPUBLIC O"
'			strTmp = ""
		Case "TOGO"
			strTmp = TXT_Africa
		Case "TOKELAU"
			strTmp = TXT_Oceania
		Case "TONGA"
			strTmp = TXT_Oceania
		Case "TRINIDAD AND TOBAGO"
			strTmp = TXT_America
		Case "TUNISIA"
			strTmp = TXT_Africa
		Case "TURKEY"
			strTmp = TXT_Asia
		Case "TURKMENISTAN"
'			strTmp = ""
		Case "TUVALU"
			strTmp = TXT_Oceania
		Case "UGANDA"
			strTmp = TXT_Africa
		Case "UKRAINE"
			strTmp = TXT_Europe
		Case "UNITED ARAB EMIRATES"
			strTmp = TXT_Asia
		Case "UNITED KINGDOM"
			strTmp = TXT_America
		Case "UNITED REPUBLIC OF TANZANIA"
			strTmp = TXT_Africa
		Case "UNITED STATES"
			strTmp = TXT_America
		Case "URUGUAY"
			strTmp = TXT_America
		Case "UZBEKISTAN"
'			strTmp = ""
		Case "VANUATU"
			strTmp = TXT_Oceania
		Case "VENEZUELA"
			strTmp = TXT_America
		Case "VIET NAM"
			strTmp = TXT_Asia
		Case "VIRGIN ISLANDS, BRITISH"
'			strTmp = ""
		Case "WESTERN SAHARA"
			strTmp = TXT_Asia
		Case "YEMEN"
			strTmp = TXT_Asia
		Case "ZAMBIA"
			strTmp = TXT_Africa
		Case "ZIMBABWE"
			strTmp = TXT_Africa
	End Select
	
	getContinent = strTmp
	
End Function

%>