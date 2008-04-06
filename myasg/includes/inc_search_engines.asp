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
' * @copyright       2003-2008 Simone Carletti
' * @license         http://www.weppos.com/asg/en/license.asp
' * @version         SVN: $Id$
' */
 
'/* 
' * Any disagreement of this license behaves the removal of rights to use this application.
' * Licensor reserve the right to bring legal action in the event of a violation of this Agreement.
' */


'Dichiara array motori
Dim aryEngine(120,5)


'Elenca motori in un array da controllare

aryEngine(1,1) = "http://abcsearch.com/"
aryEngine(1,2) = "abcsearch.com"
aryEngine(1,3) = "terms="
aryEngine(1,4) = "page="
aryEngine(1,5) = 1

aryEngine(2,1) = "http://search.abacho.com/"
aryEngine(2,2) = "Abacho"
aryEngine(2,3) = "q="
aryEngine(2,4) = "StartCounter="
aryEngine(2,5) = 2


aryEngine(3,1) = "http://100links.supereva.it/"
aryEngine(3,2) = "100Links"
aryEngine(3,3) = "q="
aryEngine(3,4) = "pag="
aryEngine(3,5) = 1

aryEngine(4,1) = "http://alltheweb.com/"
aryEngine(4,2) = "Alltheweb"
aryEngine(4,3) = "q="
aryEngine(4,4) = "o="
aryEngine(4,5) = 2

aryEngine(5,1) = "http://altavista.com/"
aryEngine(5,2) = "Altavista.com"
aryEngine(5,3) = "q="
aryEngine(5,4) = "stq="
aryEngine(5,5) = 2

aryEngine(6,1) = "http://aol.com/"
aryEngine(6,2) = "AOL"
aryEngine(6,3) = "query="
aryEngine(6,4) = ""
aryEngine(6,5) = 0

aryEngine(7,1) = "http://ask.com/"
aryEngine(7,2) = "Ask.com"
aryEngine(7,3) = "ask="
aryEngine(7,4) = "page="
aryEngine(7,5) = 1

aryEngine(8,1) = "http://arianna.libero.it/"
aryEngine(8,2) = "Arianna"
aryEngine(8,3) = "query="
aryEngine(8,4) = "pag="
aryEngine(8,5) = 1

aryEngine(9,1) = "http://search.dmoz.org/"
aryEngine(9,2) = "DMOZ"
aryEngine(9,3) = "search="
aryEngine(9,4) = ""
aryEngine(9,5) = 0

aryEngine(10,1) = "http://dogpile.com/"
aryEngine(10,2) = "Dogpile.com"
aryEngine(10,3) = "q="
aryEngine(10,4) = ""
aryEngine(10,5) = 0

aryEngine(11,1) = "http://www.excite.it/"
aryEngine(11,2) = "Excite.it"
aryEngine(11,3) = "q="
aryEngine(11,4) = "offset="
aryEngine(11,5) = 3

aryEngine(12,1) = "http://msxml.excite.com/"
aryEngine(12,2) = "Excite.com"
aryEngine(12,3) = "qkw="
aryEngine(12,4) = ""
aryEngine(12,5) = 0

aryEngine(13,1) = "http://www.godago.it/"
aryEngine(13,2) = "Godago"
aryEngine(13,3) = "keywords="

aryEngine(14,1) = "http://www.google.at/"
aryEngine(14,2) = "Google.at"
aryEngine(14,3) = "q="
aryEngine(14,4) = "start="
aryEngine(14,5) = 2

aryEngine(15,1) = "http://www.google.be/"
aryEngine(15,2) = "Google.be"
aryEngine(15,3) = "q="
aryEngine(15,4) = "start="
aryEngine(15,5) = 2

aryEngine(16,1) = "http://www.google.ca/"
aryEngine(16,2) = "Google.ca"
aryEngine(16,3) = "q="
aryEngine(16,4) = "start="
aryEngine(16,5) = 2

aryEngine(17,1) = "http://www.google.co.il/"
aryEngine(17,2) = "Google.co.il"
aryEngine(17,3) = "q="
aryEngine(17,4) = "start="
aryEngine(17,5) = 2

aryEngine(18,1) = "http://www.google.co.jp/"
aryEngine(18,2) = "Google.co.jp"
aryEngine(18,3) = "q="
aryEngine(18,4) = "start="
aryEngine(18,5) = 2

aryEngine(19,1) = "http://www.google.co.hu/"
aryEngine(19,2) = "Google.co.hu"
aryEngine(19,3) = "q="
aryEngine(19,4) = "start="
aryEngine(19,5) = 2

aryEngine(20,1) = "http://www.google.co.kr/"
aryEngine(20,2) = "Google.co.kr"
aryEngine(20,3) = "q="
aryEngine(20,4) = "start="
aryEngine(20,5) = 2

aryEngine(21,1) = "http://www.google.co.nz/"
aryEngine(21,2) = "Google.co.nz"
aryEngine(21,3) = "q="
aryEngine(21,4) = "start="
aryEngine(21,5) = 2

aryEngine(22,1) = "http://www.google.co.th/"
aryEngine(22,2) = "Google.co.th"
aryEngine(22,3) = "q="
aryEngine(22,4) = "start="
aryEngine(22,5) = 2

aryEngine(23,1) = "http://www.google.co.uk/"
aryEngine(23,2) = "Google.co.uk"
aryEngine(23,3) = "q="
aryEngine(23,4) = "start="
aryEngine(23,5) = 2

aryEngine(24,1) = "http://www.google.com.ar/"
aryEngine(24,2) = "Google.com.ar"
aryEngine(24,3) = "q="
aryEngine(24,4) = "start="
aryEngine(24,5) = 2

aryEngine(25,1) = "http://www.google.com.br/"
aryEngine(25,2) = "Google.com.br"
aryEngine(25,3) = "q="
aryEngine(25,4) = "start="
aryEngine(25,5) = 2

aryEngine(26,1) = "http://www.google.com.au/"
aryEngine(26,2) = "Google.com.au"
aryEngine(26,3) = "q="
aryEngine(26,4) = "start="
aryEngine(26,5) = 2

aryEngine(27,1) = "http://www.google.com.mt/"
aryEngine(27,2) = "Google.com.mt"
aryEngine(27,3) = "q="
aryEngine(27,4) = "start="
aryEngine(27,5) = 2

aryEngine(28,1) = "http://www.google.com.pe/"
aryEngine(28,2) = "Google.com.pe"
aryEngine(28,3) = "q="
aryEngine(28,4) = "start="
aryEngine(28,5) = 2

aryEngine(29,1) = "http://www.google.com.ru/"
aryEngine(29,2) = "Google.com.ru"
aryEngine(29,3) = "q="
aryEngine(29,4) = "start="
aryEngine(29,5) = 2

aryEngine(30,1) = "http://www.google.com/"
aryEngine(30,2) = "Google.com"
aryEngine(30,3) = "q="
aryEngine(30,4) = "start="
aryEngine(30,5) = 2

aryEngine(31,1) = "http://www.google.ch/"
aryEngine(31,2) = "Google.ch"
aryEngine(31,3) = "q="
aryEngine(31,4) = "start="
aryEngine(31,5) = 2

aryEngine(32,1) = "http://www.google.cl/"
aryEngine(32,2) = "Google.cl"
aryEngine(32,3) = "q="
aryEngine(32,4) = "start="
aryEngine(32,5) = 2

aryEngine(33,1) = "http://www.google.de/"
aryEngine(33,2) = "Google.de"
aryEngine(33,3) = "q="
aryEngine(33,4) = "start="
aryEngine(33,5) = 2

aryEngine(34,1) = "http://www.google.fi/"
aryEngine(34,2) = "Google.fi"
aryEngine(34,3) = "q="
aryEngine(34,4) = "start="
aryEngine(34,5) = 2

aryEngine(35,1) = "http://www.google.fr/"
aryEngine(35,2) = "Google.fr"
aryEngine(35,3) = "q="
aryEngine(35,4) = "start="
aryEngine(35,5) = 2

aryEngine(36,1) = "http://www.google.it/"
aryEngine(36,2) = "Google.it"
aryEngine(36,3) = "q="
aryEngine(36,4) = "start="
aryEngine(36,5) = 2

aryEngine(37,1) = "http://www.google.lt/"
aryEngine(37,2) = "Google.lt"
aryEngine(37,3) = "q="
aryEngine(37,4) = "start="
aryEngine(37,5) = 2

aryEngine(38,1) = "http://www.google.lv/"
aryEngine(38,2) = "Google.lv"
aryEngine(38,3) = "q="
aryEngine(38,4) = "start="
aryEngine(38,5) = 2

aryEngine(39,1) = "http://www.google.nl/"
aryEngine(39,2) = "Google.nl"
aryEngine(39,3) = "q="
aryEngine(39,4) = "start="
aryEngine(39,5) = 2

aryEngine(40,1) = "http://www.google.pl/"
aryEngine(40,2) = "Google.pl"
aryEngine(40,3) = "q="
aryEngine(40,4) = "start="
aryEngine(40,5) = 2

aryEngine(41,1) = "http://www.google.pt/"
aryEngine(41,2) = "Google.pt"
aryEngine(41,3) = "q="
aryEngine(41,4) = "start="
aryEngine(41,5) = 2

aryEngine(42,1) = "http://google.icq.com/"
aryEngine(42,2) = "Google ICQ"
aryEngine(42,3) = "q="
aryEngine(42,4) = "start="
aryEngine(42,5) = 2

aryEngine(43,1) = "http://www.google.ie/"
aryEngine(43,2) = "Google.ie"
aryEngine(43,3) = "q="
aryEngine(43,4) = "start="
aryEngine(43,5) = 2

aryEngine(44,1) = "http://www.google.com.tw/"
aryEngine(44,2) = "Google.com.tw"
aryEngine(44,3) = "q="
aryEngine(44,4) = "start="
aryEngine(44,5) = 2

aryEngine(45,1) = "http://groups.google.it/"
aryEngine(45,2) = "Google Groups"
aryEngine(45,3) = "q="
aryEngine(45,4) = "start="
aryEngine(45,5) = 2

aryEngine(46,1) = "http://categorie.iltrovatore.it/"
aryEngine(46,2) = "Il Trovatore"
aryEngine(46,3) = "query="
aryEngine(46,4) = "nh="
aryEngine(46,5) = 1

aryEngine(47,1) = "http://search.iltrovatore.it/"
aryEngine(47,2) = "Il Trovatore"
aryEngine(47,3) = "q="
aryEngine(47,4) = "np="
aryEngine(47,5) = 1

aryEngine(48,1) = "http://ixquick.com/"
aryEngine(48,2) = "Ixquick"
aryEngine(48,3) = "query="
aryEngine(48,4) = "startat="
aryEngine(48,5) = 2

aryEngine(49,1) = "http://cerca.lycos.it/"
aryEngine(49,2) = "Lycos.it"
aryEngine(49,3) = "q="
aryEngine(49,4) = "pag="
aryEngine(49,5) = 4

aryEngine(50,1) = "http://cerca.lycos.it/"
aryEngine(50,2) = "Lycos.it"
aryEngine(50,3) = "query="
aryEngine(50,4) = "pag="
aryEngine(50,5) = 4

aryEngine(51,1) = "http://search.lycos.com/"
aryEngine(51,2) = "Lycos.com"
aryEngine(51,3) = "query="
aryEngine(51,4) = "first="
aryEngine(51,5) = 2

aryEngine(52,1) = "http://vachercher.lycos.fr/"
aryEngine(52,2) = "Lycos.fr"
aryEngine(52,3) = "query="
aryEngine(52,4) = "pag="
aryEngine(52,5) = 4

aryEngine(53,1) = "http://suche.lycos.de/"
aryEngine(53,2) = "Lycos.de"
aryEngine(53,3) = "query="
aryEngine(53,4) = "pag="
aryEngine(53,5) = 4

aryEngine(54,1) = "http://www.metacrawler.com/"
aryEngine(54,2) = "Metacrawler"
aryEngine(54,3) = ""
aryEngine(54,4) = ""
aryEngine(54,5) = 0

aryEngine(55,1) = "http://search.netscape.com/"
aryEngine(55,2) = "Netscape Search"
aryEngine(55,3) = "query="
aryEngine(55,4) = "page="
aryEngine(55,5) = 1

aryEngine(56,1) = "http://search.msn.com/"
aryEngine(56,2) = "MSN.com"
aryEngine(56,3) = "q="
aryEngine(56,4) = "pn="
aryEngine(56,5) = 4

aryEngine(57,1) = "http://www.search.ch/"
aryEngine(57,2) = "Search.ch"
aryEngine(57,3) = "q="
aryEngine(57,4) = "rank="
aryEngine(57,5) = 2

aryEngine(58,1) = "http://www.search.com/"
aryEngine(58,2) = "Search.com"
aryEngine(58,3) = "q="
aryEngine(58,4) = "page="
aryEngine(58,5) = 4

aryEngine(59,1) = "http://search.supereva.it/"
aryEngine(59,2) = "Supereva"
aryEngine(59,3) = "q="
aryEngine(59,4) = "start="
aryEngine(59,5) = 2

aryEngine(60,1) = "http://search-dyn.tiscali.it/"
aryEngine(60,2) = "Tiscali"
aryEngine(60,3) = "key="
aryEngine(60,4) = "pg="
aryEngine(60,5) = 1

aryEngine(61,1) = "http://search.virgilio.it/"
aryEngine(61,2) = "Virgilio"
aryEngine(61,3) = "qs="
aryEngine(61,4) = "offset="
aryEngine(61,5) = 2

aryEngine(62,1) = "http://search.ke.voila.fr/"
aryEngine(62,2) = "Voilà.fr"
aryEngine(62,3) = "kw="
aryEngine(62,4) = "ap="
aryEngine(62,5) = 1

aryEngine(63,1) = "http://www.hotbot.com/"
aryEngine(63,2) = "HotBot"
aryEngine(63,3) = "query="
aryEngine(63,4) = "first="
aryEngine(63,5) = 2

aryEngine(64,1) = "http://search.yahoo.com/"
aryEngine(64,2) = "Yahoo.com"
aryEngine(64,3) = "p="
aryEngine(64,4) = "b="
aryEngine(64,5) = 5

'//	Verificare
' http://it.search.yahoo.com/search/it?x=wrb&va=ristorante%20c acciani&y=y&ei=UTF-8&fr=fp-tab-web-t&ve= 
aryEngine(65,1) = "http://it.search.yahoo.com/"
aryEngine(65,2) = "Yahoo.it"
aryEngine(65,3) = "p="
aryEngine(65,4) = "b="
aryEngine(65,5) = 5

aryEngine(66,1) = "http://fr.search.yahoo.com/"
aryEngine(66,2) = "Yahoo.fr"
aryEngine(66,3) = "p="
aryEngine(66,4) = "b="
aryEngine(66,5) = 5

aryEngine(67,1) = "http://de.search.yahoo.com/"
aryEngine(67,2) = "Yahoo.de"
aryEngine(67,3) = "p="
aryEngine(67,4) = "b="
aryEngine(67,5) = 5

aryEngine(68,1) = "http://bismark.caltanet.it/"
aryEngine(68,2) = "Bismark.it"
aryEngine(68,3) = "query="

aryEngine(69,1) = "http://www.kataweb.it/"
aryEngine(69,2) = "Kataweb"
aryEngine(69,3) = "q="
aryEngine(69,4) = "start="
aryEngine(69,5) = 2

aryEngine(70,1) = "http://search.clarence.com/"
aryEngine(70,2) = "Clarence"
aryEngine(70,3) = "query="
aryEngine(70,4) = "page="
aryEngine(70,5) = 1

aryEngine(71,1) = "http://paginegialle.virgilio.it/"
aryEngine(71,2) = "Pagine Gialle"
aryEngine(71,3) = "qs="
aryEngine(70,4) = "vrs="
aryEngine(70,5) = 2

aryEngine(72,1) = "http://www.jumpy.it/"
aryEngine(72,2) = "Jumpy"
aryEngine(72,3) = "searchWord="
aryEngine(72,4) = "offset="
aryEngine(72,5) = 2

aryEngine(73,1) = "http://www.italiapuntonet.net"
aryEngine(73,2) = "ItaliaPuntoNet"
aryEngine(73,3) = "search="
aryEngine(73,4) = "PagePosition="
aryEngine(73,5) = 1

'20.11.2003 Google Turchia
aryEngine(74,1) = "http://www.google.com.tr/"
aryEngine(74,2) = "Google.com.tr"
aryEngine(74,3) = "q="
aryEngine(74,4) = "start="
aryEngine(74,5) = 2

'22.11.2003 MSN Italia
aryEngine(75,1) = "http://search.msn.it/"
aryEngine(75,2) = "MSN.it"
aryEngine(75,3) = "q="
aryEngine(75,4) = "pn="
aryEngine(75,5) = 4

'07.12.2003 Ask co.uk
aryEngine(76,1) = "http://www.ask.co.uk/"
aryEngine(76,2) = "Ask.co.uk"
aryEngine(76,3) = "q="
aryEngine(76,4) = "b="
aryEngine(76,5) = 2

'10.12.2003 Google Pannello Siti
'http://www.go.com/"
aryEngine(77,1) = "http://go.google.com/"
aryEngine(77,2) = "Google.com"
aryEngine(77,3) = "q="
aryEngine(77,4) = "start="
aryEngine(77,5) = 2

'11.12.2003 Altavista Italia
aryEngine(78,1) = "http://it.altavista.com/"
aryEngine(78,2) = "Altavista.it"
aryEngine(78,3) = "q="
aryEngine(78,4) = "stq="
aryEngine(78,5) = 2

'11.12.2003 Virgilio Italia (2)
aryEngine(79,1) = "http://csearch.virgilio.it/"
aryEngine(79,2) = "Virgilio.it"
aryEngine(79,3) = "qs="
aryEngine(79,4) = "offset="
aryEngine(79,5) = 2

'11.12.2003 Yahoo UK
aryEngine(80,1) = "http://uk.search.yahoo.com/"
aryEngine(80,2) = "Yahoo.uk"
aryEngine(80,3) = "p="
aryEngine(80,4) = "b="
aryEngine(80,5) = 5

'11.12.2003 Google Sverige
aryEngine(81,1) = "http://www.google.se/"
aryEngine(81,2) = "Google.se"
aryEngine(81,3) = "q="
aryEngine(81,4) = "start="
aryEngine(81,5) = 2

'11.12.2003 Euuu.com
aryEngine(82,1) = "http://www.euuu.com/"
aryEngine(82,2) = "Euuu.com"
aryEngine(82,3) = "query="
aryEngine(82,4) = ""
aryEngine(82,5) = 0

'11.12.2003 MyWay (By Google)
aryEngine(83,1) = "http://mysearch.myway.com/"
aryEngine(83,2) = "MyWay"
aryEngine(83,3) = "searchfor="
aryEngine(83,4) = "fr="
aryEngine(83,5) = 2

'11.12.2003 Tuttogratis
aryEngine(84,1) = "http://www.tuttogratis.it/"
aryEngine(84,2) = "Tuttogratis"
aryEngine(84,3) = "keywords="
aryEngine(84,4) = ""
aryEngine(84,5) = 0

'11.12.2003 X-Download
aryEngine(85,1) = "http://www.xdownload.it/"
aryEngine(85,2) = "X-Download"
aryEngine(85,3) = "keyword="
aryEngine(85,4) = ""
aryEngine(85,5) = 0


'12.12.2003 Excite Inghilterra
aryEngine(86,1) = "http://www.excite.co.uk/"
aryEngine(86,2) = "Excite.co.uk"
aryEngine(86,3) = "q="
aryEngine(86,4) = "offset="
aryEngine(86,5) = 2

'12.12.2003 Excite Spagna
aryEngine(87,1) = "http://www.excite.es/"
aryEngine(87,2) = "Excite.es"
aryEngine(87,3) = "q="
aryEngine(87,4) = "offset="
aryEngine(87,5) = 6

'12.12.2003 Excite Germania
aryEngine(88,1) = "http://www.excite.de/"
aryEngine(88,2) = "Excite.de"
aryEngine(88,3) = "q="
aryEngine(88,4) = ""
aryEngine(88,5) = 0

'12.12.2003 Excite Francia
aryEngine(89,1) = "http://www.excite.fr/"
aryEngine(89,2) = "Excite.fr"
aryEngine(89,3) = "q="
aryEngine(89,4) = ""
aryEngine(89,5) = 0

'12.12.2003 Excite Giappone
aryEngine(90,1) = "http://www.excite.co.jp/"
aryEngine(90,2) = "Excite.co.jp"
aryEngine(90,3) = "q="
aryEngine(90,4) = "start="
aryEngine(90,5) = 2

'12.12.2003 Excite Olanda
aryEngine(91,1) = "http://www.excite.nl/"
aryEngine(91,2) = "Excite.nl"
aryEngine(91,3) = "q="
aryEngine(91,4) = "offset="
aryEngine(91,5) = 2

'12.12.2003 Excite Austria
aryEngine(92,1) = "http://www.excite.at/"
aryEngine(92,2) = "Excite.at"
aryEngine(92,3) = "q="
aryEngine(92,4) = "offset="
aryEngine(92,5) = 2

'12.12.2003 Inktomi.com
aryEngine(93,1) = "http://sitesearch.inktomi.com/"
aryEngine(93,2) = "Inktomi.com"
aryEngine(93,3) = "qt="
aryEngine(93,4) = ""
aryEngine(93,5) = 0

'12.12.2003 Teoma.com
aryEngine(94,1) = "http://s.teoma.com/"
aryEngine(94,2) = "Teoma.com"
aryEngine(94,3) = "q="
aryEngine(94,4) = ""
aryEngine(94,5) = 0

'16.12.2003 Google Messico
aryEngine(95,1) = "http://www.google.com.mx/"
aryEngine(95,2) = "Google.mx"
aryEngine(95,3) = "q="
aryEngine(95,4) = "start="
aryEngine(95,5) = 2

'16.12.2003 Google Italia
aryEngine(96,1) = "http://www.gogle.it/"
aryEngine(96,2) = "Google.it"
aryEngine(96,3) = "q="
aryEngine(96,4) = "start="
aryEngine(96,5) = 2

'16.12.2003 Altavista
aryEngine(97,1) = "http://www.altavista.com/"
aryEngine(97,2) = "Altavista.com"
aryEngine(97,3) = "q="
aryEngine(97,4) = "stq="
aryEngine(97,5) = 2

'16.12.2003 MSN Spagna
aryEngine(98,1) = "http://search.msn.es/"
aryEngine(98,2) = "MSN.es"
aryEngine(98,3) = "q="
aryEngine(98,4) = "pn="
aryEngine(98,5) = 4

'16.12.2003 Seznam [Non mi chiedete che è! L'ho visitato ma ci ho capito una mazza ;oP]
aryEngine(99,1) = "http://search1.seznam.cz/"
aryEngine(99,2) = "Seznam"
aryEngine(99,3) = "w="
aryEngine(99,4) = ""
aryEngine(99,5) = 0

'16.12.2003 Alexa
aryEngine(100,1) = "http://www.alexa.com/"
aryEngine(100,2) = "Alexa"
aryEngine(100,3) = "q="
aryEngine(100,4) = "page="
aryEngine(100,5) = 1

'16.12.2003 Clix [Portogallo]
aryEngine(101,1) = "http://pesquisa.clix.pt/"
aryEngine(101,2) = "Clix"
aryEngine(101,3) = "question="
aryEngine(100,4) = "position="
aryEngine(100,5) = 1

'24.12.2003 Comet Web Search
'					<----- 16 ----->
aryEngine(102,1) = "http://search.cometsystems.com/"
aryEngine(102,2) = "Comet Web Search"
aryEngine(102,3) = "qry="
aryEngine(102,4) = "start="
aryEngine(102,5) = 2

'29.12.2003 Euroseek.com
aryEngine(103,1) = "http://usseek.com/system/search.cgi"
aryEngine(103,2) = "Euroseek.com"
aryEngine(103,3) = "string="
aryEngine(103,4) = "start="
aryEngine(103,5) = 2

'02.01.2004 Google Spagna
aryEngine(104,1) = "http://www.google.es/"
aryEngine(104,2) = "Google.es"
aryEngine(104,3) = "q="
aryEngine(104,4) = "start="
aryEngine(104,5) = 2

'04.01.2004 Searchalot.com
aryEngine(105,1) = "http://www.searchalot.com/"
aryEngine(105,2) = "Searchalot.com"
aryEngine(105,3) = "q="
aryEngine(105,4) = "page="
aryEngine(105,5) = 1

'04.01.2004 Bluewin [by Inktomi]
aryEngine(106,1) = "http://search.bluewin.ch/"
aryEngine(106,2) = "Bluewin"
aryEngine(106,3) = "qry="
aryEngine(106,4) = ""
aryEngine(106,5) = 0

'11.01.2004 Google India
aryEngine(107,1) = "http://www.google.co.in/"
aryEngine(107,2) = "Google.co.in"
aryEngine(107,3) = "q="
aryEngine(107,4) = "start="
aryEngine(107,5) = 2

'16.01.2004 Google Venezuela
aryEngine(108,1) = "http://www.google.co.ve/"
aryEngine(108,2) = "Google.co.ve"
aryEngine(108,3) = "q="
aryEngine(108,4) = "start="
aryEngine(108,5) = 2

'21.01.2004 Google Vietnam
aryEngine(109,1) = "http://www.google.com.vn/"
aryEngine(109,2) = "Google.com.vn"
aryEngine(109,3) = "q="
aryEngine(109,4) = "start="
aryEngine(109,5) = 2

'24.01.2004 Gigablast
aryEngine(110,1) = "http://www.gigablast.com/"
aryEngine(110,2) = "Gigablast"
aryEngine(110,3) = "q="
aryEngine(110,4) = "q="
aryEngine(110,5) = 2

'26.02.2004 Gomeo.it
aryEngine(111,1) = "http://www.gomeo.it/"
aryEngine(111,2) = "Gomeo.it"
aryEngine(111,3) = "keyword="
aryEngine(111,4) = ""
aryEngine(111,5) = 0
 
'16.03.2004 Google Danimarca
aryEngine(112,1) = "http://www.google.dk/"
aryEngine(112,2) = "Google.dk"
aryEngine(112,3) = "q="
aryEngine(112,4) = "start="
aryEngine(112,5) = 2
 
'17.03.2004 Google Hong Kong
aryEngine(113,1) = "http://www.google.com.hk/"
aryEngine(113,2) = "Google.com.hk"
aryEngine(113,3) = "q="
aryEngine(113,4) = "start="
aryEngine(113,5) = 2 	 
 
'04.05.2004
aryEngine(114,1) = "http://www.rapace.it/"
aryEngine(114,2) = "Rapace.it"
aryEngine(114,3) = "keywords="
aryEngine(114,4) = ""
aryEngine(114,5) = 0	 						'NON Determinabile, stringa criptata.
 
'04.05.2004
'Uguale a Tuttoricerche.it in caratteristiche e comportamento, cambia solo skin
aryEngine(115,1) = "http://www.scovato.it/"
aryEngine(115,2) = "Scovato.it"
aryEngine(115,3) = "kw="						'NON Determinabile per i risultati in prima pagina, in quando non passa in QS i valori
aryEngine(115,4) = "p="
aryEngine(115,5) = 1
'http://www.scovato.it/cerca.php
'http://www.scovato.it/cerca.php?kw=weppos&p=2 	 
 
'04.05.2004
'Uguale a Scovato.it in caratteristiche e comportamento, cambia solo skin
aryEngine(116,1) = "http://www.tuttoricerche.it/"
aryEngine(116,2) = "Tuttoricerche.it"
aryEngine(116,3) = "kw="						'NON Determinabile per i risultati in prima pagina, in quando non passa in QS i valori
aryEngine(116,4) = "p="
aryEngine(116,5) = 1
 
'06.05.2004 Mozdex - Open Source Engine
aryEngine(117,1) = "http://www.mozdex.com/"
aryEngine(117,2) = "mozDex"
aryEngine(117,3) = "query="						
aryEngine(117,4) = "start="
aryEngine(117,5) = 2
  
'13.05.2004 Google Singapore
aryEngine(118,1) = "http://www.google.com.sg/"
aryEngine(118,2) = "Google.com.sg"
aryEngine(118,3) = "q="
aryEngine(118,4) = "start="
aryEngine(118,5) = 2 	 
  
'13.05.2004 Google Romania
aryEngine(119,1) = "http://www.google.ro/"
aryEngine(119,2) = "Google.ro"
aryEngine(119,3) = "q="
aryEngine(119,4) = "start="
aryEngine(119,5) = 2 	 

'06.06.2004 Google Ukraine
aryEngine(120,1) = "http://www.google.com.ua/"
aryEngine(120,2) = "Google.com.ua"
aryEngine(120,3) = "q="
aryEngine(120,4) = "start="
aryEngine(120,5) = 2


%>