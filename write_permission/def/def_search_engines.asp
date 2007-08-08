<%

'-------------------------------------------------------------------------------'
'	ASP Stats Generator															'
'	Copyright © 2003-2005 Simone Carletti [ aka weppos ]						'
'-------------------------------------------------------------------------------'


' Declare an array to hold search engine definitions.
' Remember to increment the array value to add new search engines!
' The index 1 uses regular expression pattern.
Dim aryAsgEngine(6,57)

' http://www.aolrecherches.aol.fr/search?service=WebMondial&first=1&last=10&p=wf&query=bittess

' http://*google.*/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,1) = "^http://w*\.*google\.([^/]*)/"
aryAsgEngine(2,1) = "Google"
aryAsgEngine(3,1) = "q="
aryAsgEngine(4,1) = "start="
aryAsgEngine(5,1) = 2
aryAsgEngine(6,1) = "us"

' http://*search.msn.*/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,2) = "^http://\w*\.*search.msn\.([^/]*)/"
aryAsgEngine(2,2) = "MSN"
aryAsgEngine(3,2) = "q="
aryAsgEngine(4,2) = "first="
aryAsgEngine(5,2) = 7
aryAsgEngine(6,2) = "us"

' http://*search.yahoo.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,3) = "^http://(\w*)\.*search.yahoo.com/"
aryAsgEngine(2,3) = "Yahoo"
aryAsgEngine(3,3) = "p="
aryAsgEngine(4,3) = "b="
aryAsgEngine(5,3) = 7
aryAsgEngine(6,3) = "us"

' http://*altavista.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,4) = "^http://\w*\.*altavista.(com)/"
aryAsgEngine(2,4) = "Altavista"
aryAsgEngine(3,4) = "q="
aryAsgEngine(4,4) = "stq="
aryAsgEngine(5,4) = 2
aryAsgEngine(6,4) = "us"

' http://search.lycos.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,5) = "^http://search.lycos.(com)/"
aryAsgEngine(2,5) = "Lycos"
aryAsgEngine(3,5) = "query="
aryAsgEngine(4,5) = "first="
aryAsgEngine(5,5) = 7
aryAsgEngine(6,5) = "us"

' http://*.lycos.*/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,6) = "^http://\w*\.lycos\.([^/]*)/"
aryAsgEngine(2,6) = "Lycos"
aryAsgEngine(3,6) = "query="
aryAsgEngine(4,6) = "pag="
aryAsgEngine(5,6) = 4
aryAsgEngine(6,6) = "us"

' http://msxml.excite.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,7) = "^http://msxml.excite.(com)/"
aryAsgEngine(2,7) = "Excite"
aryAsgEngine(3,7) = null
aryAsgEngine(4,7) = null
aryAsgEngine(5,7) = null
aryAsgEngine(6,7) = "us"

' http://*excite.*/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,8) = "^http://w*\.*excite\.([^/]*)/"
aryAsgEngine(2,8) = "Excite"
aryAsgEngine(3,8) = "q="
aryAsgEngine(4,8) = "offset="
aryAsgEngine(5,8) = 2
aryAsgEngine(6,8) = "us"

' http://search.aol.*/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,9) = "^http://search.aol\.([^/]*)/"
aryAsgEngine(2,9) = "AOL"
aryAsgEngine(3,9) = "query="
aryAsgEngine(4,9) = "page="
aryAsgEngine(5,9) = 4
aryAsgEngine(6,9) = "us"

' http://www.alexa.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,10) = "^http://www.alexa.(com)/"
aryAsgEngine(2,10) = "Alexa"
aryAsgEngine(3,10) = "q="
aryAsgEngine(4,10) = "page="
aryAsgEngine(5,10) = 1
aryAsgEngine(6,10) = "us"

' http://s.teoma.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,11) = "^http://s.teoma.(com)/"
aryAsgEngine(2,11) = "Teoma"
aryAsgEngine(3,11) = "q="
aryAsgEngine(4,11) = "page="
aryAsgEngine(5,11) = 1
aryAsgEngine(6,11) = "us"

' http://web.ask.*/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,12) = "^http://web.ask\.([^/]*)/"
aryAsgEngine(2,12) = "Ask"
aryAsgEngine(3,12) = "q="
aryAsgEngine(4,12) = "page="
aryAsgEngine(5,12) = 1
aryAsgEngine(6,12) = "us"

' http://*alltheweb.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,13) = "^http://w*\.*alltheweb.(com)/"
aryAsgEngine(2,13) = "Alltheweb"
aryAsgEngine(3,13) = "q="
aryAsgEngine(4,13) = "o="
aryAsgEngine(5,13) = 2
aryAsgEngine(6,13) = "us"

' http://search.dmoz.org/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,14) = "^http://search.dmoz.(org)/"
aryAsgEngine(2,14) = "Dmoz"
aryAsgEngine(3,14) = "search="
aryAsgEngine(4,14) = "start="
aryAsgEngine(5,14) = 5
aryAsgEngine(6,14) = "us"

' http://*dogpile.*/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,15) = "^http://w*\.*dogpile\.([^/]*)/"
aryAsgEngine(2,15) = "Dogpile"
aryAsgEngine(3,15) = null
aryAsgEngine(4,15) = null
aryAsgEngine(5,15) = null
aryAsgEngine(6,15) = "us"

' http://*ixquick.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,16) = "^http://w*\.*ixquick.(com)/"
aryAsgEngine(2,16) = "Ixquick"
aryAsgEngine(3,16) = "query="
aryAsgEngine(4,16) = "startat="
aryAsgEngine(5,16) = 2
aryAsgEngine(6,16) = "us"

' http://search.hotbot.*/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,17) = "^http://search.hotbot\.([^/]*)/"
aryAsgEngine(2,17) = "HotBot"
aryAsgEngine(3,17) = "query="
aryAsgEngine(4,17) = "pag="
aryAsgEngine(5,17) = 2
aryAsgEngine(6,17) = "us"

' http://search.abacho.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
' Lang : in the URL /TLD/
aryAsgEngine(1,18) = "^http://search.abacho.(com)/"
aryAsgEngine(2,18) = "Abacho"
aryAsgEngine(3,18) = "q="
aryAsgEngine(4,18) = "offset="
aryAsgEngine(5,18) = 2
aryAsgEngine(6,18) = "us"

' http://*godago.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,19) = "^http://(\w*)\.*godado.com/"
aryAsgEngine(2,19) = "Godado"
aryAsgEngine(3,19) = "keywords="
aryAsgEngine(4,19) = null
aryAsgEngine(5,19) = null
aryAsgEngine(6,19) = "us"

' http://*godago.*/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,20) = "^http://w*\.*godado\.([^/]*)/"
aryAsgEngine(2,20) = "Godado"
aryAsgEngine(3,20) = "keywords="
aryAsgEngine(4,20) = null
aryAsgEngine(5,20) = null
aryAsgEngine(6,20) = "us"

' http://*abcsearch.com/
' I , 1 , 2 , 3 , 4 , 5
aryAsgEngine(1,21) = "^http://w*\.*abcsearch.(com)/"
aryAsgEngine(2,21) = "ABC Search"
aryAsgEngine(3,21) = "Terms="
aryAsgEngine(4,21) = "offset="
aryAsgEngine(5,21) = 2
aryAsgEngine(6,21) = "us"

' http://msxml.webcrawler.com/
' X , 1 , 2 , 3 , 4 , 5
aryAsgEngine(1,22) = "^http://msxml.webcrawler.(com)/"
aryAsgEngine(2,22) = "WebCrawler"
aryAsgEngine(3,22) = null
aryAsgEngine(4,22) = null
aryAsgEngine(5,22) = null
aryAsgEngine(6,22) = "us"

' http://arianna.libero.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,23) = "^http://arianna.libero.(it)/"
aryAsgEngine(2,23) = "Arianna"
aryAsgEngine(3,23) = "query="
aryAsgEngine(4,23) = "pag="
aryAsgEngine(5,23) = 1
aryAsgEngine(6,23) = "it"

' http://search-dyn.tiscali.*/
' I , 1 , 2 , 3 , 4 , 5
aryAsgEngine(1,24) = "^http://search-dyn.tiscali\.([^/]*)/"
aryAsgEngine(2,24) = "Tiscali"
aryAsgEngine(3,24) = "key="
aryAsgEngine(4,24) = "pg="
aryAsgEngine(5,24) = 1
aryAsgEngine(6,24) = "it"

' http://search-dyn.tiscali.*/
' I , 1 , 2 , 3 , 4 , 5
' Exception .fr
aryAsgEngine(1,25) = "^http://rechercher.nomade.tiscali.(fr)/"
aryAsgEngine(2,25) = "Tiscali"
aryAsgEngine(3,25) = "s="
aryAsgEngine(4,25) = "pg="
aryAsgEngine(5,25) = 1
aryAsgEngine(6,25) = "fr"

' http://search-dyn.tiscali.*/
' I , 1 , 2 , 3 , 4 , 5
' Exception .co.uk
aryAsgEngine(1,26) = "^http://www.tiscali.(co.uk)/"
aryAsgEngine(2,26) = "Tiscali"
aryAsgEngine(3,26) = "query="
aryAsgEngine(4,26) = "start="
aryAsgEngine(5,26) = 7
aryAsgEngine(6,26) = "co.uk"

' http://search.virgilio.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,27) = "^http://search.virgilio.(it)/"
aryAsgEngine(2,27) = "Virgilio"
aryAsgEngine(3,27) = "qs="
aryAsgEngine(4,27) = "offset="
aryAsgEngine(5,27) = 2
aryAsgEngine(6,27) = "it"

' http://csearch.virgilio.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,28) = "^http://csearch.virgilio.(it)/"
aryAsgEngine(2,28) = "Virgilio"
aryAsgEngine(3,28) = "qs="
aryAsgEngine(4,28) = "vrs="
aryAsgEngine(5,28) = 2
aryAsgEngine(6,28) = "it"

' http://www.metacrawler.com/
' X , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,29) = "^http://www.metacrawler.(com)/"
aryAsgEngine(2,29) = "Metacrawler"
aryAsgEngine(3,29) = null
aryAsgEngine(4,29) = null
aryAsgEngine(5,29) = null
aryAsgEngine(6,29) = "us"

' http://search.netscape.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,30) = "^http://search.netscape.(com)/"
aryAsgEngine(2,30) = "Netscape"
aryAsgEngine(3,30) = "query="
aryAsgEngine(4,30) = "page="
aryAsgEngine(5,30) = 1
aryAsgEngine(6,30) = "us"

' http://www.search.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,31) = "^http://www.search.(com)/"
aryAsgEngine(2,31) = "Search"
aryAsgEngine(3,31) = "q="
aryAsgEngine(4,31) = null
aryAsgEngine(5,31) = null
aryAsgEngine(6,31) = "us"

' http://www.search.ch/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,32) = "^http://www.search.(ch)/"
aryAsgEngine(2,32) = "Search"
aryAsgEngine(3,32) = "q="
aryAsgEngine(4,32) = "rank="
aryAsgEngine(5,32) = 2
aryAsgEngine(6,32) = "ch"

' http://www.search.it/
' X , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,33) = "^http://www.search.(it)/"
aryAsgEngine(2,33) = "Search"
aryAsgEngine(3,33) = "srctxt="
aryAsgEngine(4,33) = null
aryAsgEngine(5,33) = null
aryAsgEngine(6,33) = "it"

' http://search.clarence.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,34) = "^http://search.clarence.(com)/"
aryAsgEngine(2,34) = "Clarence"
aryAsgEngine(3,34) = "q="
aryAsgEngine(4,34) = "page="
aryAsgEngine(5,34) = 1
aryAsgEngine(6,34) = "it"

' http://www.paginegialle.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,35) = "^http://www.paginegialle.(it)/"
aryAsgEngine(2,35) = "Pagine Gialle"
aryAsgEngine(3,35) = "qs="
aryAsgEngine(4,35) = "be="
aryAsgEngine(5,35) = 2
aryAsgEngine(6,35) = "it"

' http://www.gigablast.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,36) = "^http://www.gigablast.(com)/"
aryAsgEngine(2,36) = "Gigablast"
aryAsgEngine(3,36) = "q="
aryAsgEngine(4,36) = "q="
aryAsgEngine(5,36) = 2
aryAsgEngine(6,36) = "us"

' http://www.searchalot.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,37) = "^http://www.searchalot.(com)/"
aryAsgEngine(2,37) = "Searchalot"
aryAsgEngine(3,37) = "q="
aryAsgEngine(4,37) = "page="
aryAsgEngine(5,37) = 1
aryAsgEngine(6,37) = "us"

' http://www.kataweb.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,38) = "^http://www.kataweb.(it)/"
aryAsgEngine(2,38) = "Kataweb"
aryAsgEngine(3,38) = "q="
aryAsgEngine(4,38) = "page="
aryAsgEngine(5,38) = 1
aryAsgEngine(6,38) = "it"

' http://*mysearch.myway.com/
' X , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,39) = "^http://w*\.*mywebsearch.(com)/"
aryAsgEngine(2,39) = "My Way"
aryAsgEngine(3,39) = "searchfor="
aryAsgEngine(4,39) = "fr="
aryAsgEngine(5,39) = 2
aryAsgEngine(6,39) = "us"

' http://*mysearch.myway.com/
' X , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,40) = "^http://(\w*)\.*mysearch.myway.com/"
aryAsgEngine(2,40) = "My Way"
aryAsgEngine(3,40) = "searchfor="
aryAsgEngine(4,40) = "fr="
aryAsgEngine(5,40) = 2
aryAsgEngine(6,40) = "us"

' http://www.mysearch.com/
' X , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,41) = "^http://www.mysearch.(com)/"
aryAsgEngine(2,41) = "My Search"
aryAsgEngine(3,41) = "searchfor="
aryAsgEngine(4,41) = "fr="
aryAsgEngine(5,41) = 2
aryAsgEngine(6,41) = "us"

' http://euroseek.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,42) = "^http://w*\.*euroseek.(com)/"
aryAsgEngine(2,42) = "Euroseek"
aryAsgEngine(3,42) = "string="
aryAsgEngine(4,42) = "start="
aryAsgEngine(5,42) = 2
aryAsgEngine(6,42) = "us"

' http://search.supereva.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,43) = "^http://search.supereva.(it)/"
aryAsgEngine(2,43) = "Supereva"
aryAsgEngine(3,43) = "q="
aryAsgEngine(4,43) = "start="
aryAsgEngine(5,43) = 2
aryAsgEngine(6,43) = "it"

' http://100links.supereva.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,44) = "^http://100links.supereva.(it)/"
aryAsgEngine(2,44) = "Supereva 100Links"
aryAsgEngine(3,44) = "q="
aryAsgEngine(4,44) = "pag="
aryAsgEngine(5,44) = 1
aryAsgEngine(6,44) = "it"

' http://recherche.aol.fr/
' I , 1 , 2 , 3 , 4 , 5 , 6
' Exception .fr
aryAsgEngine(1,45) = "^http://recherche.aol.(fr)/"
aryAsgEngine(2,45) = "AOL"
aryAsgEngine(3,45) = "q="
aryAsgEngine(4,45) = "s="
aryAsgEngine(5,45) = 2
aryAsgEngine(6,45) = "fr"

' http://search.ke.voila.fr/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,46) = "^http://search.ke.voila.(fr)/"
aryAsgEngine(2,46) = "Voil&agrave;"
aryAsgEngine(3,46) = "kw="
aryAsgEngine(4,46) = "ap="
aryAsgEngine(5,46) = 1
aryAsgEngine(6,46) = "fr"

' http://categorie.iltrovatore.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,47) = "^http://categorie.iltrovatore.(it)/"
aryAsgEngine(2,47) = "Il Trovatore Directory"
aryAsgEngine(3,47) = "query="
aryAsgEngine(4,47) = "nh="
aryAsgEngine(5,47) = 1
aryAsgEngine(6,47) = "it"

' http://search.iltrovatore.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,48) = "^http://search.iltrovatore.(it)/"
aryAsgEngine(2,48) = "Il Trovatore"
aryAsgEngine(3,48) = "q="
aryAsgEngine(4,48) = "ps="
aryAsgEngine(5,48) = 2
aryAsgEngine(6,48) = "it"

' http://www.jumpy.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,49) = "^http://www.jumpy.(it)/"
aryAsgEngine(2,49) = "Jumpy"
aryAsgEngine(3,49) = "searchWord="
aryAsgEngine(4,49) = "offset="
aryAsgEngine(5,49) = 2
aryAsgEngine(6,49) = "it"

' http://www.italiapuntonet.net/
' X , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,50) = "^http://www.italiapuntonet.(net)/"
aryAsgEngine(2,50) = "ItaliaPuntoNet"
aryAsgEngine(3,50) = "search="
aryAsgEngine(4,50) = "PagePosition="
aryAsgEngine(5,50) = 1
aryAsgEngine(6,50) = "it"

' http://www.tuttogratis.it/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,51) = "^http://www.tuttogratis.(it)/"
aryAsgEngine(2,51) = "Tuttogratis"
aryAsgEngine(3,51) = "k="
aryAsgEngine(4,51) = "p="
aryAsgEngine(5,51) = 1
aryAsgEngine(6,51) = "it"

' http://www.euuu.com/
' X , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,52) = "^http://www.euuu.(com)/"
aryAsgEngine(2,52) = "Euuu"
aryAsgEngine(3,52) = "query="
aryAsgEngine(4,52) = "start="
aryAsgEngine(5,52) = 1
aryAsgEngine(6,52) = "us"

' http://search.cometsystems.com/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,53) = "^http://search.cometsystems.(com)/"
aryAsgEngine(2,53) = "Comet Web Search"
aryAsgEngine(3,53) = "qry="
aryAsgEngine(4,53) = "start="
aryAsgEngine(5,53) = 2
aryAsgEngine(6,53) = "us"

' http://search1.seznam.cz/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,54) = "^http://search1.seznam.(cz)/"
aryAsgEngine(2,54) = "Seznam"
aryAsgEngine(3,54) = "w="
aryAsgEngine(4,54) = "from="
aryAsgEngine(5,54) = 2
aryAsgEngine(6,54) = "cz"

' http://pesquisa.clix.pt/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,55) = "^http://pesquisa.clix.(pt)/"
aryAsgEngine(2,55) = "Clix"
aryAsgEngine(3,55) = "question="
aryAsgEngine(4,55) = "pnacional="
aryAsgEngine(5,55) = 1
aryAsgEngine(6,55) = "pt"

' http://search.bluewin.*/
' I , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,56) = "^http://search.bluewin\.([^/]*)/"
aryAsgEngine(2,56) = "Bluewin"
aryAsgEngine(3,56) = "qry="
aryAsgEngine(4,56) = "first="
aryAsgEngine(5,56) = 2
aryAsgEngine(6,56) = "ch"

' http://www.eniro.*/
' !-! , 1 , 2 , 3 , 4 , 5 , 6
aryAsgEngine(1,57) = "^http://www.eniro.([^/]*)/"
aryAsgEngine(2,57) = "Eniro"
aryAsgEngine(3,57) = "q="
aryAsgEngine(4,57) = "stq="
aryAsgEngine(5,57) = 2
aryAsgEngine(6,57) = "se"

' http://www.mozdex.com/
' !!! , 1 , 2 , 3 , 4 , 5 , 6
'aryAsgEngine(1) = "^http://www.mozdex.(com)/"
'aryAsgEngine(2) = "mozDex"
'aryAsgEngine(3) = "query="						
'aryAsgEngine(4) = "start="
'aryAsgEngine(5) = 2
'aryAsgEngine(6) = "us"




' == Checked and not added
' http://www.alexaweb.com/
' http://www.gomeo.it/ , http://www.gomeo.com/

' == Checked and not added (Not working)
' http://www.bismark.it/

' == Checked and removed
' http://sitesearch.inktomi.com/


%>