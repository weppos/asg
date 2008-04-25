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
' Definition - Do not translate!
'-----------------------------------------------------------------------------------------
Const infoAsgTypeLanguage = "english"
'-----------------------------------------------------------------------------------------


'-----------------------------------------------------------------------------------------
' Generali
'-----------------------------------------------------------------------------------------
Const strAsgTxtOrderBy = "Order by"
Const strAsgTxtURL = "URL"
Const strAsgTxtHits = "Hits"
Const strAsgTxtVisits = "Visits"
Const strAsgTxtSmHits = "Hits"
Const strAsgTxtSmVisits = "Visits"
Const strAsgTxtByMonth = "Group by Month"
Const strAsgTxtAll = "All"
Const strAsgTxtPage = "Page"
Const strAsgTxtOf = "of"
Const strAsgTxtAsc = "Ascending"
Const strAsgTxtDesc = "Descending"
Const strAsgTxtLastAccess = "Last Access"
Const strAsgTxtShow = "Show"
Const strAsgTxtNoRecordInDatabase = "Sorry, there are no current record in the database."
Const strAsgTxtGraph = "Graph"
Const strAsgTxtStats = "Statistics"
Const strAsgTxtOptions = "Options"
Const strAsgTxtGeneral = "General"
Const strAsgTxtProvenance = "Provenance"
Const strAsgTxtExtra = "Extra"
Const strAsgTxtByHits = "by Hits"
Const strAsgTxtByVisits = "by Visits"
Const strAsgTxtNum = "Num"

'-----------------------------------------------------------------------------------------
' statistiche.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtIndexReport = "Index Report"

'-----------------------------------------------------------------------------------------
' visitors.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtDate = "Date"
Const strAsgTxtGoToPage = "Go to Page"
Const strAsgTxtVisitorsDetails = "Visitors Details"

'-----------------------------------------------------------------------------------------
' browser.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtBrowser = "Browser"

'-----------------------------------------------------------------------------------------
' os.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtOS = "Operating System"

'-----------------------------------------------------------------------------------------
' color.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtReso = "Resolution"
Const strAsgTxtColor = "Color Deep"
Const strAsgTxtSmReso = "Reso"
Const strAsgTxtSmColor = "Bit"

'-----------------------------------------------------------------------------------------
' browser_lang.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtBrowserLanguages = "Browser Languages"

'-----------------------------------------------------------------------------------------
' system.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtSystems = "Systems"

'-----------------------------------------------------------------------------------------
' referer.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtReferer = "Referer"
Const strAsgTxtRefererIn = "Internal Referers"
Const strAsgTxtRefererOut = "External Referers"
Const strAsgTxtRefererAll = "All Referers"
Const strAsgTxtTypology = "Typology"

'-----------------------------------------------------------------------------------------
' engine.asp & query.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtSearchQuery = "Search Query"
Const strAsgTxtSearchEngine = "Search Engine"
Const strAsgTxtQuery = "Query"
Const strAsgTxtEngine = "Engine"

'-----------------------------------------------------------------------------------------
' country.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtCountry = "Country"

'-----------------------------------------------------------------------------------------
' ip_address.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtIP = "IP"
Const strAsgTxtIPAddress = "IP Address"
Const strAsgTxtNoInformationSelectedIP = "No available informations about selected IP!"
Const strAsgTxtShowIpInformation = "Show IP Information"

'-----------------------------------------------------------------------------------------
' accessi.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtVisitsPerDay = "Visits per Day"
Const strAsgTxtVisitsPerMonth = "Visits per Month"
Const strAsgTxtVisitsPerHour = "Visits per Hour"

'-----------------------------------------------------------------------------------------
' settings.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtGeneralSettings = "General Settings"
Const strAsgTxtSecuritySettings = "Security Settings"
Const strAsgTxtResetSettings = "Reset Settings"


' NEW FROM VERSION 1.2
'-----------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------
' login.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtLogin = "Login"
Const strAsgTxtLoginCompleted = "Login successfully completed!"
Const strAsgTxtEntryAllowed = "Entry to statt allowed"
Const strAsgTxtClickToLogout = "Click here to Logout"
Const strAsgTxtWrongPassword = "Typed password is not correct"
Const strAsgTxtTypePassword = "Type password"
Const strAsgTxtEntryPassword = "Entry Password"

'-----------------------------------------------------------------------------------------
' settings.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtSiteName = "Site Name"
Const strAsgTxtSiteURLlocal = "LOCAL Site URL"
Const strAsgTxtSiteURLremote = "REMOTE Site URL"
Const strAsgTxtSiteEmail = "Site Email"
Const strAsgTxtConfigSettings = "Configuration Settings"
Const strAsgTxtCountSettings = "Count Settings"
Const strAsgTxtMonitSettings = "Tracking Settings"
Const strAsgTxtMonitOptions = "Tracking Options"
Const strAsgTxtStartVisits = "Starting Visits"
Const strAsgTxtStartHits = "Starting Hits"
Const strAsgTxtFilterIPaddr = "Filtered IP Addresses"
Const strAsgTxtEnableMonit = "Enable Monitoring"
Const strAsgTxtCountServerAsReferer = "Count server as a Referer"
Const strAsgTxtstrAsgIPPathQS = "strAsgIP Path Querysting"
Const strAsgTxtDailyMonit = "Daily Monitoring"
Const strAsgTxtHourlyMonit = "Hourly Monitoring"

'-----------------------------------------------------------------------------------------
' settings_security.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtNewPassword = "New Password"
Const strAsgTxtConfirmPassword = "Confirm Password"
Const strAsgTxtUpdateSuccessfullyCompleted = "Update successfully completed!"
Const strAsgTxtStatsProtection = "Stats Protection"
Const strAsgTxtStatsProtectionLevel = "Protection level"
Const strAsgTxtNone = "None"
Const strAsgTxtLimited = "Limited"
Const strAsgTxtFull = "Full"
Const strAsgTypeOnlyToChangePassword = "To fill only if you want to change password!"
Const strAsgTxtAttentionPasswordNotMatching = "ATTENTION: typed passwords not matching!"


'-----------------------------------------------------------------------------------------
' settings_reset.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtErrorOccured = "An error occured!"
Const strAsgTxtCheckTableMatching = "Check that table name matches with settings configuration."
Const strAsgTxtTableReset = "Reset Table"
Const strAsgTxtDetailContent = "Contains general informations and users stats"
Const strAsgTxtSystemContent = "Contains users systems informations"
Const strAsgTxtHourlyContent = "Contains hourly traking informations"
Const strAsgTxtDailyContent = "Contains daily tracking informations"
Const strAsgTxtIPContent = "Contains IP addresses"
Const strAsgTxtLanguageContent = "Contains users browser languages informations"
Const strAsgTxtRefererContent = "Countains refereres informations"
Const strAsgTxtPageContent = "Contains pages visited by users"
Const strAsgTxtQueryContent = "Contains engines and search queries"
Const strAsgTxtResetAllTables = "Reset all Tables"


'-----------------------------------------------------------------------------------------
' settings_reset_execute.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtExecutionReport = "Execution Report"
Const strAsgTxtTable = "Table"
Const strAsgTxtCorrectlyDeleted = "correctly deleted!"
Const strAsgTxtDatabaseSuccessfullyCompactedOn = "Database successfully compacted on "
Const strAsgTxtDatabaseSuccessfullyRenamedTo = "Database successfully renamed to "

Const strAsgTxtError = "Error"
Const strAsgTxtLogout = "Logout"
Const strAsgTxtUpdate = "Update"
Const strAsgTxtAnd = "and"


'-----------------------------------------------------------------------------------------
' statistiche.asp
'-----------------------------------------------------------------------------------------
Const strAsgTxtVisitsInformations = "Visits Informations"
Const strAsgTxtGeneralInformations = "General Informations"
Const strAsgTxtGeneralAverageInformations = "General Average Informations"
Const strAsgTxtYearlyInformations = "Yearly Informations"
Const strAsgTxtToday = "Today"
Const strAsgTxtYesterday = "Yesterday"
Const strAsgTxtPerDay = "per Day"
Const strAsgTxtPerHour = "per Hour"
Const strAsgTxtLastMonth = "last Month"


' NEW FROM VERSION 1.3
'-----------------------------------------------------------------------------------------

Const strAsgTxtDetails = "Details"
Const strAsgTxtGoingToBeRedirected = "You are going to be redirected"
Const strAsgTxtClickToRedirect = "Click here if your browser doesn't redirect you"
Const strAsgTxtJanuary = "January"
Const strAsgTxtFebruary = "February"
Const strAsgTxtMarch = "March"
Const strAsgTxtApril = "April"
Const strAsgTxtMay = "May"
Const strAsgTxtJune = "June"
Const strAsgTxtJuly = "July"
Const strAsgTxtAugust = "August"
Const strAsgTxtSeptember = "September"
Const strAsgTxtOctober = "October"
Const strAsgTxtNovember = "November"
Const strAsgTxtDecember = "December"
Const strAsgTxtPath = "Path"
Const strAsgTxtGroupByPath = "Group by Path"
Const strAsgTxtGroupByPage = "Group by Page"
Const strAsgTxtGroupByEngine = "Group by Engine"
Const strAsgTxtGroupByQuery = "Group by Query"
Const strAsgTxtIPTracking = "IP Tracking"
Const strAsgTxtFor = "for"
Const strAsgTxtTime = "Time"
Const strAsgTxtMissedDataToElab = "Not enoungh information to continue elaboration process!"
Const strAsgTxtCloseWindow = "Close Window"
Const strAsgTxtView = "View"


' NEW FROM VERSION 1.4
'-----------------------------------------------------------------------------------------

Const strAsgTxtInsufficientPermission = "You don't have permissions to view this page!"
Const strAsgTxtAction = "Action"
Const strAsgTxtAddToList = "Add to List"
Const strAsgTxtResetAndAddToList = "Reset List and Add Item"
Const strAsgTxtMonthlyCalendar = "Monthly Calendar"

Const strAsgTxtSunday = "Sunday"


' NEW FROM VERSION 2.0
'-----------------------------------------------------------------------------------------
Const strAsgTxtStatsOfTheMonth = "Statistics of the Month"
Const strAsgTxtStatsOfTheYear = "Statistics of the Year"
Const strAsgTxtCalendar = "Calendar"
Const strAsgTxtDomain = "Domain"
Const strAsgTxtGroupByDomain = "Group by Domain"
Const strAsgTxtGroupByReferer = "Group by Referer"
Const strAsgTxtInformationsToExitByIpRange = "The * wildcard character can be used to block IP ranges. <br />Ex. To block the range '200.200.200.0 - 255' you would use '200.200.200.*'"
Const strAsgTxtServerinformations = "Server Informations"
Const strAsgTxtFullVersion = "Full Version"
Const strAsgTxtProgramTools = "Program Tools"
Const strAsgTxtSERPreports = "SERP Reports"
Const strAsgTxtExclusionSettings = "Stats Exclusion"
Const strAsgTxtExitByIP = "Exclusion by IP"
Const strAsgTxtExitByCookie = "Exclusion by Cookie"
Const strAsgTxtExcludePC = "Exclude PC from statistics"
Const strAsgTxtIncludePC = "Include PC in statistics"
Const strAsgTxtDateSettings = "Date Settings"
Const strAsgTxtTimeZoneOffSet = "Time zone offset"
Const strAsgTxtOffSetClientServer = "from Server time"
Const strAsgTxtOffSetServerToGMT = "from Server to Greenwich meridian (GMT)"
Const strAsgTxtOffSetGMTtoUser = "from Greenwich meridian (GMT) to your Time"
Const strAsgTxtOn = "on"
Const strAsgTxtAt = "at"
Const strAsgTxtServerDateTimeAre = "Server date and time are"
Const strAsgTxtReportDateTimeAre = "Report date and time are"
Const strAsgTxtCountryContent = "Contains informations about users country"
Const strAsgTxtMonth = "Month"
Const strAsgTxtMonths = "Monts"
Const strAsgTxtWeek = "Week"
Const strAsgTxtWeeks = "Weeks"
Const strAsgTxtDataReset = "Reset data"
Const strAsgTxtOlderThan = "older than"
Const strAsgTxtCurrent = "this"
Const strAsgTxtOnlineUsers = "Users online"
Const strAsgTxtTop = "Top"


' NEW FROM VERSION 2.1
'-----------------------------------------------------------------------------------------
Const strAsgTxtSearch = "Search"
Const strAsgTxtSearchFoundNoResults = "Sorry, your search found no results."
Const strAsgTxtAdvice = "Advice"
Const strAsgTxtTablesWithWarningIconNeedsReset = "Tables marked with the alert signal need a data reset."
Const strAsgTxtRecords = "Records"
Const strAsgTxtFrom = "from"
Const strAsgTxtTo = "to"
Const strAsgCompactDatabase = "Compact Database"
Const strAsgTxtIISversion = "IIS version"
Const strAsgTxtProtocolVersion = "Protocol version"
Const strAsgTxtYourIpIs = "Your IP is"
Const strAsgTxtServerName = "Server Name"
Const strAsgTxtThisPCisActually = "This PC is actually"
Const strAsgTxtIncluded = "included"
Const strAsgTxtExcluded = "excluded"
Const strAsgTxtIntoMonitoringProcess = "into monitoring process"
Const strAsgTxtUsingApplication = "Using application"
Const strAsgTxtPageMonitoringString = "Monitoring string to add in pages you'd like to track"


Const ASG_TEXT_PAGE_GENERATED_IN = "This page was generated in %{seconds} seconds"

%>