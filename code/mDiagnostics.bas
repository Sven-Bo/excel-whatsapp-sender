Attribute VB_Name = "mDiagnostics"
'---------------------------------------------------------------------------------------
' Module    : mDiagnostics
' Author    : Sven Bosau
' Purpose   : Captures diagnostic information to help debug XPath mismatches
'             for users with different WhatsApp Web variants.
'             Outputs a formatted text file with settings, XPath match results,
'             and discovered DOM elements.
'---------------------------------------------------------------------------------------

Option Explicit

Private Const DIAG_FILE_PREFIX As String = "WhatsAppBlaster_Diagnostics_"
Private Const SEPARATOR As String = "================================================================================"

' ========================================================================================
'                              PUBLIC ENTRY POINT
' ========================================================================================

Public Sub RunDiagnostics()
    On Error GoTo ErrHandler

    Dim diagOutput As String
    Dim filePath As String
    Dim botObj As Object
    Dim byObj As Object
    Dim wsSettings As Worksheet
    Dim wsBackend As Worksheet
    Dim userResponse As VbMsgBoxResult

    ' Summary tracking
    Dim issueCount As Long
    Dim warnCount As Long
    Dim summaryText As String
    issueCount = 0
    warnCount = 0
    summaryText = ""

    ' --- Intro ---
    Dim introMsg As String
    introMsg = "IMPORTANT: Please read each step carefully and follow the instructions exactly." & vbCrLf
    introMsg = introMsg & "Each step will ask you to do something specific in WhatsApp Web." & vbCrLf
    introMsg = introMsg & "Wait for each instruction before clicking or typing anything." & vbCrLf & vbCrLf
    introMsg = introMsg & "This tool will:" & vbCrLf
    introMsg = introMsg & "  1. Open Chrome and navigate to WhatsApp Web" & vbCrLf
    introMsg = introMsg & "  2. Ask you to do 2 quick steps (search for something, open a chat)" & vbCrLf
    introMsg = introMsg & "  3. Save a diagnostics file you can send to support" & vbCrLf & vbCrLf
    introMsg = introMsg & "No messages will be sent or read. Only technical page structure" & vbCrLf
    introMsg = introMsg & "is captured. The contact name you search for may appear in the file." & vbCrLf
    introMsg = introMsg & "The file is saved locally on your computer." & vbCrLf & vbCrLf
    introMsg = introMsg & "Click OK to start."
    userResponse = MsgBox(introMsg, vbOKCancel + vbExclamation, "WhatsApp Bot - Diagnostics")

    If userResponse = vbCancel Then Exit Sub

    ' ========== SECTION 1: System & App Info ==========
    diagOutput = FormatHeader("WHATSAPP BLASTER DIAGNOSTICS REPORT")
    diagOutput = diagOutput & "Generated:     " & Format(Now, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    diagOutput = diagOutput & "App Version:   " & VERSION_NUMBER & vbCrLf
    diagOutput = diagOutput & "Excel Version: " & Application.Version & vbCrLf
    diagOutput = diagOutput & "OS:            " & Application.OperatingSystem & vbCrLf
    diagOutput = diagOutput & "Computer:      " & Environ("COMPUTERNAME") & vbCrLf
    diagOutput = diagOutput & "User:          " & Environ("USERNAME") & vbCrLf
    diagOutput = diagOutput & vbCrLf

    ' ========== SECTION 2: Settings ==========
    diagOutput = diagOutput & FormatHeader("SETTINGS")
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Worksheets("Settings")
    On Error GoTo ErrHandler

    If Not wsSettings Is Nothing Then
        diagOutput = diagOutput & ReadSettingLine(wsSettings, "KeepLoginCredentials")
        diagOutput = diagOutput & ReadSettingLine(wsSettings, "SendingMethod")
        diagOutput = diagOutput & ReadSettingLine(wsSettings, "UseDefaultChromeBinary")
        diagOutput = diagOutput & ReadSettingLine(wsSettings, "CustomChromeBinaryPath")
        diagOutput = diagOutput & ReadSettingLine(wsSettings, "DefaultDelay")
        diagOutput = diagOutput & ReadSettingLine(wsSettings, "DelayTimeAttachment")
        diagOutput = diagOutput & ReadSettingLine(wsSettings, "useRandomDelay")
        diagOutput = diagOutput & ReadSettingLine(wsSettings, "minRandomDelay")
        diagOutput = diagOutput & ReadSettingLine(wsSettings, "maxRandomDelay")
        diagOutput = diagOutput & ReadSettingLine(wsSettings, "KeepOriginalFormatting")
    Else
        diagOutput = diagOutput & "  [Settings worksheet not found]" & vbCrLf
    End If
    diagOutput = diagOutput & vbCrLf

    ' ========== SECTION 3: Backend Settings ==========
    diagOutput = diagOutput & FormatHeader("BACKEND SETTINGS")
    On Error Resume Next
    Set wsBackend = ThisWorkbook.Worksheets("Backend_Settings")
    On Error GoTo ErrHandler

    If Not wsBackend Is Nothing Then
        diagOutput = diagOutput & ReadSettingLine(wsBackend, "RETRIEVE_LATEST_XPATHS")
        diagOutput = diagOutput & ReadSettingLine(wsBackend, "LAST_XPATH_RETRIEVED")
        diagOutput = diagOutput & ReadNamedRangeValue("CONNECTION_MODE")
    Else
        diagOutput = diagOutput & "  [Backend_Settings worksheet not found]" & vbCrLf
    End If
    diagOutput = diagOutput & vbCrLf

    ' ========== Refresh XPaths from API (if enabled) ==========
    If Not wsBackend Is Nothing Then
        Dim retrieveXPaths As Boolean
        On Error Resume Next
        retrieveXPaths = CBool(wsBackend.Range("RETRIEVE_LATEST_XPATHS").Value)
        On Error GoTo ErrHandler

        If retrieveXPaths Then
            diagOutput = diagOutput & "  Retrieving latest XPaths from API..." & vbCrLf
            On Error Resume Next
            Dim xpathRefreshed As Boolean
            xpathRefreshed = ParseXPathsFromAPI()
            On Error GoTo ErrHandler
            If xpathRefreshed Then
                diagOutput = diagOutput & "  XPaths refreshed successfully." & vbCrLf
            Else
                diagOutput = diagOutput & "  Could not refresh XPaths (API unavailable). Using cached values." & vbCrLf
            End If
        Else
            diagOutput = diagOutput & "  RETRIEVE_LATEST_XPATHS is disabled. Using cached values." & vbCrLf
        End If
    End If
    diagOutput = diagOutput & vbCrLf

    ' ========== SECTION 4: Current XPath Values ==========
    diagOutput = diagOutput & FormatHeader("CURRENT XPATH VALUES (Primary)")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathSearchInputField")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathTextInputField")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathNoContactFound")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathInvalidPhoneNumber")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathAttachmentButton")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathMultipleAttachmentButton")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathDocumentAttachmentButton")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathMediaAttachmentButton")
    diagOutput = diagOutput & ReadNamedRangeValue("CSSClassModalPopup")
    diagOutput = diagOutput & vbCrLf

    diagOutput = diagOutput & FormatHeader("CURRENT XPATH VALUES (Alternate)")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathSearchInputField_Alt")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathTextInputField_Alt")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathNoContactFound_Alt")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathInvalidPhoneNumber_Alt")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathAttachmentButton_Alt")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathMultipleAttachmentButton_Alt")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathDocumentAttachmentButton_Alt")
    diagOutput = diagOutput & ReadNamedRangeValue("XPathMediaAttachmentButton_Alt")
    diagOutput = diagOutput & ReadNamedRangeValue("CSSClassModalPopup_Alt")
    diagOutput = diagOutput & vbCrLf

    ' ========== SECTION 5: Browser Diagnostics ==========
    On Error GoTo SeleniumError
    Set byObj = CreateObject("Selenium.By")
    On Error GoTo ErrHandler

    Set botObj = InitDiagWebDriver(wsSettings)
    If botObj Is Nothing Then
        diagOutput = diagOutput & FormatHeader("BROWSER DIAGNOSTICS")
        diagOutput = diagOutput & "  [Could not initialize Chrome WebDriver]" & vbCrLf
        GoTo WriteFile
    End If

    On Error GoTo ChromeError
    botObj.Start "chrome"
    On Error GoTo ErrHandler

    ' Browser info
    diagOutput = diagOutput & FormatHeader("BROWSER INFO")
    Dim userAgent As String
    userAgent = botObj.ExecuteScript("return navigator.userAgent;")
    diagOutput = diagOutput & "  User Agent: " & userAgent & vbCrLf

    botObj.Get "https://web.whatsapp.com/"

    ' --- LOGIN PROMPT ---
    Dim loginMsg As String
    loginMsg = "Chrome has opened WhatsApp Web." & vbCrLf & vbCrLf
    loginMsg = loginMsg & "If you are not logged in yet, scan the QR code with your phone." & vbCrLf
    loginMsg = loginMsg & "Wait until you see your chat list." & vbCrLf & vbCrLf
    loginMsg = loginMsg & "DO NOT click or type anything in WhatsApp yet." & vbCrLf
    loginMsg = loginMsg & "Just wait for chats to load, then click OK here."
    userResponse = MsgBox(loginMsg, vbOKCancel + vbInformation, "Diagnostics - Login")

    If userResponse = vbCancel Then GoTo BrowserCleanup

    WaitForPageLoadDiag botObj, 15

    Dim waVersion As String
    waVersion = GetWAVersionDiag(botObj)
    diagOutput = diagOutput & "  WhatsApp Web Version: " & waVersion & vbCrLf
    diagOutput = diagOutput & vbCrLf

    ' ========== PHASE 1: Idle State ==========
    diagOutput = diagOutput & FormatHeader("PHASE 1: IDLE STATE - Discovered Elements")
    diagOutput = diagOutput & DiscoverElements(botObj)
    diagOutput = diagOutput & vbCrLf

    ' Phase 1b: XPath match test in idle state
    diagOutput = diagOutput & FormatHeader("PHASE 1b: IDLE STATE - XPath Match Test")
    diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathSearchInputField")
    diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathSearchInputField_Alt")
    diagOutput = diagOutput & TestCSS(botObj, byObj, "CSSClassModalPopup")
    diagOutput = diagOutput & TestCSS(botObj, byObj, "CSSClassModalPopup_Alt")
    diagOutput = diagOutput & vbCrLf

    ' Track whether search field was found in idle state (the correct test state)
    Dim searchFieldFoundIdle As Boolean
    searchFieldFoundIdle = IsXPathFound(botObj, byObj, "XPathSearchInputField") Or _
                           IsXPathFound(botObj, byObj, "XPathSearchInputField_Alt")

    ' ========== PHASE 2: Search & No-Contact-Found ==========
    ' Inject a click listener to capture the search bar element
    On Error Resume Next
    botObj.ExecuteScript "window._diagClickedEl=null; document.addEventListener('click',function(e){window._diagClickedEl=e.target;});"
    On Error GoTo ErrHandler

    Dim step1Msg As String
    step1Msg = "STEP 1 of 2: Test search and 'no results' detection" & vbCrLf & vbCrLf
    step1Msg = step1Msg & "In WhatsApp Web:" & vbCrLf
    step1Msg = step1Msg & "  1. Click on the SEARCH BAR at the top" & vbCrLf
    step1Msg = step1Msg & "  2. Type exactly:  zzz999" & vbCrLf
    step1Msg = step1Msg & "  3. Wait until you see a 'No results' message" & vbCrLf & vbCrLf
    step1Msg = step1Msg & "Then come back here and click OK."
    userResponse = MsgBox(step1Msg, vbOKCancel + vbInformation, "Diagnostics - Step 1 of 2")

    If userResponse = vbOK Then
        ' Capture what element was clicked (search bar identification)
        diagOutput = diagOutput & FormatHeader("PHASE 2a: SEARCH FIELD IDENTIFICATION")
        On Error Resume Next
        Dim clickedInfo As String
        clickedInfo = botObj.ExecuteScript( _
            "var el=window._diagClickedEl; if(!el)return 'No click captured.'; " & _
            "var o='<'+el.tagName+'> '; " & _
            "if(el.getAttribute('role'))o+='role=""'+el.getAttribute('role')+'""  '; " & _
            "if(el.getAttribute('data-tab'))o+='data-tab=""'+el.getAttribute('data-tab')+'""  '; " & _
            "if(el.getAttribute('contenteditable'))o+='contenteditable=""'+el.getAttribute('contenteditable')+'""  '; " & _
            "if(el.getAttribute('aria-label'))o+='aria-label=""'+el.getAttribute('aria-label')+'""  '; " & _
            "if(el.getAttribute('data-testid'))o+='data-testid=""'+el.getAttribute('data-testid')+'""  '; " & _
            "if(el.type)o+='type=""'+el.type+'""  '; " & _
            "o+='\n  Classes: '+String(el.className).substring(0,200); " & _
            "var p=el.parentElement; if(p)o+='\n  Parent: <'+p.tagName+'> '+String(p.className).substring(0,100); " & _
            "return o;")
        diagOutput = diagOutput & "  Clicked element: " & clickedInfo & vbCrLf & vbCrLf

        ' Capture the currently focused element (should be the search field with zzz999 typed)
        Dim focusedInfo As String
        focusedInfo = botObj.ExecuteScript( _
            "var el=document.activeElement; if(!el)return 'No focused element.'; " & _
            "var o='<'+el.tagName+'> '; " & _
            "if(el.getAttribute('role'))o+='role=""'+el.getAttribute('role')+'""  '; " & _
            "if(el.getAttribute('data-tab'))o+='data-tab=""'+el.getAttribute('data-tab')+'""  '; " & _
            "if(el.getAttribute('contenteditable'))o+='contenteditable=""'+el.getAttribute('contenteditable')+'""  '; " & _
            "if(el.getAttribute('aria-label'))o+='aria-label=""'+el.getAttribute('aria-label')+'""  '; " & _
            "if(el.type)o+='type=""'+el.type+'""  '; " & _
            "o+='\n  Classes: '+String(el.className).substring(0,200); " & _
            "return o;")
        diagOutput = diagOutput & "  Focused element: " & focusedInfo & vbCrLf & vbCrLf

        ' Capture suggested XPath while the search field is focused
        Dim sugSearch As String
        Dim sugJs As String
        sugJs = "var el=document.activeElement; "
        sugJs = sugJs & "if(!el||!(el.tagName==='INPUT'||el.getAttribute('contenteditable')==='true'||el.getAttribute('role')==='textbox'))return ''; "
        sugJs = sugJs & "var tag=el.tagName.toLowerCase(); "
        sugJs = sugJs & "var al=el.getAttribute('aria-label'); "
        sugJs = sugJs & "var dt=el.getAttribute('data-tab'); "
        sugJs = sugJs & "var rl=el.getAttribute('role'); "
        sugJs = sugJs & "if(al)return '//'+tag+'[@aria-label='+JSON.stringify(al)+']'; "
        sugJs = sugJs & "if(dt)return '//'+tag+'[@data-tab='+JSON.stringify(dt)+']'; "
        sugJs = sugJs & "if(rl)return '//'+tag+'[@role='+JSON.stringify(rl)+']'; "
        sugJs = sugJs & "return '//'+tag+'[@contenteditable='+JSON.stringify('true')+']';"
        sugSearch = botObj.ExecuteScript(sugJs)
        If Len(sugSearch) > 0 Then
            sugSearch = vbCrLf & "           SUGGESTED XPATH: " & sugSearch
        End If
        On Error GoTo ErrHandler

        ' Search field XPath was tested in Phase 1b (idle/empty state).
        ' WhatsApp changes the aria-label when text is typed, so we use the idle result.
        If searchFieldFoundIdle Then
            AddOK summaryText, "Search field XPath is working."
        Else
            AddIssue summaryText, issueCount, "SEARCH FIELD: Neither XPathSearchInputField nor its Alt matched in idle state." & vbCrLf & _
                "           The bot will not be able to search for contacts." & sugSearch
        End If

        ' Test no-contact-found XPaths (zzz999 should show "no results")
        diagOutput = diagOutput & FormatHeader("PHASE 2b: NO CONTACT FOUND - XPath Match Test")
        diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathNoContactFound")
        diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathNoContactFound_Alt")
        diagOutput = diagOutput & vbCrLf

        If Not IsXPathFound(botObj, byObj, "XPathNoContactFound") And Not IsXPathFound(botObj, byObj, "XPathNoContactFound_Alt") Then
            AddWarn summaryText, warnCount, "NO CONTACT FOUND: Neither XPathNoContactFound nor its Alt matched." & vbCrLf & _
                "           The bot will still work (cursor-in-footer JS check prevents wrong-chat sends)," & vbCrLf & _
                "           but invalid contacts will take longer to skip and show a generic error."
        Else
            AddOK summaryText, "'No contact found' detection is working."
        End If

        diagOutput = diagOutput & vbCrLf
    End If

    ' ========== PHASE 3: Chat Open ==========
    Dim step2Msg As String
    step2Msg = "STEP 2 of 2: Open a chat" & vbCrLf & vbCrLf
    step2Msg = step2Msg & "  1. Clear the search bar (select all and delete, or click X)" & vbCrLf
    step2Msg = step2Msg & "  2. Click on any contact to open a chat" & vbCrLf
    step2Msg = step2Msg & "  3. Wait until you see the message input area at the bottom" & vbCrLf & vbCrLf
    step2Msg = step2Msg & "Then come back here and click OK."
    userResponse = MsgBox(step2Msg, vbOKCancel + vbInformation, "Diagnostics - Step 2 of 2")

    If userResponse = vbOK Then
        diagOutput = diagOutput & FormatHeader("PHASE 3: CHAT OPEN - Discovered Elements")
        diagOutput = diagOutput & DiscoverChatElements(botObj)
        diagOutput = diagOutput & vbCrLf

        diagOutput = diagOutput & FormatHeader("PHASE 3: CHAT OPEN - XPath Match Test")
        diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathTextInputField")
        diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathTextInputField_Alt")
        diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathAttachmentButton")
        diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathAttachmentButton_Alt")
        diagOutput = diagOutput & vbCrLf

        ' Test the JavaScript cursor-in-footer check (used by saved contact flow)
        Dim footerCheckResult As Variant
        On Error Resume Next
        footerCheckResult = botObj.ExecuteScript("var el = document.activeElement; return el && el.closest('footer') !== null;")
        On Error GoTo ErrHandler
        diagOutput = diagOutput & "  JS Footer Check (cursor in footer): " & CStr(footerCheckResult) & vbCrLf

        ' Also verify the <footer> element exists in the DOM at all
        Dim footerExists As Variant
        On Error Resume Next
        footerExists = botObj.ExecuteScript("return document.querySelector('#main footer') !== null;")
        On Error GoTo ErrHandler
        diagOutput = diagOutput & "  JS Footer Exists (#main footer):    " & CStr(footerExists) & vbCrLf
        diagOutput = diagOutput & vbCrLf

        If footerExists = False Then
            AddIssue summaryText, issueCount, "FOOTER ELEMENT: The <footer> element was not found inside #main." & vbCrLf & _
                "           The cursor-in-footer JS check (used for saved contact verification) will not work." & vbCrLf & _
                "           WhatsApp Web may have changed its page structure."
        Else
            AddOK summaryText, "Footer element exists in #main (JS cursor check will work)."
        End If

        ' Track: text input field
        If Not IsXPathFound(botObj, byObj, "XPathTextInputField") And Not IsXPathFound(botObj, byObj, "XPathTextInputField_Alt") Then
            Dim sugText As String
            sugText = SuggestTextInputXPath(botObj)
            AddIssue summaryText, issueCount, "TEXT INPUT FIELD: Neither XPathTextInputField nor its Alt matched with a chat open." & vbCrLf & _
                "           The bot will not be able to type or send messages." & sugText
        Else
            AddOK summaryText, "Text input field XPath is working."
        End If

        ' Track: attachment button
        If Not IsXPathFound(botObj, byObj, "XPathAttachmentButton") And Not IsXPathFound(botObj, byObj, "XPathAttachmentButton_Alt") Then
            Dim sugAttach As String
            sugAttach = SuggestAttachmentXPath(botObj)
            AddIssue summaryText, issueCount, "ATTACHMENT BUTTON: Neither XPathAttachmentButton nor its Alt matched." & vbCrLf & _
                "           The bot will not be able to send media or documents." & sugAttach
        Else
            AddOK summaryText, "Attachment button XPath is working."
        End If

        ' ========== PHASE 3b: Attachment Dropdown ==========
        ' Click the attachment button to open the dropdown, then capture its contents
        diagOutput = diagOutput & FormatHeader("PHASE 3b: ATTACHMENT DROPDOWN")
        Dim attachClicked As Boolean
        attachClicked = False

        On Error Resume Next
        ' Try primary attachment button
        If botObj.IsElementPresent(byObj.XPath("//span[@data-icon='plus-rounded']")) Then
            botObj.FindElementByXPath("//span[@data-icon='plus-rounded']").Click
            attachClicked = True
        ElseIf botObj.IsElementPresent(byObj.XPath("//span[@data-icon='clip']")) Then
            botObj.FindElementByXPath("//span[@data-icon='clip']").Click
            attachClicked = True
        ElseIf botObj.IsElementPresent(byObj.XPath("//span[@data-icon='plus']")) Then
            botObj.FindElementByXPath("//span[@data-icon='plus']").Click
            attachClicked = True
        End If
        On Error GoTo ErrHandler

        If attachClicked Then
            botObj.Wait 1500

            diagOutput = diagOutput & DiscoverAttachmentDropdown(botObj)
            diagOutput = diagOutput & vbCrLf

            diagOutput = diagOutput & FormatHeader("PHASE 3b: ATTACHMENT DROPDOWN - XPath Match Test")
            diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathMediaAttachmentButton")
            diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathMediaAttachmentButton_Alt")
            diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathDocumentAttachmentButton")
            diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathDocumentAttachmentButton_Alt")
            diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathMultipleAttachmentButton")
            diagOutput = diagOutput & TestXPath(botObj, byObj, "XPathMultipleAttachmentButton_Alt")
            diagOutput = diagOutput & vbCrLf

            ' Track: media button
            If Not IsXPathFound(botObj, byObj, "XPathMediaAttachmentButton") And Not IsXPathFound(botObj, byObj, "XPathMediaAttachmentButton_Alt") Then
                Dim sugMedia As String
                sugMedia = SuggestDropdownItemXPath(botObj, "photo|video|media|image")
                AddIssue summaryText, issueCount, "MEDIA BUTTON: Neither XPathMediaAttachmentButton nor its Alt matched in the dropdown." & vbCrLf & _
                    "           The bot will not be able to send photos/videos." & sugMedia
            Else
                AddOK summaryText, "Media attachment button XPath is working."
            End If

            If Not IsXPathFound(botObj, byObj, "XPathDocumentAttachmentButton") And Not IsXPathFound(botObj, byObj, "XPathDocumentAttachmentButton_Alt") Then
                Dim sugDoc As String
                sugDoc = SuggestDropdownItemXPath(botObj, "document|file")
                AddIssue summaryText, issueCount, "DOCUMENT BUTTON: Neither XPathDocumentAttachmentButton nor its Alt matched in the dropdown." & vbCrLf & _
                    "           The bot will not be able to send documents." & sugDoc
            Else
                AddOK summaryText, "Document attachment button XPath is working."
            End If

            ' Close the dropdown
            On Error Resume Next
            Dim ksObj As Object
            Set ksObj = CreateObject("Selenium.Keys")
            botObj.SendKeys ksObj.Escape
            botObj.Wait 500
            On Error GoTo ErrHandler
        Else
            diagOutput = diagOutput & "  Could not click attachment button (none of the known icons found)." & vbCrLf
            diagOutput = diagOutput & vbCrLf
            AddIssue summaryText, issueCount, "ATTACHMENT DROPDOWN: Could not open the attachment dropdown. None of the known button icons were found."
        End If
    End If

BrowserCleanup:
    On Error Resume Next
    If Not botObj Is Nothing Then botObj.Quit
    Set botObj = Nothing
    On Error GoTo ErrHandler

WriteFile:
    ' ========== SUMMARY ==========
    diagOutput = diagOutput & FormatHeader("DIAGNOSTIC SUMMARY")
    If issueCount = 0 And warnCount = 0 Then
        diagOutput = diagOutput & "  No issues detected. All tested XPaths and JS checks passed." & vbCrLf
        diagOutput = diagOutput & "  If the user is still experiencing problems, the issue may be" & vbCrLf
        diagOutput = diagOutput & "  timing-related rather than selector-related." & vbCrLf
    Else
        diagOutput = diagOutput & "  Issues: " & issueCount & "    Warnings: " & warnCount & vbCrLf
        diagOutput = diagOutput & vbCrLf
        diagOutput = diagOutput & summaryText
        If issueCount > 0 Then
            diagOutput = diagOutput & vbCrLf
            diagOutput = diagOutput & "  NEXT STEPS:" & vbCrLf
            diagOutput = diagOutput & "  1. Review the 'Discovered Elements' sections above to find the correct" & vbCrLf
            diagOutput = diagOutput & "     selectors for this user's WhatsApp Web version." & vbCrLf
            diagOutput = diagOutput & "  2. Look for matching data-icon, data-tab, data-testid, or role attributes." & vbCrLf
            diagOutput = diagOutput & "  3. Update the API XPaths (Version_A or Version_B)." & vbCrLf
        End If
    End If
    diagOutput = diagOutput & vbCrLf

    diagOutput = diagOutput & FormatHeader("END OF DIAGNOSTICS REPORT")

    filePath = GetDiagFilePath()
    If Len(filePath) = 0 Then
        MsgBox "Could not find a writable folder for the diagnostics file." & vbCrLf & _
               "Please try running from a local drive.", vbCritical, "Diagnostics"
        GoTo CleanUp
    End If

    WriteDiagFile filePath, diagOutput

    MsgBox "Diagnostics saved to:" & vbCrLf & vbCrLf & _
           filePath & vbCrLf & vbCrLf & _
           "Please send this file to support.", _
           vbInformation, "Diagnostics Complete"

    GoTo CleanUp

SeleniumError:
    diagOutput = diagOutput & FormatHeader("BROWSER DIAGNOSTICS")
    diagOutput = diagOutput & "  Selenium not installed or not configured." & vbCrLf
    diagOutput = diagOutput & "  Error: " & Err.Description & vbCrLf
    On Error GoTo ErrHandler
    GoTo WriteFile

ChromeError:
    diagOutput = diagOutput & FormatHeader("BROWSER DIAGNOSTICS")
    diagOutput = diagOutput & "  Chrome could not be started." & vbCrLf
    diagOutput = diagOutput & "  Error: " & Err.Description & vbCrLf
    On Error Resume Next
    If Not botObj Is Nothing Then botObj.Quit
    Set botObj = Nothing
    On Error GoTo ErrHandler
    GoTo WriteFile

ErrHandler:
    On Error Resume Next
    If Not botObj Is Nothing Then botObj.Quit
    Set botObj = Nothing
    On Error GoTo 0
    MsgBox "Diagnostics encountered an error:" & vbCrLf & _
           Err.Description, vbCritical, "Diagnostics Error"

CleanUp:
    Set botObj = Nothing
    Set byObj = Nothing
End Sub

' ========================================================================================
'                              WEBDRIVER INITIALIZATION
' ========================================================================================

'---------------------------------------------------------------------------------------
' Function  : InitDiagWebDriver
' Purpose   : Creates a Selenium WebDriver with the same configuration as the bot.
'             Simplified error handling since this is a diagnostic tool.
'---------------------------------------------------------------------------------------
Private Function InitDiagWebDriver(ByVal wsSettings As Worksheet) As Object
    On Error GoTo ErrHandler

    Dim webDriver As Object
    Set webDriver = CreateObject("Selenium.WebDriver")
    webDriver.Timeouts.ImplicitWait = ImplicitWait
    webDriver.Timeouts.PageLoad = PageLoad
    webDriver.Timeouts.Server = TimeoutServer

    webDriver.AddArgument "--disable-popup-blocking"
    webDriver.AddArgument "--disable-notifications"
    webDriver.AddArgument "--lang=en"

    If Not wsSettings Is Nothing Then
        ' Reuse existing Chrome user data folder (so the user stays logged in)
        Dim keepLogin As String
        On Error Resume Next
        keepLogin = wsSettings.Range("KeepLoginCredentials").Value
        On Error GoTo ErrHandler

        If keepLogin = "Yes" Then
            Dim userDataPath As String
            On Error Resume Next
            userDataPath = ThisWorkbook.Names("UserDataFolderPath").RefersToRange.Value
            On Error GoTo ErrHandler
            If Len(userDataPath) > 0 Then
                webDriver.AddArgument "--user-data-dir=" & userDataPath
            End If
        End If

        ' Custom Chrome binary
        Dim useDefault As String
        On Error Resume Next
        useDefault = LCase(Trim(wsSettings.Range("UseDefaultChromeBinary").Value))
        On Error GoTo ErrHandler

        If useDefault = "no" Then
            Dim customPath As String
            On Error Resume Next
            customPath = Trim(Replace(wsSettings.Range("CustomChromeBinaryPath").Value, """", ""))
            On Error GoTo ErrHandler
            If Len(customPath) > 0 And Dir(customPath) <> "" Then
                webDriver.SetBinary customPath
            End If
        End If
    End If

    Set InitDiagWebDriver = webDriver
    Exit Function

ErrHandler:
    Set InitDiagWebDriver = Nothing
End Function

' ========================================================================================
'                              DOM ELEMENT DISCOVERY (JavaScript)
' ========================================================================================

'---------------------------------------------------------------------------------------
' Function  : DiscoverElements
' Purpose   : Discovers UI elements in the idle state using universal attributes:
'             contenteditable, data-icon, data-testid, role=textbox.
'---------------------------------------------------------------------------------------
Private Function DiscoverElements(botObj As Object) As String
    On Error Resume Next
    Dim js As String

    ' --- Contenteditable elements (search field, message field candidates) ---
    js = "var o=''; var els=document.querySelectorAll('[contenteditable=""true""]'); "
    js = js & "o+='Contenteditable elements found: '+els.length+'\n'; "
    js = js & "for(var i=0;i<els.length;i++){ var e=els[i]; "
    js = js & "o+='  '+(i+1)+'. <'+e.tagName+'> '; "
    js = js & "if(e.getAttribute('role'))o+='role=""'+e.getAttribute('role')+'""  '; "
    js = js & "if(e.getAttribute('data-tab'))o+='data-tab=""'+e.getAttribute('data-tab')+'""  '; "
    js = js & "if(e.getAttribute('aria-label'))o+='aria-label=""'+e.getAttribute('aria-label')+'""  '; "
    js = js & "if(e.getAttribute('title'))o+='title=""'+e.getAttribute('title')+'""  '; "
    js = js & "if(e.getAttribute('data-testid'))o+='data-testid=""'+e.getAttribute('data-testid')+'""  '; "
    js = js & "o+='\n     Classes: '+String(e.className).substring(0,200)+'\n'; "
    js = js & "var p=e.parentElement; "
    js = js & "if(p)o+='     Parent: <'+p.tagName+'> '+String(p.className).substring(0,150)+'\n'; "
    js = js & "o+='\n'; } return o;"
    DiscoverElements = botObj.ExecuteScript(js) & vbCrLf

    ' --- Data-icon elements (attachment, send, emoji buttons etc.) ---
    js = "var o=''; var els=document.querySelectorAll('[data-icon]'); "
    js = js & "o+='Data-icon elements found: '+els.length+'\n'; "
    js = js & "for(var i=0;i<els.length;i++){ var e=els[i]; "
    js = js & "o+='  '+(i+1)+'. <'+e.tagName+'> data-icon=""'+e.getAttribute('data-icon')+'""'; "
    js = js & "var b=e.closest('button,[role=""button""]'); "
    js = js & "if(b&&b!==e)o+=' in <'+b.tagName+'> '+String(b.className).substring(0,80); "
    js = js & "if(e.getAttribute('data-testid'))o+=' data-testid=""'+e.getAttribute('data-testid')+'""'; "
    js = js & "o+='\n'; } return o;"
    DiscoverElements = DiscoverElements & botObj.ExecuteScript(js) & vbCrLf

    ' --- Unique data-testid values (most stable selectors) ---
    js = "var o=''; var els=document.querySelectorAll('[data-testid]'); var seen={}; "
    js = js & "for(var i=0;i<els.length;i++){ var tid=els[i].getAttribute('data-testid'); "
    js = js & "if(!seen[tid])seen[tid]=0; seen[tid]++; } "
    js = js & "var keys=Object.keys(seen).sort(); "
    js = js & "o+='Unique data-testid values: '+keys.length+'\n'; "
    js = js & "for(var i=0;i<keys.length;i++){ o+='  '+keys[i]+' (x'+seen[keys[i]]+')\n'; } "
    js = js & "return o;"
    DiscoverElements = DiscoverElements & botObj.ExecuteScript(js)

    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Function  : DiscoverSearchResults
' Purpose   : Captures the search results panel structure after the user has typed
'             a contact name in the search bar.
'---------------------------------------------------------------------------------------
Private Function DiscoverSearchResults(botObj As Object) As String
    On Error Resume Next
    Dim js As String

    ' Search result list items
    js = "var o=''; var side=document.getElementById('side'); "
    js = js & "if(!side)return 'Side panel (#side) not found.\n'; "
    js = js & "var items=side.querySelectorAll('[role=""listitem""]'); "
    js = js & "o+='Search result items (role=listitem): '+items.length+'\n\n'; "
    js = js & "for(var i=0;i<Math.min(items.length,5);i++){ var e=items[i]; "
    js = js & "o+='  Result '+(i+1)+':\n'; "
    js = js & "o+='    Tag: <'+e.tagName+'>\n'; "
    js = js & "o+='    Classes: '+String(e.className).substring(0,200)+'\n'; "
    js = js & "if(e.getAttribute('data-testid'))o+='    data-testid: '+e.getAttribute('data-testid')+'\n'; "
    js = js & "var spans=e.querySelectorAll('span[title]'); "
    js = js & "for(var j=0;j<spans.length;j++){ "
    js = js & "o+='    Span title: ""'+spans[j].getAttribute('title')+'""  \n'; } "
    js = js & "o+='    Inner HTML (first 500 chars):\n'; "
    js = js & "o+='      '+e.innerHTML.substring(0,500).replace(/\n/g,' ')+'\n\n'; } "

    ' Contenteditable elements in the side panel during search
    js = js & "var ce=side.querySelectorAll('[contenteditable=""true""]'); "
    js = js & "o+='\nContenteditable in side panel: '+ce.length+'\n'; "
    js = js & "for(var i=0;i<ce.length;i++){ var el=ce[i]; "
    js = js & "o+='  '+(i+1)+'. <'+el.tagName+'> '; "
    js = js & "if(el.getAttribute('data-tab'))o+='data-tab=""'+el.getAttribute('data-tab')+'""  '; "
    js = js & "if(el.getAttribute('role'))o+='role=""'+el.getAttribute('role')+'""  '; "
    js = js & "if(el.getAttribute('aria-label'))o+='aria-label=""'+el.getAttribute('aria-label')+'""  '; "
    js = js & "o+='\n     Text content: ""'+el.textContent.substring(0,100)+'""  \n'; } "
    js = js & "return o;"
    DiscoverSearchResults = botObj.ExecuteScript(js)

    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Function  : DiscoverChatElements
' Purpose   : Captures the chat area structure (header, footer/message input,
'             attachment buttons) after the user has opened a chat.
'---------------------------------------------------------------------------------------
Private Function DiscoverChatElements(botObj As Object) As String
    On Error Resume Next
    Dim js As String

    ' Chat header
    js = "var o=''; var main=document.getElementById('main'); "
    js = js & "if(!main)return 'Main panel (#main) not found.\n'; "
    js = js & "o+='--- Chat Header ---\n'; var header=main.querySelector('header'); "
    js = js & "if(header){ var spans=header.querySelectorAll('span[title],span[dir]'); "
    js = js & "o+='  Header spans: '+spans.length+'\n'; "
    js = js & "for(var i=0;i<Math.min(spans.length,10);i++){ o+='    <SPAN> '; "
    js = js & "if(spans[i].getAttribute('title'))o+='title=""'+spans[i].getAttribute('title')+'""  '; "
    js = js & "if(spans[i].getAttribute('dir'))o+='dir=""'+spans[i].getAttribute('dir')+'""  '; "
    js = js & "if(spans[i].getAttribute('data-testid'))o+='data-testid=""'+spans[i].getAttribute('data-testid')+'""  '; "
    js = js & "o+='text=""'+spans[i].textContent.substring(0,50)+'""  \n'; } "
    js = js & "}else{o+='  No <header> found in #main\n';} "

    ' Footer / message input
    js = js & "o+='\n--- Footer / Message Input ---\n'; "
    js = js & "var footer=main.querySelector('footer'); "
    js = js & "if(footer){ var ce=footer.querySelectorAll('[contenteditable=""true""]'); "
    js = js & "o+='  Contenteditable in footer: '+ce.length+'\n'; "
    js = js & "for(var i=0;i<ce.length;i++){ var e=ce[i]; o+='    '+(i+1)+'. <'+e.tagName+'> '; "
    js = js & "if(e.getAttribute('role'))o+='role=""'+e.getAttribute('role')+'""  '; "
    js = js & "if(e.getAttribute('data-tab'))o+='data-tab=""'+e.getAttribute('data-tab')+'""  '; "
    js = js & "if(e.getAttribute('aria-label'))o+='aria-label=""'+e.getAttribute('aria-label')+'""  '; "
    js = js & "if(e.getAttribute('data-testid'))o+='data-testid=""'+e.getAttribute('data-testid')+'""  '; "
    js = js & "o+='\n       Classes: '+String(e.className).substring(0,200)+'\n'; } "
    js = js & "var icons=footer.querySelectorAll('[data-icon]'); "
    js = js & "o+='\n  Data-icon in footer: '+icons.length+'\n'; "
    js = js & "for(var i=0;i<icons.length;i++){ "
    js = js & "o+='    '+(i+1)+'. data-icon=""'+icons[i].getAttribute('data-icon')+'""'; "
    js = js & "if(icons[i].getAttribute('data-testid'))o+=' data-testid=""'+icons[i].getAttribute('data-testid')+'""'; "
    js = js & "o+='\n'; } "
    js = js & "}else{o+='  No <footer> found in #main\n';} "

    ' Attachment area
    js = js & "o+='\n--- Attachment Area ---\n'; "
    js = js & "var att=main.querySelectorAll('[data-icon*=""attach""],[data-icon*=""plus""],[data-icon=""clip""]'); "
    js = js & "o+='  Attachment-related icons: '+att.length+'\n'; "
    js = js & "for(var i=0;i<att.length;i++){ var e=att[i]; "
    js = js & "o+='    '+(i+1)+'. <'+e.tagName+'> data-icon=""'+e.getAttribute('data-icon')+'""'; "
    js = js & "var btn=e.closest('button,[role=""button""]'); "
    js = js & "if(btn)o+=' in <'+btn.tagName+'> class=""'+String(btn.className).substring(0,80)+'""'; "
    js = js & "o+='\n'; } return o;"
    DiscoverChatElements = botObj.ExecuteScript(js)

    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Function  : DiscoverAttachmentDropdown
' Purpose   : Captures the attachment dropdown menu elements after the plus/clip
'             button has been clicked. Looks for menu items, spans with text,
'             and data-icon elements inside the dropdown.
'---------------------------------------------------------------------------------------
Private Function DiscoverAttachmentDropdown(botObj As Object) As String
    On Error Resume Next
    Dim js As String

    ' Discover only menuitem elements (the actual dropdown options)
    ' Using role="menuitem" avoids capturing chat content from LI/button elements
    js = "var o=''; "
    js = js & "var items=document.querySelectorAll('[role=""menuitem""]'); "
    js = js & "var visible=[]; "
    js = js & "for(var i=0;i<items.length;i++){ "
    js = js & "var r=items[i].getBoundingClientRect(); "
    js = js & "if(r.width>0&&r.height>0)visible.push(items[i]); } "
    js = js & "o+='Dropdown menu items (role=menuitem): '+visible.length+'\n'; "
    js = js & "for(var i=0;i<visible.length;i++){ var e=visible[i]; "
    js = js & "var txt=e.textContent.trim().substring(0,80); "
    js = js & "o+='  '+(i+1)+'. <'+e.tagName+'> '; "
    js = js & "if(e.getAttribute('data-testid'))o+='data-testid=""'+e.getAttribute('data-testid')+'""  '; "
    js = js & "o+='text=""'+txt+'""  \n'; "
    js = js & "var icons=e.querySelectorAll('[data-icon]'); "
    js = js & "for(var j=0;j<icons.length;j++){ "
    js = js & "o+='     data-icon=""'+icons[j].getAttribute('data-icon')+'""  \n'; } "
    js = js & "var spans=e.querySelectorAll('span'); "
    js = js & "for(var j=0;j<spans.length;j++){ var st=spans[j].textContent.trim(); "
    js = js & "if(st.length>0&&st.length<60)o+='     <SPAN> text=""'+st+'""  \n'; } "
    js = js & "o+='\n'; } "

    ' Also capture any recently appeared overlay/popup containers
    js = js & "o+='--- Popup/Overlay Containers ---\n'; "
    js = js & "var popups=document.querySelectorAll('[data-animate-dropdown-item=""true""],"
    js = js & " [class*=""popup""], [class*=""dropdown""], [class*=""menu""]'); "
    js = js & "o+='  Found: '+popups.length+'\n'; "
    js = js & "for(var i=0;i<Math.min(popups.length,5);i++){ var e=popups[i]; "
    js = js & "o+='  '+(i+1)+'. <'+e.tagName+'> class=""'+String(e.className).substring(0,120)+'""  \n'; "
    js = js & "o+='     innerHTML (first 300 chars): '+e.innerHTML.substring(0,300).replace(/\\n/g,' ')+'\n\n'; } "
    js = js & "return o;"
    DiscoverAttachmentDropdown = botObj.ExecuteScript(js)

    On Error GoTo 0
End Function

' ========================================================================================
'                              XPATH / CSS MATCH TESTING
' ========================================================================================

'---------------------------------------------------------------------------------------
' Function  : TestXPath
' Purpose   : Tests whether a configured XPath (from a named range) matches any element
'             on the current page. Returns a formatted result line showing FOUND/NOT FOUND
'             and element details if found.
'---------------------------------------------------------------------------------------
Private Function TestXPath(botObj As Object, byObj As Object, ByVal rangeName As String) As String
    On Error Resume Next

    Dim xpathValue As String
    xpathValue = ""
    xpathValue = Trim(CStr(ThisWorkbook.Names(rangeName).RefersToRange.Value))

    If Len(xpathValue) = 0 Then
        TestXPath = "  " & PadRight(rangeName, 40) & "[not configured]" & vbCrLf
        Exit Function
    End If

    Dim found As Boolean
    found = botObj.IsElementPresent(byObj.XPath(xpathValue))

    If found Then
        Dim el As Object
        Dim tagName As String, className As String, testId As String
        Set el = botObj.FindElementByXPath(xpathValue)
        If Not el Is Nothing Then
            tagName = el.tagName
            className = Left(el.Attribute("class"), 120)
            testId = el.Attribute("data-testid")
        End If
        TestXPath = "  " & PadRight(rangeName, 40) & "FOUND" & vbCrLf & _
                    "    XPath: " & xpathValue & vbCrLf & _
                    "    Matched: <" & tagName & "> class=""" & className & """" & _
                    IIf(Len(testId) > 0, " data-testid=""" & testId & """", "") & vbCrLf
    Else
        TestXPath = "  " & PadRight(rangeName, 40) & "NOT FOUND" & vbCrLf & _
                    "    XPath: " & xpathValue & vbCrLf
    End If

    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Function  : TestCSS
' Purpose   : Same as TestXPath but for CSS selectors.
'---------------------------------------------------------------------------------------
Private Function TestCSS(botObj As Object, byObj As Object, ByVal rangeName As String) As String
    On Error Resume Next

    Dim cssValue As String
    cssValue = ""
    cssValue = Trim(CStr(ThisWorkbook.Names(rangeName).RefersToRange.Value))

    If Len(cssValue) = 0 Then
        TestCSS = "  " & PadRight(rangeName, 40) & "[not configured]" & vbCrLf
        Exit Function
    End If

    Dim found As Boolean
    found = botObj.IsElementPresent(byObj.css(cssValue))

    If found Then
        Dim el As Object
        Dim tagName As String
        Set el = botObj.FindElementByCss(cssValue)
        If Not el Is Nothing Then tagName = el.tagName
        TestCSS = "  " & PadRight(rangeName, 40) & "FOUND" & vbCrLf & _
                  "    CSS: " & cssValue & vbCrLf & _
                  "    Matched: <" & tagName & ">" & vbCrLf
    Else
        TestCSS = "  " & PadRight(rangeName, 40) & "NOT FOUND" & vbCrLf & _
                  "    CSS: " & cssValue & vbCrLf
    End If

    On Error GoTo 0
End Function

' ========================================================================================
'                              SETTINGS READERS
' ========================================================================================

Private Function ReadSettingLine(ws As Worksheet, ByVal rangeName As String) As String
    On Error Resume Next
    Dim val As String
    val = CStr(ws.Range(rangeName).Value)
    If Err.Number <> 0 Then val = "[not found]"
    On Error GoTo 0
    ReadSettingLine = "  " & PadRight(rangeName, 30) & val & vbCrLf
End Function

Private Function ReadNamedRangeValue(ByVal rangeName As String) As String
    On Error Resume Next
    Dim val As String
    val = CStr(ThisWorkbook.Names(rangeName).RefersToRange.Value)
    If Err.Number <> 0 Then val = "[not defined]"
    On Error GoTo 0
    ReadNamedRangeValue = "  " & PadRight(rangeName, 40) & val & vbCrLf
End Function

' ========================================================================================
'                              FILE OUTPUT
' ========================================================================================

'---------------------------------------------------------------------------------------
' Function  : GetDiagFilePath
' Purpose   : Determines a writable path for the diagnostics file.
'             Tries: workbook folder (with OneDrive handling) > Desktop > Temp.
'---------------------------------------------------------------------------------------
Private Function GetDiagFilePath() As String
    On Error Resume Next

    Dim fileName As String
    fileName = DIAG_FILE_PREFIX & Format(Now, "yyyy-mm-dd_hhmmss") & ".txt"

    ' Option 1: Workbook folder (handle OneDrive paths)
    Dim currentDir As String
    currentDir = ThisWorkbook.path
    If Len(currentDir) > 0 Then
        Dim localDir As String
        localDir = GetLocalPath( _
            fullPath:=DecodeURL(currentDir), _
            rebuildCache:=True, _
            returnInputOnFail:=True)
        If Len(localDir) > 0 And IsFolderEditable(localDir) Then
            GetDiagFilePath = BuildPath(localDir, fileName)
            Exit Function
        End If
    End If

    ' Option 2: Desktop
    Dim desktopPath As String
    desktopPath = Environ("USERPROFILE")
    If Len(desktopPath) > 0 Then
        desktopPath = BuildPath(desktopPath, "Desktop")
        If IsFolderEditable(desktopPath) Then
            GetDiagFilePath = BuildPath(desktopPath, fileName)
            Exit Function
        End If
    End If

    ' Option 3: Temp folder
    Dim tempPath As String
    tempPath = Environ("TEMP")
    If Len(tempPath) > 0 And IsFolderEditable(tempPath) Then
        GetDiagFilePath = BuildPath(tempPath, fileName)
        Exit Function
    End If

    GetDiagFilePath = ""
    On Error GoTo 0
End Function

Private Sub WriteDiagFile(ByVal filePath As String, ByVal content As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
End Sub

' ========================================================================================
'                              FORMATTING HELPERS
' ========================================================================================

Private Function FormatHeader(ByVal title As String) As String
    FormatHeader = SEPARATOR & vbCrLf & _
                   title & vbCrLf & _
                   SEPARATOR & vbCrLf
End Function

Private Function PadRight(ByVal text As String, ByVal totalWidth As Long) As String
    If Len(text) >= totalWidth Then
        PadRight = text & " "
    Else
        PadRight = text & Space(totalWidth - Len(text))
    End If
End Function

' ========================================================================================
'                              SUMMARY HELPERS
' ========================================================================================

Private Function IsXPathFound(botObj As Object, byObj As Object, ByVal rangeName As String) As Boolean
    On Error Resume Next
    Dim xpathValue As String
    xpathValue = Trim(CStr(ThisWorkbook.Names(rangeName).RefersToRange.Value))
    If Len(xpathValue) = 0 Then
        IsXPathFound = False
        Exit Function
    End If
    IsXPathFound = botObj.IsElementPresent(byObj.XPath(xpathValue))
    On Error GoTo 0
End Function

Private Sub AddIssue(ByRef summaryText As String, ByRef issueCount As Long, ByVal msg As String)
    issueCount = issueCount + 1
    summaryText = summaryText & "  [ISSUE] " & msg & vbCrLf & vbCrLf
End Sub

Private Sub AddWarn(ByRef summaryText As String, ByRef warnCount As Long, ByVal msg As String)
    warnCount = warnCount + 1
    summaryText = summaryText & "  [WARN]  " & msg & vbCrLf & vbCrLf
End Sub

Private Sub AddOK(ByRef summaryText As String, ByVal msg As String)
    summaryText = summaryText & "  [OK]    " & msg & vbCrLf
End Sub

' ========================================================================================
'                              XPATH SUGGESTION HELPERS
' ========================================================================================

'---------------------------------------------------------------------------------------
' Each Suggest* function runs a targeted JS query to find the likely element
' and returns a formatted suggestion string (or empty if nothing found).
'---------------------------------------------------------------------------------------

Private Function SuggestSearchFieldXPath(botObj As Object) As String
    ' Note: This is a fallback. The primary suggestion is captured inline in Phase 2
    ' using document.activeElement while the search field is focused.
    On Error Resume Next
    Dim js As String, result As String

    js = "var side=document.getElementById('side'); if(!side)return ''; "
    js = js & "var el=side.querySelector('input[role=""textbox""]'); "
    js = js & "if(el){ if(el.getAttribute('data-tab'))return '//input[@data-tab='+JSON.stringify(el.getAttribute('data-tab'))+']'; "
    js = js & "return '//input[@role=""textbox""]'; } "
    js = js & "var el2=side.querySelector('[contenteditable=""true""]'); "
    js = js & "if(el2)return '//'+el2.tagName.toLowerCase()+'[@contenteditable=""true""]'; "
    js = js & "return '';"
    result = botObj.ExecuteScript(js)
    If Len(result) > 0 Then
        SuggestSearchFieldXPath = vbCrLf & "           SUGGESTED XPATH: " & result
    End If
    On Error GoTo 0
End Function

Private Function SuggestTextInputXPath(botObj As Object) As String
    On Error Resume Next
    Dim js As String, result As String

    js = "var main=document.getElementById('main'); if(!main)return ''; "
    js = js & "var footer=main.querySelector('footer'); if(!footer)return ''; "
    js = js & "var el=footer.querySelector('[contenteditable=""true""]'); if(!el)return ''; "
    js = js & "if(el.getAttribute('role')==='textbox')return '//*[@id=""main""]//footer//div[@role=""textbox""]'; "
    js = js & "if(el.getAttribute('data-tab'))return '//*[@id=""main""]//footer//'+el.tagName.toLowerCase()+'[@data-tab='+JSON.stringify(el.getAttribute('data-tab'))+']'; "
    js = js & "return '//*[@id=""main""]//footer//'+el.tagName.toLowerCase()+'[@contenteditable=""true""]';"
    result = botObj.ExecuteScript(js)
    If Len(result) > 0 Then
        SuggestTextInputXPath = vbCrLf & "           SUGGESTED XPATH: " & result
    End If
    On Error GoTo 0
End Function

Private Function SuggestAttachmentXPath(botObj As Object) As String
    On Error Resume Next
    Dim js As String, result As String

    js = "var icons=['plus-rounded','clip','plus','attach','attachment']; "
    js = js & "for(var i=0;i<icons.length;i++){ "
    js = js & "var el=document.querySelector('[data-icon='+JSON.stringify(icons[i])+']'); "
    js = js & "if(el&&el.getBoundingClientRect().width>0)return '//span[@data-icon='+JSON.stringify(icons[i])+']'; } "
    js = js & "return '';"
    result = botObj.ExecuteScript(js)
    If Len(result) > 0 Then
        SuggestAttachmentXPath = vbCrLf & "           SUGGESTED XPATH: " & result
    End If
    On Error GoTo 0
End Function

Private Function SuggestDropdownItemXPath(botObj As Object, ByVal keywords As String) As String
    On Error Resume Next
    Dim js As String, result As String

    js = "var kw='" & keywords & "'.split('|'); "
    js = js & "var all=document.querySelectorAll('[role=""menuitem""]'); "
    js = js & "for(var i=0;i<all.length;i++){ var el=all[i]; "
    js = js & "var r=el.getBoundingClientRect(); if(r.width==0)continue; "
    js = js & "var txt=el.textContent.trim().toLowerCase(); "
    js = js & "for(var j=0;j<kw.length;j++){ if(txt.indexOf(kw[j])>=0){ "
    js = js & "var spans=el.querySelectorAll('span'); "
    js = js & "for(var k=0;k<spans.length;k++){ var st=spans[k].textContent.trim(); "
    js = js & "if(st.length>2&&st.length<50)return '//span[contains(text(),'+JSON.stringify(st)+')]'; } "
    js = js & "return '(found text: '+JSON.stringify(txt.substring(0,50))+' - build XPath from this)'; }}} "
    js = js & "return '';"
    result = botObj.ExecuteScript(js)
    If Len(result) > 0 Then
        SuggestDropdownItemXPath = vbCrLf & "           SUGGESTED XPATH: " & result
    End If
    On Error GoTo 0
End Function

' ========================================================================================
'                              BROWSER HELPERS
' ========================================================================================

Private Sub WaitForPageLoadDiag(botObj As Object, ByVal maxWaitSec As Long)
    On Error Resume Next
    Dim startTime As Single
    startTime = Timer
    Do While Timer - startTime < maxWaitSec
        If botObj.ExecuteScript("return document.readyState") = "complete" Then Exit Do
        botObj.Wait 500
    Loop
    On Error GoTo 0
End Sub

Private Function GetWAVersionDiag(botObj As Object) As String
    On Error Resume Next
    Dim ver As String
    ver = botObj.ExecuteScript( _
        "try { " & _
        "  if(window.Debug&&window.Debug.VERSION)return window.Debug.VERSION; " & _
        "  var el=document.querySelector('[data-app-version]'); " & _
        "  if(el)return el.getAttribute('data-app-version'); " & _
        "  return 'Unknown'; " & _
        "} catch(e){return 'Unknown';}")
    If Err.Number <> 0 Or Len(ver) = 0 Then ver = "Unknown"
    GetWAVersionDiag = ver
    On Error GoTo 0
End Function




