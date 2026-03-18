Attribute VB_Name = "mWhatsAppBOT"
'---------------------------------------------------------------------------------------
' Module     : mWhatsAppBOT
' Author     : Sven Bosau
' Website    : https://pythonandvba.com
' Email      : sven@pythonandvba.com
'---------------------------------------------------------------------------------------

Option Explicit

Public BOT As Object, By As Object, ks As Object
Public wb As Workbook
Public ws As Worksheet
Public wsSettings As Worksheet
Public i As Long

' XPath variables (primary)
Public xPathInvalidPhoneNumber As String
Public xPathTextInputField As String
Public xPathSearchInputField As String
Public xPathNoContactFound As String
Public xPathAttachmentButton As String
Public xPathMultipleAttachmentButton As String
Public CSSClassModalPopup As String

' XPath variables (alternate for A/B fallback)
Public xPathInvalidPhoneNumber_Alt As String
Public xPathTextInputField_Alt As String
Public xPathSearchInputField_Alt As String
Public xPathNoContactFound_Alt As String
Public xPathAttachmentButton_Alt As String
Public xPathMultipleAttachmentButton_Alt As String
Public CSSClassModalPopup_Alt As String

' XPath for clicking the first search result (saved contacts)
Public xPathFirstSearchResult As String
Public xPathFirstSearchResult_Alt As String

Private m_WAVersion As String

' ========================================================================================
'                         HELPER FUNCTIONS FOR ROBUSTNESS
' ========================================================================================

'---------------------------------------------------------------------------------------
' Function  : WaitForElement
' Purpose   : Polls for an element using XPath, retrying every pollMs until maxWaitMs.
'             Returns True if found, False if timed out.
'---------------------------------------------------------------------------------------
Private Function WaitForElement(ByVal xp As String, Optional ByVal maxWaitMs As Long = 10000, Optional ByVal pollMs As Long = 300) As Boolean
    Dim elapsed As Long
    elapsed = 0
    Do While elapsed < maxWaitMs
        If BOT.IsElementPresent(By.XPath(xp)) Then
            WaitForElement = True
            Exit Function
        End If
        BOT.Wait pollMs
        elapsed = elapsed + pollMs
    Loop
    WaitForElement = False
End Function

'---------------------------------------------------------------------------------------
' Function  : WaitForElementCss
' Purpose   : Same as WaitForElement but uses CSS selector.
'---------------------------------------------------------------------------------------
Private Function WaitForElementCss(ByVal css As String, Optional ByVal maxWaitMs As Long = 10000, Optional ByVal pollMs As Long = 300) As Boolean
    Dim elapsed As Long
    elapsed = 0
    Do While elapsed < maxWaitMs
        If BOT.IsElementPresent(By.css(css)) Then
            WaitForElementCss = True
            Exit Function
        End If
        BOT.Wait pollMs
        elapsed = elapsed + pollMs
    Loop
    WaitForElementCss = False
End Function

'---------------------------------------------------------------------------------------
' Function  : WaitForElementGone
' Purpose   : Waits until an element is no longer present (e.g. loading spinner disappears).
'---------------------------------------------------------------------------------------
Private Function WaitForElementGone(ByVal xp As String, Optional ByVal maxWaitMs As Long = 10000, Optional ByVal pollMs As Long = 300) As Boolean
    Dim elapsed As Long
    elapsed = 0
    Do While elapsed < maxWaitMs
        If Not BOT.IsElementPresent(By.XPath(xp)) Then
            WaitForElementGone = True
            Exit Function
        End If
        BOT.Wait pollMs
        elapsed = elapsed + pollMs
    Loop
    WaitForElementGone = False
End Function

'---------------------------------------------------------------------------------------
' Function  : FindElementWithFallback
' Purpose   : Tries to find an element using the primary XPath. If not found, tries
'             the alternate XPath. Returns the element or Nothing.
'---------------------------------------------------------------------------------------
Private Function FindElementWithFallback(ByVal xpPrimary As String, ByVal xpAlt As String, Optional ByVal maxWaitMs As Long = 5000) As Object
    ' Try primary first
    If WaitForElement(xpPrimary, maxWaitMs, 300) Then
        On Error Resume Next
        Set FindElementWithFallback = BOT.FindElementByXPath(xpPrimary)
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If
    ' Try alternate
    If Len(xpAlt) > 0 Then
        If WaitForElement(xpAlt, maxWaitMs, 300) Then
            On Error Resume Next
            Set FindElementWithFallback = BOT.FindElementByXPath(xpAlt)
            If Err.Number = 0 Then
                On Error GoTo 0
                Exit Function
            End If
            Err.Clear
            On Error GoTo 0
        End If
    End If
    ' Neither found
    Set FindElementWithFallback = Nothing
End Function

'---------------------------------------------------------------------------------------
' Function  : IsElementPresentWithFallback
' Purpose   : Returns True if element is present using primary or alternate XPath.
'---------------------------------------------------------------------------------------
Private Function IsElementPresentWithFallback(ByVal xpPrimary As String, ByVal xpAlt As String) As Boolean
    If BOT.IsElementPresent(By.XPath(xpPrimary)) Then
        IsElementPresentWithFallback = True
    ElseIf Len(xpAlt) > 0 Then
        If BOT.IsElementPresent(By.XPath(xpAlt)) Then
            IsElementPresentWithFallback = True
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : IsCssElementPresentWithFallback
' Purpose   : Returns True if element is present using primary or alternate CSS.
'---------------------------------------------------------------------------------------
Private Function IsCssElementPresentWithFallback(ByVal cssPrimary As String, ByVal cssAlt As String) As Boolean
    If BOT.IsElementPresent(By.css(cssPrimary)) Then
        IsCssElementPresentWithFallback = True
    ElseIf Len(cssAlt) > 0 Then
        If BOT.IsElementPresent(By.css(cssAlt)) Then
            IsCssElementPresentWithFallback = True
        End If
    End If
End Function

'---------------------------------------------------------------------------------------
' Sub       : ClearSearchField
' Purpose   : Robustly clears the search input field with verification.
'---------------------------------------------------------------------------------------
Private Sub ClearSearchField()
    Dim clearAttempts As Long
    Dim el As Object
    
    For clearAttempts = 1 To 5
        Set el = FindElementWithFallback(xPathSearchInputField, xPathSearchInputField_Alt, 3000)
        If el Is Nothing Then Exit Sub
        
        el.SendKeys ks.Control & "a"
        BOT.Wait 100
        el.SendKeys ks.Delete
        BOT.Wait 300
        
        ' Verify field is empty by checking its text
        If Len(Trim(el.Attribute("textContent"))) = 0 Then
            Exit Sub
        End If
    Next clearAttempts
End Sub

'---------------------------------------------------------------------------------------
' Function  : VerifyChatOpen
' Purpose   : Verifies that a chat is actually open by checking for the text input field.
'             Returns True if the text input field is present (chat is open and ready).
'---------------------------------------------------------------------------------------
Private Function VerifyChatOpen(Optional ByVal maxWaitMs As Long = 8000) As Boolean
    ' Check primary text input field
    If WaitForElement(xPathTextInputField, maxWaitMs, 300) Then
        VerifyChatOpen = True
        Exit Function
    End If
    ' Check alternate text input field
    If Len(xPathTextInputField_Alt) > 0 Then
        If WaitForElement(xPathTextInputField_Alt, 3000, 300) Then
            VerifyChatOpen = True
            Exit Function
        End If
    End If
    VerifyChatOpen = False
End Function

'---------------------------------------------------------------------------------------
' Function  : GetWhatsAppVersion
' Purpose   : Extracts the WhatsApp Web version string from the page via JavaScript.
'             Returns "Unknown" if it cannot be determined.
'---------------------------------------------------------------------------------------
Private Function GetWhatsAppVersion() As String
    On Error Resume Next
    Dim ver As String
    ver = BOT.ExecuteScript( _
        "try { " & _
        "  if (window.Debug && window.Debug.VERSION) return window.Debug.VERSION; " & _
        "  if (typeof __x_config !== 'undefined' && __x_config.version) return __x_config.version; " & _
        "  var el = document.querySelector('[data-app-version]'); " & _
        "  if (el) return el.getAttribute('data-app-version'); " & _
        "  var scripts = document.querySelectorAll('script[src]'); " & _
        "  for (var i = 0; i < scripts.length; i++) { " & _
        "    var m = scripts[i].src.match(/(\d+\.\d{4,}\.\d+)/); " & _
        "    if (m) return m[1]; " & _
        "  } " & _
        "  return 'Unknown'; " & _
        "} catch(e) { return 'Unknown'; }")
    If Err.Number <> 0 Or Len(ver) = 0 Then
        GetWhatsAppVersion = "Unknown"
    Else
        GetWhatsAppVersion = ver
    End If
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Function  : SafeLoadNamedRange
' Purpose   : Safely loads a named range value from a worksheet. Returns "" if the
'             named range does not exist. This avoids errors when alternate XPaths
'             are not yet defined in Backend_Settings.
'---------------------------------------------------------------------------------------
Private Function SafeLoadNamedRange(ByVal wsSheet As Worksheet, ByVal rangeName As String) As String
    On Error Resume Next
    SafeLoadNamedRange = wsSheet.Range(rangeName).Value
    If Err.Number <> 0 Then
        SafeLoadNamedRange = ""
    End If
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Sub       : DismissModalPopup
' Purpose   : Dismisses any modal popup (e.g., "end-to-end encrypted" notifications)
'             using primary or alternate CSS selector.
'---------------------------------------------------------------------------------------
Private Sub DismissModalPopup()
    On Error Resume Next
    If BOT.IsElementPresent(By.css(CSSClassModalPopup)) Then
        BOT.SendKeys (ks.Escape)
        BOT.Wait 300
    ElseIf Len(CSSClassModalPopup_Alt) > 0 Then
        If BOT.IsElementPresent(By.css(CSSClassModalPopup_Alt)) Then
            BOT.SendKeys (ks.Escape)
            BOT.Wait 300
        End If
    End If
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Function  : DeleteChromeUserData
' Purpose   : Deletes the Chrome user data folder.
'             Fixes corrupted profile issues (e.g. SessionNotCreatedError) and
'             stale lock files left behind after an unclean shutdown.
'             Note: The user will need to re-scan the WhatsApp QR code after this.
' Returns   : True if the folder was deleted successfully or didn't exist.
'---------------------------------------------------------------------------------------
Private Function DeleteChromeUserData() As Boolean
    On Error Resume Next
    
    Dim userDataPath As String
    Dim rng As Range
    
    Set rng = ThisWorkbook.Names("UserDataFolderPath").RefersToRange
    If rng Is Nothing Then
        DeleteChromeUserData = False
        Exit Function
    End If
    
    userDataPath = rng.Value
    If Len(userDataPath) = 0 Then
        DeleteChromeUserData = False
        Exit Function
    End If
    
    ' Use LibFileTools.DeleteFolder (handles locked files, recursive deletion, etc.)
    DeleteChromeUserData = LibFileTools.DeleteFolder(userDataPath, deleteContents:=True, failIfMissing:=False)
    
    On Error GoTo 0
End Function


Public Sub WhatsAppBOT()

    ErrorHandling.ClearAdditionalErrorInfo
    
    ' Disable events at the beginning to prevent worksheet events from interfering
    Application.EnableEvents = False
    
    ' Retrieve the latest XPath values from the API
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Backend_Settings")
    If ws.Range("RETRIEVE_LATEST_XPATHS").Value = True Then
        ' Try to parse XPaths
        On Error GoTo XPathError
        If Not ParseXPathsFromAPI() Then
            ' If it's a server connection error, the function returns False silently and we continue.
        End If
        On Error GoTo ErrHandler
    End If
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = True
    
    ' Load primary XPaths
    xPathInvalidPhoneNumber = ws.Range("XPathInvalidPhoneNumber").Value
    xPathTextInputField = ws.Range("XPathTextInputField").Value
    xPathSearchInputField = ws.Range("XPathSearchInputField").Value
    xPathNoContactFound = ws.Range("XPathNoContactFound").Value
    xPathAttachmentButton = ws.Range("XPathAttachmentButton").Value
    xPathMultipleAttachmentButton = ws.Range("XPathAttachmentButton").Value
    CSSClassModalPopup = ws.Range("CSSClassModalPopup").Value
    
    ' Load alternate XPaths for A/B fallback (safe load - blank if not defined)
    xPathInvalidPhoneNumber_Alt = SafeLoadNamedRange(ws, "XPathInvalidPhoneNumber_Alt")
    xPathTextInputField_Alt = SafeLoadNamedRange(ws, "XPathTextInputField_Alt")
    xPathSearchInputField_Alt = SafeLoadNamedRange(ws, "XPathSearchInputField_Alt")
    xPathNoContactFound_Alt = SafeLoadNamedRange(ws, "XPathNoContactFound_Alt")
    xPathAttachmentButton_Alt = SafeLoadNamedRange(ws, "XPathAttachmentButton_Alt")
    xPathMultipleAttachmentButton_Alt = SafeLoadNamedRange(ws, "XPathMultipleAttachmentButton_Alt")
    CSSClassModalPopup_Alt = SafeLoadNamedRange(ws, "CSSClassModalPopup_Alt")
    
    ' Load search result XPaths for clicking contacts
    xPathFirstSearchResult = SafeLoadNamedRange(ws, "XPathFirstSearchResult")
    xPathFirstSearchResult_Alt = SafeLoadNamedRange(ws, "XPathFirstSearchResult_Alt")
    
    Dim MessageValue As String
    Dim KeepLoginCredentials As String
    Dim SearchText As String
    Dim DefaultDelay As Long
    Dim RandomDelay As Long
    Dim LastRowStatus As Long
    Dim LastRowNumber As Long
    Dim LastRowText As Long
    Dim IsValidContact As Boolean
    Dim SendingMethod As String

    ' Variables for daily message limit
    Dim AppName As String
    Dim SettingSection As String
    Dim LastSentDateKey As String
    Dim MessageCountKey As String
    Dim lastSentDate As String
    Dim messageCount As Long
    Dim MaxMessages As Long
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("BOT")
    Set wsSettings = Settings
                
    KeepLoginCredentials = wsSettings.Range("KeepLoginCredentials")
    SendingMethod = wsSettings.Range("SendingMethod")
    
    ' Add additional information for Error Handler
    ErrorHandling.AppendAdditionalErrorInfo "Sending Method: " & SendingMethod
    ErrorHandling.AppendAdditionalErrorInfo "Using Default Chrome Browser: " & LCase(Trim(wsSettings.Range("UseDefaultChromeBinary").Value))

    ' --- Daily Message Limit for Free Version ---
    ' Check limit BEFORE initializing Selenium or the browser to avoid unnecessary resource usage
    AppName = "WhatsAppBlaster"
    SettingSection = "Settings"
    LastSentDateKey = "LastSentDate"
    MessageCountKey = "MessageCount"
    MaxMessages = 20 ' Set the daily message limit

    ' Get the last sent date and message count from the registry
    On Error Resume Next
    lastSentDate = GetSetting(AppName, SettingSection, LastSentDateKey, "")
    messageCount = CLng(GetSetting(AppName, SettingSection, MessageCountKey, 0))
    On Error GoTo ErrHandler

    ' If the last sent date is not today, reset the message count
    If lastSentDate <> CStr(Date) Then
        messageCount = 0
        On Error Resume Next
        SaveSetting AppName, SettingSection, LastSentDateKey, CStr(Date)
        SaveSetting AppName, SettingSection, MessageCountKey, CStr(messageCount)
        On Error GoTo ErrHandler
    End If

    ' Check limit BEFORE starting the browser
    If messageCount >= MaxMessages Then
        Dim purchaseResponse As VbMsgBoxResult
        purchaseResponse = MsgBox("You have reached the daily limit of " & MaxMessages & " messages for the free version." & vbCrLf & vbCrLf & _
               "Would you like to upgrade to the PRO version for unlimited messaging?", vbYesNo + vbQuestion, "Daily Limit Reached")

        If purchaseResponse = vbYes Then
            On Error Resume Next
            ActiveWorkbook.FollowHyperlink LINK_WHATSAPP_PRO, NewWindow:=True
            On Error GoTo 0
        End If

        Application.EnableEvents = True
        Exit Sub
    End If
    ' --- End of Daily Message Limit ---

    ' Check if Selenium is installed (late binding so the project compiles
    ' even when the Selenium type library is not registered on this machine)
    On Error GoTo SeleniumNotInstalledError
    Set By = CreateObject("Selenium.By")
    Set ks = CreateObject("Selenium.Keys")
    On Error GoTo ErrHandler
    
    ' Initialize the WebDriver with proper configuration
    Set BOT = InitWebDriver(KeepLoginCredentials)

    ' Determine last rows
    LastRowStatus = ws.Cells(rows.Count, BotColumn.wcStatus).End(xlUp).Row
    LastRowNumber = ws.Cells(rows.Count, BotColumn.wcNumber).End(xlUp).Row
    LastRowText = ws.Cells(rows.Count, BotColumn.wcText).End(xlUp).Row

    ' Set status column width immediately to ensure visibility
    ws.Columns(BotColumn.wcStatus).ColumnWidth = 30
    
    ' Clear status cells
    If LastRowStatus > FirstRow - 1 Then
        ws.Range(ws.Cells(FirstRow, BotColumn.wcStatus), _
            ws.Cells(LastRowStatus, BotColumn.wcStatus)).ClearContents
            
        ' Remove any hyperlinks from the status column to prevent underlines
        On Error Resume Next
        Dim hyperlinksRange As Range
        Set hyperlinksRange = ws.Range(ws.Cells(FirstRow, BotColumn.wcStatus), _
            ws.Cells(LastRowStatus, BotColumn.wcStatus))
        
        Dim cell As Range
        For Each cell In hyperlinksRange
            If Not IsEmpty(cell.Hyperlinks) Then
                If cell.Hyperlinks.Count > 0 Then
                    cell.Hyperlinks.Delete
                End If
            End If
            ' Set font underline property to False
            cell.Font.Underline = False
        Next cell
        On Error GoTo ErrHandler
    End If

    ' Init New Chrome instance & navigate to WebWhatsApp
    On Error GoTo ChromeStartError
    BOT.Start "chrome"
    On Error GoTo ErrHandler
    
    ' Capture the user agent information for error handling
    Dim userAgentString As String
    userAgentString = BOT.ExecuteScript("return navigator.userAgent;")
    ErrorHandling.AppendAdditionalErrorInfo "Browser: " & userAgentString
    
    BOT.Get "https://web.whatsapp.com/"

QRCodePrompt:
    ' Ask user to scan the QR code. Once logged in, continue with the macro
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("A new browser window has opened for Web WhatsApp." & vbCrLf & _
                          "If you're not logged in yet, please scan the QR code with your phone." & vbCrLf & _
                          "If you already see your chats, click OK to continue." & vbCrLf & _
                          "Click Cancel to exit.", _
                          vbOKCancel + vbInformation, "WhatsApp Bot - Login")
    If userResponse = vbCancel Then
        BOT.Quit
        GoTo CleanUp
    End If

    If Not WaitForPageLoad(BOT, 10) Then
        ' Handle the scenario where the page didn't load in time
        MsgBox "Page did not load within the specified time.", vbCritical, "WhatsApp Bot"
        BOT.Quit
        GoTo CleanUp
    End If

    ' Capture WhatsApp Web version for debugging
    m_WAVersion = GetWhatsAppVersion()
    ErrorHandling.AppendAdditionalErrorInfo "WhatsApp Web Version: " & m_WAVersion

    DefaultDelay = wsSettings.Range("DefaultDelay").Value

    '--- Track only the last phone number and whether its chat is open
    Dim lastNumber As String
    Dim lastChatOpen As Boolean
    
    lastNumber = ""         ' Start empty
    lastChatOpen = False    ' Not open yet

    ' Loop through each row
    For i = FirstRow To LastRowNumber
        ' Check if the daily message limit has been reached
        If messageCount >= MaxMessages Then
            purchaseResponse = MsgBox("You have reached the daily limit of " & MaxMessages & " messages for the free version." & vbCrLf & vbCrLf & _
                   "Would you like to upgrade to the PRO version for unlimited messaging?", vbYesNo + vbQuestion, "Daily Limit Reached")
            
            If purchaseResponse = vbYes Then
                On Error Resume Next
                ActiveWorkbook.FollowHyperlink LINK_WHATSAPP_PRO, NewWindow:=True
                On Error GoTo 0 ' Or handle error appropriately
            End If
            
            GoTo CleanUp
        End If

        MessageValue = ws.Cells(i, BotColumn.wcText).Value
        SearchText = ws.Cells(i, BotColumn.wcNumber).Value
        
        If SearchText = vbNullString Then
            ws.Cells(i, BotColumn.wcStatus).Value = _
                "Error: No number provided | " & Format(Now, "mm/dd/yyyy HH:mm:ss")
            ws.Columns(BotColumn.wcStatus).AutoFit
            If ws.Columns(BotColumn.wcStatus).ColumnWidth < 30 Then
                ws.Columns(BotColumn.wcStatus).ColumnWidth = 30
            End If
            GoTo NextIteration
        End If
        
        If Len(Trim(MessageValue)) = 0 Then
            ws.Cells(i, BotColumn.wcStatus).Value = _
                "Error: No message provided | " & Format(Now, "mm/dd/yyyy HH:mm:ss")
            ws.Columns(BotColumn.wcStatus).AutoFit
            If ws.Columns(BotColumn.wcStatus).ColumnWidth < 30 Then
                ws.Columns(BotColumn.wcStatus).ColumnWidth = 30
            End If
            GoTo NextIteration
        End If

        RandomDelay = CalcRandomDelay()
        
        '--- Decide if we need to open the chat or not ---
        Dim needToOpenChat As Boolean
        needToOpenChat = True
        
        ' If same number as last time and last chat was open/successful, skip re-check
        If (SearchText = lastNumber) And (lastChatOpen = True) Then
            needToOpenChat = False
        End If
        
        If needToOpenChat Then

            IsValidContact = IsValidContactSavedNumber(i)
            
            If Not IsValidContact Then
                ' Mark chat not open, skip sending
                lastChatOpen = False
                GoTo NextIteration
            Else
                ' If it's valid, store the number as "open"
                lastNumber = SearchText
                lastChatOpen = True
            End If
        End If
        
        ' If we get here, the chat is open (either from this iteration or previous)
        ' Check if the message contains emojis before deciding which sending method to use
        If mEmojis.HasEmoji(MessageValue) Then
            ' Update status in the sheet with a message about PRO version
            ws.Cells(i, BotColumn.wcStatus).Value = _
                "Error: Message contains emoji. Upgrade to PRO for emoji support." & _
                " | " & Format(Now, "mm/dd/yyyy HH:mm:ss")

            ' Remove any existing underline formatting before adding hyperlink
            ws.Cells(i, BotColumn.wcStatus).Font.Underline = False
            
            ' Add a hyperlink to the PRO version purchase page
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, BotColumn.wcStatus), _
                              Address:=LINK_WHATSAPP_PRO, _
                              TextToDisplay:=ws.Cells(i, BotColumn.wcStatus).Value

            ws.Columns(BotColumn.wcStatus).AutoFit
            If ws.Columns(BotColumn.wcStatus).ColumnWidth < 30 Then
                ws.Columns(BotColumn.wcStatus).ColumnWidth = 30
            End If
            
        Else
            ' Send the message using the selected method
            If SendingMethod = "Default" Then
                SendTextWithClipboard MessageValue
            Else
                SendTextmessage MessageValue
            End If
            
            ' Increment the message count and save it to the registry
            messageCount = messageCount + 1
            On Error Resume Next
            SaveSetting AppName, SettingSection, MessageCountKey, CStr(messageCount)
            On Error GoTo ErrHandler
           
            ' Wait the default + random delay
            BOT.Wait DefaultDelay
            BOT.Wait RandomDelay
        End If
        
NextIteration:
    Next i

    On Error Resume Next
    BOT.Quit
    Set BOT = Nothing
    On Error GoTo ErrHandler
    
    ufSuccess.Show
    Set ufSuccess = Nothing

    ' Auto-adjust the status column width after all messages have been processed
    ws.Columns(BotColumn.wcStatus).AutoFit
    If ws.Columns(BotColumn.wcStatus).ColumnWidth < 30 Then
        ws.Columns(BotColumn.wcStatus).ColumnWidth = 30
    End If
    
    GoTo CleanUp
    
ErrHandler:
    Dim errHandlerNum As Long
    Dim errHandlerSrc As String
    Dim errHandlerDesc As String
    Dim errHandlerLine As Long
    errHandlerNum = Err.Number
    errHandlerSrc = Err.Source
    errHandlerDesc = Err.Description
    errHandlerLine = Erl
    On Error Resume Next
    If Not BOT Is Nothing Then BOT.Quit
    Set BOT = Nothing
    On Error GoTo 0
    DisplayError errHandlerSrc, errHandlerDesc, "mWhatsAppBOT.WhatsAppBOT", errHandlerLine
    GoTo CleanUp
    
ChromeStartError:
    Dim chromeErrNum As Long
    Dim chromeErrDesc As String
    chromeErrNum = Err.Number
    chromeErrDesc = Err.Description
    
    ' Resume to exit the active error state ? without this, any subsequent
    ' statement that triggers an error would cause an unhandled runtime error.
    Resume ChromeStartErrorResume
ChromeStartErrorResume:
    On Error GoTo 0
    
    ' Detect whether this is a missing Chrome/ChromeDriver issue vs. a version
    ' mismatch issue vs. a profile issue ? each needs a different message.
    Dim chromeErrDescLower As String
    chromeErrDescLower = LCase(chromeErrDesc)
    
    Dim isMissingBinary As Boolean
    isMissingBinary = (InStr(1, chromeErrDescLower, "cannot find chrome binary") > 0) Or _
                      (InStr(1, chromeErrDescLower, "chrome failed to start") > 0) Or _
                      (InStr(1, chromeErrDescLower, "chrome not found") > 0) Or _
                      (InStr(1, chromeErrDescLower, "chromedriver") > 0 And InStr(1, chromeErrDescLower, "not found") > 0)
    
    Dim isVersionMismatch As Boolean
    isVersionMismatch = (InStr(1, chromeErrDescLower, "sessionnotcreatederror") > 0) Or _
                        (InStr(1, chromeErrDescLower, "session not created") > 0) Or _
                        (InStr(1, chromeErrDescLower, "this version of chromedriver only supports") > 0)
    
    Dim resetResponse As VbMsgBoxResult
    
    If isMissingBinary Then
        ' Chrome or ChromeDriver is not installed ? profile reset won't help
        MsgBox "Chrome browser could not be started correctly." & vbCrLf & vbCrLf & _
               "Error details:" & vbCrLf & _
               "  Error #" & chromeErrNum & ": " & chromeErrDesc & vbCrLf & vbCrLf & _
               "It appears that Chrome or the ChromeDriver is not installed or not found." & vbCrLf & _
               "Please follow the setup guide to install the required components.", _
               vbExclamation, "Chrome Setup Required"
        GoTo ChromeStartErrorFinal
    ElseIf isVersionMismatch Then
        ' ChromeDriver version does not match Chrome version ? profile reset won't help
        MsgBox "Chrome browser could not be started correctly." & vbCrLf & vbCrLf & _
               "Error details:" & vbCrLf & _
               "  Error #" & chromeErrNum & ": " & chromeErrDesc & vbCrLf & vbCrLf & _
               "Your ChromeDriver version does not match your Chrome browser version." & vbCrLf & vbCrLf & _
               "To fix this, either:" & vbCrLf & _
               "  1. Update your ChromeDriver to match your Chrome version, or" & vbCrLf & _
               "  2. Use 'Chrome for Testing' (recommended) which does not auto-update." & vbCrLf & vbCrLf & _
               "Please follow the setup guide for instructions.", _
               vbExclamation, "ChromeDriver Version Mismatch"
        GoTo ChromeStartErrorFinal
    Else
        ' Profile or other issue ? offer reset
        resetResponse = MsgBox("Chrome browser could not be started correctly." & vbCrLf & vbCrLf & _
                               "Error details:" & vbCrLf & _
                               "  Error #" & chromeErrNum & ": " & chromeErrDesc & vbCrLf & vbCrLf & _
                               "This can often be fixed by resetting the Chrome user profile." & vbCrLf & _
                               "(You will need to re-scan the WhatsApp QR code afterwards.)" & vbCrLf & vbCrLf & _
                               "Would you like to reset the Chrome profile and retry?", _
                               vbYesNo + vbExclamation, "Chrome Setup Required")
    End If
    
    If resetResponse = vbYes Then
        On Error Resume Next
        If Not BOT Is Nothing Then BOT.Quit
        Set BOT = Nothing
        On Error GoTo 0
        
        If DeleteChromeUserData() Then
            ' Retry starting Chrome
            Dim retryErr As Long
            On Error Resume Next
            Set BOT = InitWebDriver(KeepLoginCredentials)
            retryErr = Err.Number
            On Error GoTo 0
            
            ' Check InitWebDriver succeeded before calling .Start
            If retryErr <> 0 Or BOT Is Nothing Then
                GoTo ChromeStartErrorFinal
            End If
            
            On Error Resume Next
            BOT.Start "chrome"
            retryErr = Err.Number
            On Error GoTo 0
            
            If retryErr <> 0 Then
                ' Retry also failed ? fall through to ChromeStartErrorFinal
                On Error Resume Next
                If Not BOT Is Nothing Then BOT.Quit
                Set BOT = Nothing
                On Error GoTo 0
                GoTo ChromeStartErrorFinal
            End If
            
            On Error GoTo ErrHandler
            
            ' If we get here, the retry succeeded
            Dim retryUserAgent As String
            retryUserAgent = BOT.ExecuteScript("return navigator.userAgent;")
            ErrorHandling.AppendAdditionalErrorInfo "Browser: " & retryUserAgent
            BOT.Get "https://web.whatsapp.com/"
            GoTo ChromeRetrySuccess
        Else
            MsgBox "Could not delete the Chrome user profile." & vbCrLf & _
                   "Please close all Chrome windows and try again.", _
                   vbOKOnly + vbExclamation, "WhatsApp Bot"
        End If
    End If
    
    ' If reset was declined or failed, offer the setup guide
ChromeStartErrorFinal:
    Dim chromeSetupResponse As VbMsgBoxResult
    chromeSetupResponse = MsgBox("Would you like to open the Chrome setup guide to fix this issue?", _
                               vbYesNo + vbQuestion, "Chrome Setup Required")
    
    If chromeSetupResponse = vbYes Then
        On Error Resume Next
        ThisWorkbook.FollowHyperlink "https://pythonandvba.com/go/whatsappblaster-how-to-set-up-chrome", NewWindow:=True
        On Error GoTo 0
    End If
    
    MsgBox "The WhatsApp Bot will now exit. Please follow the Chrome setup guide and try again.", _
           vbOKOnly + vbInformation, "WhatsApp Bot"
    
    On Error Resume Next
    If Not BOT Is Nothing Then BOT.Quit
    Set BOT = Nothing
    On Error GoTo 0
    GoTo CleanUp
    
ChromeRetrySuccess:
    GoTo QRCodePrompt

SeleniumNotInstalledError:
    Dim seleniumResponse As VbMsgBoxResult
    
    seleniumResponse = MsgBox("Selenium is not installed or properly configured." & vbCrLf & vbCrLf & _
                            "Selenium is required for the WhatsApp Bot to function." & vbCrLf & vbCrLf & _
                            "Would you like to watch the getting started video tutorial that explains how to set up Selenium?", _
                            vbYesNo + vbExclamation, "Selenium Setup Required")
    
    If seleniumResponse = vbYes Then
        ThisWorkbook.FollowHyperlink "https://pythonandvba.com/whatsapp-pro-tutorial", NewWindow:=True
    End If
    
    MsgBox "The WhatsApp Bot will now exit. Please follow the tutorial to set up Selenium and try again.", _
           vbOKOnly + vbInformation, "WhatsApp Bot"
    
    On Error Resume Next
    If Not BOT Is Nothing Then BOT.Quit
    Set BOT = Nothing
    On Error GoTo 0
    GoTo CleanUp

XPathError:
    ' Clean up and exit ? save error details before cleanup clears them
    Dim xpErrSrc As String
    Dim xpErrDesc As String
    Dim xpErrLine As Long
    xpErrSrc = Err.Source
    xpErrDesc = Err.Description
    xpErrLine = Erl
    On Error Resume Next
    If Not BOT Is Nothing Then BOT.Quit
    Set BOT = Nothing
    On Error GoTo 0
    DisplayError xpErrSrc, xpErrDesc, "mWhatsAppBOT.WhatsAppBOT", xpErrLine
    GoTo CleanUp

CleanUp:
    On Error Resume Next
    If Not BOT Is Nothing Then BOT.Quit
    Set BOT = Nothing
    On Error GoTo 0
    Application.EnableEvents = True
    Exit Sub
End Sub


Private Function IsValidContactSavedNumber(ByVal i As Long) As Boolean
    ' ================= CONFIGURATION =================
    Dim MaxSearchRetries As Long: MaxSearchRetries = 3          ' Retry the entire search up to 3 times
    Dim WaitAfterType As Long: WaitAfterType = 1500             ' Wait for search results to load after typing
    Dim WaitForResultsMax As Long: WaitForResultsMax = 5000     ' Max wait for search results to appear
    Dim WaitAfterClick As Long: WaitAfterClick = 1500           ' Wait after clicking a search result
    Dim ChatVerifyTimeout As Long: ChatVerifyTimeout = 8000     ' Max wait to verify chat opened

    On Error GoTo ErrHandler
    
    Dim SearchText As String
    SearchText = ws.Cells(i, BotColumn.wcNumber)
    
    Dim searchRetry As Long
    Dim searchEl As Object
    Dim contactClicked As Boolean
    Dim typedContent As String
    contactClicked = False
    
    For searchRetry = 1 To MaxSearchRetries
        ' ================= STEP 1: Find and clear the search field =================
        Set searchEl = FindElementWithFallback(xPathSearchInputField, xPathSearchInputField_Alt, 5000)
        If searchEl Is Nothing Then
            ' Search field not found, retry
            BOT.Wait 1000
            GoTo RetrySearch
        End If
        
        ' Clear the search field robustly
        ClearSearchField
        BOT.Wait 300
        
        ' ================= STEP 2: Type the search text =================
        Set searchEl = FindElementWithFallback(xPathSearchInputField, xPathSearchInputField_Alt, 3000)
        If searchEl Is Nothing Then GoTo RetrySearch
        
        searchEl.SendKeys (SearchText)
        BOT.Wait 500
        
        ' ================= STEP 2b: Verify text was actually typed =================
        typedContent = ""
        On Error Resume Next
        typedContent = searchEl.Attribute("textContent")
        On Error GoTo ErrHandler
        
        If Len(Trim(typedContent)) = 0 Then
            If searchRetry = MaxSearchRetries Then
                ws.Cells(i, BotColumn.wcStatus).Value = _
                    "Error: Search field not responding (XPath may be outdated) | WA: " & m_WAVersion & " | " & Format(Now, "mm/dd/yyyy HH:mm:ss")
                IsValidContactSavedNumber = False
                Exit Function
            End If
            GoTo RetrySearch
        End If
        
        ' ================= STEP 3: Wait for search results to load =================
        BOT.Wait WaitAfterType
        
        ' ================= STEP 4: Check if "no contact found" =================
        If IsElementPresentWithFallback(xPathNoContactFound, xPathNoContactFound_Alt) Then
            ClearSearchField
            
            ' Update status message to include information about the PRO version
            ws.Cells(i, BotColumn.wcStatus).Value = _
                "Error: No contact found. Upgrade to PRO to send messages to unsaved contacts | " & Format(Now, "mm/dd/yyyy HH:mm:ss")
            
            ' Remove any existing underline formatting before adding hyperlink
            ws.Cells(i, BotColumn.wcStatus).Font.Underline = False
            
            ' Add a hyperlink to the PRO version purchase page
            ws.Hyperlinks.Add Anchor:=ws.Cells(i, BotColumn.wcStatus), _
                              Address:=LINK_WHATSAPP_PRO, _
                              TextToDisplay:=ws.Cells(i, BotColumn.wcStatus).Value
            
            IsValidContactSavedNumber = False
            Exit Function
        End If
        
        ' ================= STEP 5: Click the first search result =================
        Dim resultEl As Object
        Set resultEl = Nothing
        
        ' Try primary search result XPath
        If Len(xPathFirstSearchResult) > 0 Then
            If WaitForElement(xPathFirstSearchResult, WaitForResultsMax, 300) Then
                Set resultEl = BOT.FindElementByXPath(xPathFirstSearchResult)
            End If
        End If
        
        ' Try alternate search result XPath
        If resultEl Is Nothing And Len(xPathFirstSearchResult_Alt) > 0 Then
            If WaitForElement(xPathFirstSearchResult_Alt, 2000, 300) Then
                Set resultEl = BOT.FindElementByXPath(xPathFirstSearchResult_Alt)
            End If
        End If
        
        ' If we found a search result, click it; otherwise fall back to Enter
        If Not resultEl Is Nothing Then
            On Error Resume Next
            resultEl.Click
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo ErrHandler
                ' Clicking failed, fall back to pressing Enter
                BOT.SendKeys (ks.Enter)
            End If
            On Error GoTo ErrHandler
        Else
            ' No search result element found, fall back to pressing Enter
            BOT.SendKeys (ks.Enter)
        End If
        
        BOT.Wait WaitAfterClick
        
        ' ================= STEP 6: Verify the chat actually opened =================
        If VerifyChatOpen(ChatVerifyTimeout) Then
            contactClicked = True
            Exit For
        End If
        
RetrySearch:
        BOT.Wait 500
    Next searchRetry
    
    ' ================= FINAL RESULT =================
    If contactClicked Then
        IsValidContactSavedNumber = True
        ' Dismiss any modal popups (e.g., end-to-end encrypted notifications)
        DismissModalPopup
    Else
        ' All retries exhausted - could not open the chat
        ClearSearchField
        ws.Cells(i, BotColumn.wcStatus).Value = _
            "Error: Could not open chat for contact | WA: " & m_WAVersion & " | " & Format(Now, "mm/dd/yyyy HH:mm:ss")
        IsValidContactSavedNumber = False
    End If
    
    Exit Function
    
ErrHandler:
    RaiseError Err.Number, Err.Source, "mWhatsAppBOT.IsValidContactSavedNumber", Err.Description, Erl
End Function


'---------------------------------------------------------------------------------------
' Procedure : CalcRandomDelay
' Purpose   : Calculate Random Number using Application.WorksheetFunction.RandBetween
'---------------------------------------------------------------------------------------
'
Private Function CalcRandomDelay() As Long

    On Error GoTo ErrHandler
    Dim minRandomDelay As Long
    Dim maxRandomDelay As Long
    Dim RandomDelay As Long
    Dim useRandomDelay As String
    
    useRandomDelay = wsSettings.Range("useRandomDelay").Value
    If useRandomDelay = "Yes" Then
        minRandomDelay = wsSettings.Range("minRandomDelay").Value
        maxRandomDelay = wsSettings.Range("maxRandomDelay").Value
        ' Guard: swap if user entered min > max
        If minRandomDelay > maxRandomDelay Then
            Dim tmp As Long
            tmp = minRandomDelay
            minRandomDelay = maxRandomDelay
            maxRandomDelay = tmp
        End If
        RandomDelay = Application.WorksheetFunction.RandBetween(minRandomDelay, maxRandomDelay)
    Else
        RandomDelay = 0
    End If
    
    CalcRandomDelay = RandomDelay
    Exit Function
    
ErrHandler:
    RaiseError Err.Number, Err.Source, "mWhatsAppBOT.CalcRandomDelay", Err.Description, Erl

End Function

'---------------------------------------------------------------------------------------
' Procedure : InitWebDriver
' Purpose   : Initialize the Chrome WebDriver.
'             - If "UseDefaultChromeBinary" = "Yes", we use the system's default Chrome.
'             - If "UseDefaultChromeBinary" = "No", we must have a valid "CustomChromeBinaryPath"
'               ending in chrome.exe (no quotes). If invalid, we exit with a user-friendly message.
'
'             Also handles .NET Framework error with a Yes/No prompt.
'---------------------------------------------------------------------------------------
'
Private Function InitWebDriver(KeepLoginCredentials As String) As Object

    On Error GoTo ErrHandler

    Dim DefaultChromeDecision As String      ' "Yes" or "No"
    Dim CustomChromePath As String          ' Full path to chrome.exe
    Dim webDriver As Object                  ' Local variable for WebDriver (late binding)
    
    ' Read from the Settings worksheet:
    DefaultChromeDecision = LCase(Trim(wsSettings.Range("UseDefaultChromeBinary").Value))
    CustomChromePath = wsSettings.Range("CustomChromeBinaryPath").Value
    
    ' Remove any quotation marks in case user pasted them like "C:\Chrome\chrome.exe"
    CustomChromePath = Replace(CustomChromePath, """", "")
    CustomChromePath = Trim(CustomChromePath)

    ' Create the Selenium WebDriver (late binding)
    Set webDriver = CreateObject("Selenium.WebDriver")
    webDriver.Timeouts.ImplicitWait = ImplicitWait
    webDriver.Timeouts.PageLoad = PageLoad
    webDriver.Timeouts.Server = TimeoutServer
    
    ' Add command-line arguments to Chrome
    webDriver.AddArgument "--disable-popup-blocking"
    webDriver.AddArgument "--disable-notifications"

    ' Set browser language to English
    webDriver.AddArgument "--lang=en"
    
    ' Keep user login (avoid scanning QR each time)
    If KeepLoginCredentials = "Yes" Then
    
        Dim chosenFolder As String
        Dim fallbackOption As Integer  ' 1 = current folder, 2 = LOCALAPPDATA, 3 = Temp
        
        fallbackOption = 0
    
        ' --- Option 1: Use current workbook's folder if available ---
        Dim currentDir As String
        currentDir = ThisWorkbook.path
        If currentDir <> "" Then
            ' Convert the path (if possible). If conversion fails, GetLocalPath returns the input.
            Dim localCurrent As String
            localCurrent = GetLocalPath( _
                                fullPath:=DecodeURL(currentDir), _
                                rebuildCache:=True, _
                                returnInputOnFail:=True _
                            )
            If Len(localCurrent) > 0 And IsFolderEditable(localCurrent) Then
                chosenFolder = BuildPath(localCurrent, "ChromeUserData")
                fallbackOption = 1
            End If
        End If
    
        ' --- Option 2: Use a folder in LOCALAPPDATA ("WhatsAppBlaster") ---
        If chosenFolder = "" Then
            Dim appDataFolder As String
            appDataFolder = BuildPath(Environ("LOCALAPPDATA"), "WhatsAppBlaster")
            Call CreateFolder(appDataFolder)
            If IsFolderEditable(appDataFolder) Then
                chosenFolder = BuildPath(appDataFolder, "ChromeUserData")
                fallbackOption = 2
            End If
        End If
    
        ' --- Option 3: Fall back to the Temp folder ---
        If chosenFolder = "" Then
            chosenFolder = BuildPath(Environ("Temp"), "ChromeUserData")
            fallbackOption = 3
        End If
    
        ' Create the chosen folder (if it doesn't already exist)
        Call CreateFolder(chosenFolder)

        ' Verify that the folder is writable.
        If Not IsFolderEditable(chosenFolder) Then
            Dim errMsg As String
            errMsg = "Error (Version: " & VERSION_NUMBER & "): Unable to create a writable folder for Chrome user data." & vbCrLf & _
                     "Please try running WhatsApp Blaster as an administrator." & vbCrLf & _
                     "Hint: Right-click your Excel shortcut and select 'Run as administrator'." & vbCrLf & _
                     "Alternatively, save your file on a local drive where you have write permissions."
            MsgBox errMsg, vbCritical, "WhatsApp Bot"
            Err.Raise 1003, "WhatsApp Bot", "Failed to create user data directory. (Version: " & VERSION_NUMBER & ")"
            Exit Function
        End If
        ' Log the chosen folder in a named range "UserDataFolderPath" for later reference.
        Dim rng As Range
        On Error Resume Next
        Set rng = ThisWorkbook.Names("UserDataFolderPath").RefersToRange
        On Error GoTo 0
        
        If rng Is Nothing Then
            Dim errorMsg As String
            errorMsg = "Oops! Something seems off (Version: " & VERSION_NUMBER & ")." & vbCrLf & vbCrLf & _
                       "A required setting in your file is missing or has been modified." & vbCrLf & _
                       "This may affect login persistence." & vbCrLf & vbCrLf & _
                       "Recommended fix: Download a fresh copy of the WhatsApp Blaster template and try again." & vbCrLf & vbCrLf & _
                       "Technical info for support:" & vbCrLf & _
                       "Expected setting: 'UserDataFolderPath'" & vbCrLf & _
                       "Attempted folder: " & chosenFolder
                       
            MsgBox errorMsg, vbExclamation, "WhatsApp Bot - Missing Setting"
            Err.Raise 1004, "WhatsApp Bot", "Missing named range 'UserDataFolderPath'. Attempted folder: " & chosenFolder & " (Version: " & VERSION_NUMBER & ")"
            Exit Function
        Else
            rng.Value = chosenFolder
        End If
    
        ' Set the Chrome user data directory for Selenium.
        webDriver.AddArgument "--user-data-dir=" & chosenFolder
    
    End If
    
    ' If user selected "No", attempt to use a custom Chrome
    If DefaultChromeDecision = "no" Then
        
        ' 1) Make sure we have something
        If Len(CustomChromePath) = 0 Then
            MsgBox "You selected 'No' for 'Use Default Chrome', but no path was provided." & vbCrLf & _
                "Please enter a valid path to chrome.exe and try again.", _
                vbExclamation, "WhatsApp Bot"
            Err.Raise vbObjectError + 1005, "WhatsApp Bot", "No custom Chrome path provided. (Version: " & VERSION_NUMBER & ")"
            Exit Function
        End If
        
        ' 2) Must end with "chrome.exe"
        If Right(LCase(CustomChromePath), 10) <> "chrome.exe" Then
            MsgBox "The path you entered does not end with 'chrome.exe':" & vbCrLf & _
                CustomChromePath & vbCrLf & vbCrLf & _
                "Please correct the path and try again.", _
                vbExclamation, "WhatsApp Bot"
            Err.Raise vbObjectError + 1006, "WhatsApp Bot", "Custom Chrome path does not end with chrome.exe: " & CustomChromePath & " (Version: " & VERSION_NUMBER & ")"
            Exit Function
        End If
        
        ' 3) Check if the file actually exists
        If Dir(CustomChromePath) = "" Then
            MsgBox "We could not find the file here:" & vbCrLf & _
                CustomChromePath & vbCrLf & vbCrLf & _
                "Please check the path and try again.", _
                vbExclamation, "WhatsApp Bot"
            Err.Raise vbObjectError + 1007, "WhatsApp Bot", "Chrome binary not found at: " & CustomChromePath & " (Version: " & VERSION_NUMBER & ")"
            Exit Function
        End If
        
        ' If all checks passed, set the custom binary
        webDriver.SetBinary CustomChromePath
        
    End If
    
    ' Return the configured driver
    Set InitWebDriver = webDriver

    Exit Function
    
    ' ----------------------
    ' Error Handler
    ' ----------------------
ErrHandler:
    Select Case Err.Number
        
        ' .NET Framework not installed or not enabled
        Case -2146232576
            Dim response As VbMsgBoxResult
            response = MsgBox( _
                "It looks like the .NET Framework is missing or not enabled." & vbCrLf & _
                "Do you want to open the troubleshooting page?" & vbCrLf & _
                "(If you select 'No', the program will exit.)", _
                vbYesNo + vbQuestion, "WhatsApp Bot - .NET Framework Required")

            If response = vbYes Then
                ActiveWorkbook.FollowHyperlink "https://pythonandvba.com/automation-error", NewWindow:=True
            End If
            Err.Raise -2146232576, "WhatsApp Bot", ".NET Framework is missing or not enabled. (Version: " & VERSION_NUMBER & ")"

        Case Else
            RaiseError Err.Number, Err.Source, "mWhatsAppBOT.InitWebDriver", Err.Description, Erl
    End Select
    
End Function

' Reads from or writes to the Windows clipboard using htmlfile
Function Clipboard$(Optional s$)
    On Error GoTo ClipErr
    Dim v: v = s  'Cast to Variant for 64-bit VBA support
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(s): .SetData "text", v     ' Write to clipboard
                Case Else:   Clipboard = .GetData("text")  ' Read from clipboard
            End Select
        End With
    End With
    Exit Function
ClipErr:
    ' Clipboard may be locked by another application; retry once after a brief pause
    Static retried As Boolean
    If Not retried Then
        retried = True
        Application.Wait Now + TimeSerial(0, 0, 1)
        Resume
    End If
    retried = False
    Clipboard = ""
End Function


Private Function WaitForPageLoad(ByVal driver As Object, Optional maxWait As Long = 10) As Boolean
    Dim startTime As Single
    startTime = Timer

    Do While Timer - startTime < maxWait
        ' Check if the document is fully loaded
        If driver.ExecuteScript("return document.readyState") = "complete" Then
            WaitForPageLoad = True
            Exit Function
        End If
        driver.Wait 500  ' Wait for 0.5 seconds before retrying
    Loop

    ' If timeout is reached without full load, return False
    WaitForPageLoad = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : SendTextmessage
' Purpose   : Send 'normal' text message. Errors will be returned in the respective row
'---------------------------------------------------------------------------------------
'
Private Sub SendTextmessage(MessageValue As String)

    On Error GoTo ErrHandler
    Dim arrTextMessage As Variant
    Dim LenOfArray As Integer
    Dim line As Long
    Dim MaxRetries As Long: MaxRetries = 3
    Dim sendRetry As Long

    ''' Ensure the text input field is ready (with retry and fallback)
    Dim textInputEl As Object
    For sendRetry = 1 To MaxRetries
        Set textInputEl = FindElementWithFallback(xPathTextInputField, xPathTextInputField_Alt, 5000)
        If Not textInputEl Is Nothing Then Exit For
        BOT.Wait 1000
    Next sendRetry
    
    If textInputEl Is Nothing Then
        ws.Cells(i, BotColumn.wcStatus).Value = _
            "Error: Text input field not found | WA: " & m_WAVersion & " | " & Format(Now, "mm/dd/yyyy HH:mm:ss")
        Exit Sub
    End If
    
    textInputEl.Click
    BOT.Wait 300

    ''' Split text message based on "|" to identify new paragraph
    arrTextMessage = Split(MessageValue, "|")
    
    '''' Length of variable. If only one line, it returns 1
    LenOfArray = UBound(arrTextMessage) - LBound(arrTextMessage) + 1
    
    ''' Iterate over array and press Shift + Enter to create new paragraph
    For line = LBound(arrTextMessage) To UBound(arrTextMessage)
        BOT.Wait (500)
        BOT.SendKeys (arrTextMessage(line))
        BOT.Wait (500)
        If LenOfArray > 1 And line < UBound(arrTextMessage) Then
            ''' Create a new line by pressing Shift & Enter
            BOT.Keyboard.KeyDown (ks.Shift)
            BOT.SendKeys (ks.Enter)
            BOT.Keyboard.KeyUp (ks.Shift)
            BOT.Wait (500)
        End If
    Next line
    BOT.Wait (500)
    BOT.SendKeys (ks.Enter)
    ws.Cells(i, BotColumn.wcStatus).Value = "Sent: " & Format(Now, "mm/dd/yyyy HH:mm:ss")
    
    ' Ensure status column has minimum width of 30
    ws.Columns(BotColumn.wcStatus).AutoFit
    If ws.Columns(BotColumn.wcStatus).ColumnWidth < 30 Then
        ws.Columns(BotColumn.wcStatus).ColumnWidth = 30
    End If
    
    Exit Sub
    
ErrHandler:
    ws.Cells(i, BotColumn.wcStatus).Value = _
        "Error (Version: " & VERSION_NUMBER & "): " & Err.Number & "_" & Err.Description & ", " & Format(Now, "mm/dd/yyyy HH:mm:ss")
    
    ' Ensure status column has minimum width of 30
    ws.Columns(BotColumn.wcStatus).AutoFit
    If ws.Columns(BotColumn.wcStatus).ColumnWidth < 30 Then
        ws.Columns(BotColumn.wcStatus).ColumnWidth = 30
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SendTextWithClipboard
' Purpose   : Send 'normal' text message. Emojis will also work. Errors will be returned in the respective row
'---------------------------------------------------------------------------------------
Private Sub SendTextWithClipboard(MessageValue As String)
    On Error GoTo ErrHandler
    
    Dim MaxRetries As Long: MaxRetries = 3
    Dim sendRetry As Long
    
    ' 1) Convert pipe symbols "|" to line breaks.
    Dim finalText As String
    finalText = Replace(MessageValue, "|", vbCrLf)
    
    ' 2) Copy the entire text (including emojis) to the clipboard
    Clipboard finalText
    
    ' 3) Click the text input field in WhatsApp (with retry and fallback)
    Dim textInputEl As Object
    For sendRetry = 1 To MaxRetries
        Set textInputEl = FindElementWithFallback(xPathTextInputField, xPathTextInputField_Alt, 5000)
        If Not textInputEl Is Nothing Then Exit For
        BOT.Wait 1000
    Next sendRetry
    
    If textInputEl Is Nothing Then
        ws.Cells(i, BotColumn.wcStatus).Value = _
            "Error: Text input field not found | WA: " & m_WAVersion & " | " & Format(Now, "mm/dd/yyyy HH:mm:ss")
        Exit Sub
    End If
    
    textInputEl.Click
    BOT.Wait 300
    
    ' 4) Press Ctrl+V to paste the content
    BOT.Keyboard.KeyDown (ks.Control)
    BOT.SendKeys "v"
    BOT.Keyboard.KeyUp (ks.Control)
    BOT.Wait 300

    ' 5) Press Enter to send
    BOT.SendKeys ks.Enter
    
    ' 6) Update status in your sheet
    ws.Cells(i, BotColumn.wcStatus).Value = "Sent: " & Format(Now, "mm/dd/yyyy HH:mm:ss")
    
    ' Ensure status column has minimum width of 30
    ws.Columns(BotColumn.wcStatus).AutoFit
    If ws.Columns(BotColumn.wcStatus).ColumnWidth < 30 Then
        ws.Columns(BotColumn.wcStatus).ColumnWidth = 30
    End If
    
    Exit Sub
    
ErrHandler:
    ws.Cells(i, BotColumn.wcStatus).Value = _
        "Error (Version: " & VERSION_NUMBER & "): " & Err.Number & "_" & Err.Description & ", " & Format(Now, "mm/dd/yyyy HH:mm:ss")
    
    ' Ensure status column has minimum width of 30
    ws.Columns(BotColumn.wcStatus).AutoFit
    If ws.Columns(BotColumn.wcStatus).ColumnWidth < 30 Then
        ws.Columns(BotColumn.wcStatus).ColumnWidth = 30
    End If
End Sub

' Clear the WhatsApp BOT sheet data with user confirmation
' This sub will only clear the content, not delete rows
Public Sub ClearBotData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim lastRow As Long
    Dim hasData As Boolean
    Dim userResponse As VbMsgBoxResult
    Dim confirmMessage As String
    Dim successMessage As String
    Dim totalMessages As Long
    
    ' Check if the worksheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("BOT")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        RaiseError vbObjectError + 1004, "", "mWhatsAppBOT.ClearBotData", "The worksheet 'BOT' does not exist.", Erl
        Exit Sub
    End If
    
    ' Check if there's data in the sheet
    hasData = False
    lastRow = ws.Cells(ws.rows.Count, BotColumn.wcNumber).End(xlUp).Row
    
    ' If lastRow is less than FirstRow, there's no data
    If lastRow >= FirstRow Then
        ' Check if there's actual content in the cells
        On Error Resume Next
        Set dataRange = ws.Range(ws.Cells(FirstRow, BotColumn.wcNumber), ws.Cells(lastRow, BotColumn.wcStatus))
        On Error GoTo ErrorHandler
        
        ' Check if any cell in the range has content
        Dim cell As Range
        For Each cell In dataRange
            If Not IsEmpty(cell.Value) Then
                hasData = True
                Exit For
            End If
        Next cell
    End If
    
    ' If there's no data, inform the user and exit
    If Not hasData Then
        MsgBox "The WhatsApp BOT sheet is already empty.", vbInformation, "Clear BOT Data"
        Exit Sub
    End If
    
    ' Count messages
    totalMessages = 0
    
    Dim i As Long
    For i = FirstRow To lastRow
        If Not IsEmpty(ws.Cells(i, BotColumn.wcNumber).Value) Then
            totalMessages = totalMessages + 1
        End If
    Next i
    
    ' Prepare confirmation message with data summary
    confirmMessage = "Are you sure you want to clear all WhatsApp message data?" & vbNewLine & vbNewLine
    confirmMessage = confirmMessage & "Total messages to be cleared: " & totalMessages & vbNewLine & vbNewLine
    
    ' Add a preview of the data (first 3 rows or less)
    confirmMessage = confirmMessage & "Preview of messages to be cleared:" & vbNewLine
    
    Dim previewRows As Long
    previewRows = Application.WorksheetFunction.Min(3, lastRow - FirstRow + 1)
    
    For i = FirstRow To FirstRow + previewRows - 1
        If Not IsEmpty(ws.Cells(i, BotColumn.wcNumber).Value) Then
            ' Show phone number (masked for privacy)
            Dim phoneNumber As String
            phoneNumber = ws.Cells(i, BotColumn.wcNumber).Value
            If Len(phoneNumber) > 4 Then
                phoneNumber = Left(phoneNumber, 3) & "..." & Right(phoneNumber, 3)
            End If
            
            ' Show message preview (truncated)
            Dim messageText As String
            messageText = ""
            If Not IsEmpty(ws.Cells(i, BotColumn.wcText).Value) Then
                messageText = ws.Cells(i, BotColumn.wcText).Value
                If Len(messageText) > 25 Then
                    messageText = Left(messageText, 25) & "..."
                End If
            End If
            
            ' Show status
            Dim status As String
            status = ""
            If Not IsEmpty(ws.Cells(i, BotColumn.wcStatus).Value) Then
                status = ws.Cells(i, BotColumn.wcStatus).Value
            End If
            
            confirmMessage = confirmMessage & "- " & phoneNumber & ": " & messageText & " [" & status & "]" & vbNewLine
        End If
    Next i
    
    ' Add note if there are more rows
    If lastRow > FirstRow + previewRows - 1 Then
        confirmMessage = confirmMessage & "- And " & (lastRow - FirstRow - previewRows + 1) & " more message(s)..." & vbNewLine
    End If
    
    confirmMessage = confirmMessage & vbNewLine & "This action cannot be undone. Do you want to continue?"
    
    ' Ask for confirmation
    userResponse = MsgBox(confirmMessage, vbQuestion + vbYesNo, "Confirm Clear BOT Data")
    
    ' If user confirms, clear the data
    If userResponse = vbYes Then
        ' Clear the content without deleting rows
        On Error Resume Next
        ws.Range(ws.Cells(FirstRow, BotColumn.wcNumber), ws.Cells(lastRow, BotColumn.wcStatus)).ClearContents
        
        ' Check if there was an error
        If Err.Number <> 0 Then
            On Error GoTo ErrorHandler
            RaiseError Err.Number, Err.Source, "mWhatsAppBOT.ClearBotData", "Failed to clear the BOT data: " & Err.Description, Erl
            Exit Sub
        End If
        
        On Error GoTo ErrorHandler
        
        ' Set specific column widths for a clean appearance
        On Error Resume Next
        ws.Columns(BotColumn.wcNumber).ColumnWidth = 20    ' Receiver column
        ws.Columns(BotColumn.wcText).ColumnWidth = 90      ' Message column
        ws.Columns(BotColumn.wcStatus).ColumnWidth = 30    ' Status column
        If Err.Number <> 0 Then
            ' If there's an error setting column widths, just continue
            Err.Clear
        End If
        
        ' Set standard row height for all data rows
        On Error Resume Next
        ws.rows(FirstRow & ":" & lastRow).RowHeight = 15
        If Err.Number <> 0 Then
            ' If there's an error setting row heights, just continue
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        
        ' Show success message
        successMessage = "WhatsApp BOT data has been cleared successfully." & vbNewLine & vbNewLine
        successMessage = successMessage & "You can now:" & vbNewLine
        successMessage = successMessage & "1. Enter new phone numbers and messages" & vbNewLine
        successMessage = successMessage & "2. Run the WhatsApp BOT to send new messages"
        
        MsgBox successMessage, vbInformation, "BOT Data Cleared"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Display error and exit
    DisplayError Err.Source, Err.Description, "mWhatsAppBOT.ClearBotData", Erl
End Sub


Public Sub OpenFAQ()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink LINK_FAQ, NewWindow:=True
    On Error GoTo 0
End Sub
Public Sub OpenWhatsAppPRO()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink LINK_WHATSAPP_PRO, NewWindow:=True
    On Error GoTo 0
End Sub

Public Sub OpenFeatureRequest()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink LINK_FEATURE_REQUEST, NewWindow:=True
    On Error GoTo 0
End Sub

Public Sub OpenMessageDetailsDocs()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink LINK_MESSAGE_DETAILS_DOCS, NewWindow:=True
    On Error GoTo 0
End Sub
Public Sub OpenUsageGuidelines()
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink LINK_USAGE_GUIDELINES, NewWindow:=True
    On Error GoTo 0
End Sub

Public Sub PlaceholderTutorial()
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=PLACEHOLDER_TUTORIAL_LINK, NewWindow:=True
    On Error GoTo 0
End Sub
