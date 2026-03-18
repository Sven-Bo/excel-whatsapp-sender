Attribute VB_Name = "mXpaths"
Option Explicit

'==============================
' Module    : mXpaths
' Purpose   : Handles XPath retrieval and management for WhatsApp Bot
' Author    : Sven Bosau
'==============================

'==============================
' Configuration Settings
'==============================
Private Const START_ROW As Long = 8 ' Row to start writing data
Private Const API_BASE_URL As String = "https://api.pythonandvba.com/xpaths/"
Private Const LAST_XPATH_RETRIEVED As String = "LAST_XPATH_RETRIEVED"
Private Const DATE_FORMAT As String = "yyyy-mm-dd hh:mm:ss"

'==============================
' Public Methods
'==============================

' Public entry point for retrieving XPaths from API
Public Sub RunParseXPathsFromAPI()
    Call ParseXPathsFromAPI
End Sub

' Retrieves and parses XPaths from the API
' Returns: Boolean - True if successful, False otherwise
Public Function ParseXPathsFromAPI() As Boolean
    
    ' Initialize return value
    ParseXPathsFromAPI = False
    
    ' Declare variables
    Dim http As Object
    Dim json As Object
    Dim item As Variant
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim connectionMode As String
    Dim isServerConnectionError As Boolean
    
    ' Initialize error tracking
    isServerConnectionError = False
    
    On Error GoTo ErrorHandler
    
    '-----------------------------------
    ' 1. Validate environment
    '-----------------------------------
    
    ' Ensure the worksheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Backend_Settings")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1001, "mXpaths.ParseXPathsFromAPI", "The 'Backend_Settings' worksheet does not exist."
        Exit Function
    End If
    
    ' Get the connection mode from the named range
    On Error Resume Next
    connectionMode = ws.Range("CONNECTION_MODE").Value
    On Error GoTo ErrorHandler
    
    If IsEmpty(connectionMode) Then
        ' Default if not specified
        connectionMode = "WINHTTP"
    End If
    
    '-----------------------------------
    ' 2. Set up HTTP request
    '-----------------------------------
    
    ' Create the HTTP object based on connection mode
    If UCase(connectionMode) = "WINHTTP" Then
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Else
        ' Use ServerXML as the default
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    End If
    
    ' Make a GET request to the API
    On Error Resume Next
    http.Open "GET", API_BASE_URL & "Version_A", False
    http.Send
    
    ' Check if there was an error during the request
    If Err.Number <> 0 Then
        ' This is a server connection error, handle it silently
        isServerConnectionError = True
        Err.Clear
        GoTo CleanUp
    End If
    
    '-----------------------------------
    ' 3. Process the response
    '-----------------------------------
    
    ' Check if the request was successful
    If http.status = 200 Then
        ' Parse the JSON response
        On Error Resume Next
        Set json = JsonConverter.ParseJson(http.responseText)
        
        ' Check if there was an error parsing the JSON
        If Err.Number <> 0 Then
            Dim jsonErrorDesc As String
            jsonErrorDesc = Err.Description
            Err.Clear
            
            On Error GoTo ErrorHandler
            Err.Raise vbObjectError + 1003, "mXpaths.ParseXPathsFromAPI", "Failed to parse the JSON response: " & jsonErrorDesc
            Exit Function
        End If
        
        On Error GoTo ErrorHandler
        
        If json Is Nothing Then
            Err.Raise vbObjectError + 1004, "mXpaths.ParseXPathsFromAPI", "Failed to parse the JSON response."
            Exit Function
        End If
        
        '-----------------------------------
        ' 4. Update the worksheet
        '-----------------------------------
        
        ' Start writing from configured row
        rowIndex = START_ROW
        
        ' Clear previous data starting from START_ROW
        On Error Resume Next
        ws.rows(START_ROW & ":" & ws.rows.Count).ClearContents
        On Error GoTo ErrorHandler
        
        ' Loop through each item in the JSON array and create named ranges
        Dim xpathName As String
        
        For Each item In json
            xpathName = item("XPathName")
            ' Write data to the sheet
            ws.Cells(rowIndex, 1).Value = xpathName
            ws.Cells(rowIndex, 2).Value = item("XPathValue")
            ws.Cells(rowIndex, 3).Value = "Version_A"
            ws.Cells(rowIndex, 4).Value = item("LastUpdated")
            
            ' Create or update named range pointing to the value cell
            CreateOrUpdateNamedRange xpathName, ws.Cells(rowIndex, 2)
            
            rowIndex = rowIndex + 1
        Next item
        
        '-----------------------------------
        ' 5. Fetch Version_B (Alternate/Fallback XPaths)
        '-----------------------------------
        Dim httpB As Object
        Dim jsonB As Object
        
        ' Create a new HTTP object for Version_B
        If UCase(connectionMode) = "WINHTTP" Then
            Set httpB = CreateObject("WinHttp.WinHttpRequest.5.1")
        Else
            Set httpB = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        End If
        
        On Error Resume Next
        httpB.Open "GET", API_BASE_URL & "Version_B", False
        httpB.Send
        
        If Err.Number = 0 Then
            If httpB.status = 200 Then
                Set jsonB = JsonConverter.ParseJson(httpB.responseText)
                Err.Clear
                On Error GoTo ErrorHandler
                
                If Not jsonB Is Nothing Then
                    Dim itemB As Variant
                    For Each itemB In jsonB
                        xpathName = itemB("XPathName") & "_Alt"
                        ws.Cells(rowIndex, 1).Value = xpathName
                        ws.Cells(rowIndex, 2).Value = itemB("XPathValue")
                        ws.Cells(rowIndex, 3).Value = "Version_B"
                        ws.Cells(rowIndex, 4).Value = itemB("LastUpdated")
                        
                        ' Create or update named range for the alternate XPath
                        CreateOrUpdateNamedRange xpathName, ws.Cells(rowIndex, 2)
                        
                        rowIndex = rowIndex + 1
                    Next itemB
                End If
            End If
        Else
            Err.Clear
        End If
        ' Note: If Version_B fetch fails, we continue with primary XPaths only.
        ' The _Alt variables will remain empty ("") which is handled gracefully by the bot.
        On Error GoTo ErrorHandler
        
        Set httpB = Nothing
        Set jsonB = Nothing
        
        ' Autofit columns for better readability
        ws.Columns("A:D").AutoFit
        
        ' Update the named range for the last retrieved time
        On Error Resume Next
        ws.Range(LAST_XPATH_RETRIEVED).Value = Format(Now, DATE_FORMAT)
        On Error GoTo ErrorHandler
        
        If Err.Number <> 0 Then
            Err.Raise vbObjectError + 1005, "mXpaths.ParseXPathsFromAPI", "The named range '" & LAST_XPATH_RETRIEVED & "' is missing or cannot be updated."
            Exit Function
        End If
        
        ' Success - set return value to True
        ParseXPathsFromAPI = True
    Else
        ' Handle non-200 status codes
        ' Check if it's a server connection error (status codes 500-599)
        If http.status >= 500 And http.status <= 599 Then
            ' This is a server error, handle it silently
            isServerConnectionError = True
        Else
            ' This is another type of error
            Err.Raise vbObjectError + 1006, "mXpaths.ParseXPathsFromAPI", "API Request Error: " & http.status & " - " & http.statusText
            Exit Function
        End If
    End If
    
CleanUp:
    ' Clean up resources
    Set http = Nothing
    Set json = Nothing
    Set ws = Nothing
    
    Exit Function
    
ErrorHandler:
    ' Clean up resources
    Set http = Nothing
    Set json = Nothing
    Set ws = Nothing
    
    ' Let the error bubble up to be handled by the calling procedure
    RaiseError Err.Number, Err.Source, "mXpaths.ParseXPathsFromAPI", Err.Description, Erl
End Function

'------------------------------
' CreateOrUpdateNamedRange creates or updates a workbook-level named range
' pointing to a specific cell. This allows the bot to read XPaths by name.
'------------------------------
Private Sub CreateOrUpdateNamedRange(ByVal rangeName As String, ByVal targetCell As Range)
    On Error Resume Next
    ThisWorkbook.Names(rangeName).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=rangeName, RefersTo:=targetCell
End Sub
