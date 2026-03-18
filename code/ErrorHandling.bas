Attribute VB_Name = "ErrorHandling"
'---------------------------------------------------------------------------------------
' Module    : ErrorHandling
' Purpose   : Error Handling traceback for subs/func withtin one module
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------
' Error Handling Module Configuration
'---------------------------------------------------------------------
' Webhook Configuration
Private Const WEBHOOK_URL As String = "https://connect.pabbly.com/workflow/sendwebhookdata/IjU3NjYwNTY5MDYzMjA0MzM1MjY0NTUzNjUxMzci_pc"
Private Const APP_NAME As String = "WhatsApp Blaster FREE Version"

' Global variable to store additional information for error reporting
Public AdditionalErrorInfo As String

Private Const LINE_NO_TEXT As String = "Line Number: "
' Used to prevent multiple error messages
Dim AlreadyUsed As Boolean

' Reraises an error and adds line number and current procedure name
Sub RaiseError(ByVal errorNo As Long _
                , ByVal src As String _
                , ByVal proc As String _
                , ByVal desc As String _
                , ByVal lineNo As Long)

    Dim sSource As String

    ' If called for the first time then add line number
    If AlreadyUsed = False Then
        
        ' Add error line number if present
        If lineNo <> 0 Then
            sSource = vbNewLine & LINE_NO_TEXT & lineNo & " "
        End If

        ' Add procedure to source
        sSource = sSource & vbNewLine & proc
        AlreadyUsed = True
        
    Else
        ' If error has already been raised simply add on procedure name
        sSource = src & vbNewLine & proc
    End If
    
    ' Pause the code here when debugging
    '(To Debug: "Tools->VBA Properties" from the menu.
    ' Add "Debugging=1" to the     ' "Conditional Compilation Arguments.)
#If Debugging = 1 Then
    Debug.Assert False
#End If

    ' Reraise the error so it will be caught in the caller procedure
    ' (Note: If the code stops here, make sure DisplayError has been
    ' placed in the topmost procedure)
    If errorNo <= 0 Then
        errorNo = 1000 ' Default error code
    End If

    Err.Raise errorNo, sSource, desc

End Sub

' Displays the error when it reaches the topmost sub
' Note: You can add a call to logging from this sub
Sub DisplayError(ByVal src As String, ByVal desc As String _
                    , ByVal sProcname As String, lineNo As Long)

    ' Check If the error happens in topmost sub
    If AlreadyUsed = False Then
        ' Reset string to remove "VBAProject" and add line number if it exists
        src = IIf(lineNo = 0, "", LINE_NO_TEXT & lineNo)
    End If
    
    ' Store original error information
    Dim origErrNumber As Long
    Dim origErrDesc As String
    origErrNumber = Err.Number
    origErrDesc = Err.Description
    
    ' Send webhook notification for all errors
    On Error Resume Next
    SendErrorWebhook src, sProcname, origErrNumber, origErrDesc
    On Error GoTo 0

    ' Build the final message
    Dim sMsg As String
    sMsg = "System Information: " & vbNewLine & _
            "Microsoft Excel version " & Application.Version & _
            " running on " & Application.OperatingSystem
    
    ' Add additional information if available
    If Len(AdditionalErrorInfo) > 0 Then
        sMsg = sMsg & vbNewLine & AdditionalErrorInfo
    End If
    
    sMsg = sMsg & vbNewLine & vbNewLine & "The following error occurred: " & vbNewLine & origErrNumber & ": " & origErrDesc _
           & vbNewLine & vbNewLine & "Error Location is: "
    sMsg = sMsg & src & vbNewLine & sProcname & vbNewLine & vbNewLine & _
           "Version Number: " & VERSION_NUMBER & vbNewLine & _
           "Make sure you are using the latest version by clicking on 'Check for Updates'" & vbNewLine & vbNewLine & _
           "If you are using the latest version but still getting an error, please note the above details before contacting support: " & ERROR_EMAIL

    ' Display the message
    MsgBox sMsg, title:="Error"

    ' reset the boolean value
    AlreadyUsed = False

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SendErrorWebhook
' Purpose   : Sends error information to the webhook for all errors
'---------------------------------------------------------------------------------------
Private Sub SendErrorWebhook(ByVal src As String, ByVal sProcname As String, _
                            ByVal errorNumber As Long, ByVal errorDesc As String)
    ' All errors are suppressed
    On Error Resume Next

    Dim Username As String
#If Mac Then
   Username = Environ("USER")
#Else
   Username = Environ("UserName")
#End If
    
    ' Create a simple JSON string using string concatenation
    Dim payload As String
    payload = "{"
    
    ' App information
    payload = payload & """app_name"": " & JsonConverter.ConvertToJson(APP_NAME) & ","
    payload = payload & """app_version"": " & JsonConverter.ConvertToJson(VERSION_NUMBER) & ","
    payload = payload & """timestamp"": " & JsonConverter.ConvertToJson(Format(Now, "yyyy-mm-dd hh:mm:ss")) & ","
    payload = payload & """os"": " & JsonConverter.ConvertToJson(Application.OperatingSystem) & ","
    payload = payload & """user_name"": " & JsonConverter.ConvertToJson(Username) & ","
    
    ' Error information - exactly as displayed in the message box
    payload = payload & """error_number"": " & errorNumber & ","
    payload = payload & """error_description"": " & JsonConverter.ConvertToJson(errorDesc) & ","
    
    ' Include any additional error information with the procedure name
    Dim fullErrorLocation As String
    fullErrorLocation = src & vbNewLine & sProcname
    
    If Len(AdditionalErrorInfo) > 0 Then
        fullErrorLocation = fullErrorLocation & vbNewLine & "Additional Info: " & AdditionalErrorInfo
    End If
    
    payload = payload & """error_location"": " & JsonConverter.ConvertToJson(fullErrorLocation)
    
    payload = payload & "}"
    
    ' Use WinHttpRequest to send the webhook
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Configure the request
    http.Open "POST", WEBHOOK_URL, False
    http.SetRequestHeader "Content-Type", "application/json"
    
    ' Send the request
    http.Send payload
    
    ' No error handling or response checking
    Set http = Nothing
    
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetAdditionalErrorInfo
' Purpose   : Sets additional information to be included in error reports
'---------------------------------------------------------------------------------------
Public Sub SetAdditionalErrorInfo(ByVal infoText As String)
    AdditionalErrorInfo = infoText
End Sub

'---------------------------------------------------------------------------------------
' Procedure : AppendAdditionalErrorInfo
' Purpose   : Appends information to the existing additional error info
'---------------------------------------------------------------------------------------
Public Sub AppendAdditionalErrorInfo(ByVal infoText As String)
    If Len(AdditionalErrorInfo) > 0 Then
        AdditionalErrorInfo = AdditionalErrorInfo & vbNewLine & infoText
    Else
        AdditionalErrorInfo = infoText
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ClearAdditionalErrorInfo
' Purpose   : Clears any additional error information
'---------------------------------------------------------------------------------------
Public Sub ClearAdditionalErrorInfo()
    AdditionalErrorInfo = ""
End Sub


