VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Settings.Name <> "Settings" Then
    Settings.Name = "Settings"
End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error Resume Next ' Prevent unexpected crashes during execution

    Dim DefaultChromePath As String
    Dim ws As Worksheet
    Dim useDefaultRange As Range
    Dim customPathRange As Range
    Dim sendingMethodRange As Range
    Dim selectedValue As String
    Dim Username As String

    ' Get the current Windows username dynamically
    Username = Environ("USERNAME")

    ' Define the default Chrome path with the actual username
    DefaultChromePath = "C:\Users\" & Username & "\AppData\Local\SeleniumBasic\chrome-win64\chrome-win64\chrome.exe"

    ' Set references
    Set ws = Me ' Assuming this is on the settings worksheet
    Set useDefaultRange = ws.Range("UseDefaultChromeBinary") ' Named range for UseDefaultChromeBinary
    Set customPathRange = ws.Range("CustomChromeBinaryPath") ' Named range for CustomChromeBinaryPath
    Set sendingMethodRange = ws.Range("SendingMethod") ' Named range for SendingMethod

    ' Handle changes for UseDefaultChromeBinary
    If Not Intersect(Target, useDefaultRange) Is Nothing Then
        Application.EnableEvents = False ' Prevent recursive events

        ' If "No," populate default path in CustomChromeBinaryPath (if empty)
        If LCase(useDefaultRange.Value) = "no" Then
            If Len(Trim(customPathRange.Value)) = 0 Then
                customPathRange.Value = DefaultChromePath
            End If

        ' If "Yes," clear CustomChromeBinaryPath but only if it matches the default
        ElseIf LCase(useDefaultRange.Value) = "yes" Then
            If customPathRange.Value = DefaultChromePath Then
                customPathRange.ClearContents
            End If
        End If

        Application.EnableEvents = True
    End If

    ' Handle changes for SendingMethod
    If Not Intersect(Target, sendingMethodRange) Is Nothing Then
        Application.EnableEvents = False ' Prevent recursive events

        ' Get the selected value
        selectedValue = LCase(Trim(sendingMethodRange.Value))

        ' Show a hint based on the selected method
        Select Case selectedValue
            Case "default"
                MsgBox "Default mode supports special characters and line breaks. It's also faster but stops working if your screen is locked (e.g., after pressing Windows + L).", _
                       vbInformation, "Sending Method"
            Case "screen lock compatible"
                MsgBox "Screen Lock Compatible mode works even if your screen is locked (e.g., after pressing Windows + L). However, it does not support some special characters and may be slower.", _
                       vbInformation, "Sending Method"
        End Select

        Application.EnableEvents = True
    End If

    On Error GoTo 0 ' Reset error handling to default after execution

End Sub








