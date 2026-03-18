Attribute VB_Name = "mUpdateHandler"
' ------------------------------------------------------
' Name:      mUpdateHandler
' Author:    Sven Bosau
' Website:   https://pythonandvba.com
' YouTube:   https://youtube.com/@CodingIsFun
' Email:     support@pythonandvba.com
' Date:      9/25/2021
' Purpose:   Check if new file is available to download
' ------------------------------------------------------
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As LongPtr
#Else
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
    "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
    szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If
' ------------------------------------------------------
' Name:      CheckForUpdates
' Purpose:   The idea is, that the input variables are stored as Global variables,e.g. in 'mGlobals' Module
'
' Input:
'   VersionControlfile: This textfile is stored on a server and holds the current version number, e.g. "1.2"
'   LatestFileUrl: The latest version/file, which will be downloaded
'   CurrentVersion: Should written as a string, examples: "5.1", "0.9", ..'
'
' Example Input:
'   VersionControlfile = "https://raw.githubusercontent.com/Sven-Bo/excel-whatsapp-sender/master/VersionControl.txt"
'   LatestFileUrl= "https://github.com/Sven-Bo/excel-whatsapp-sender/raw/master/WhatsApp_BOT_BASIC.xlsm"
'   CurrentVersion = "5.1"
' ------------------------------------------------------
Public Const VersionControlfile As String = "https://raw.githubusercontent.com/Sven-Bo/excel-whatsapp-sender/master/VersionControl.txt"
Public Const VersionHistoryfile As String = "https://raw.githubusercontent.com/Sven-Bo/excel-whatsapp-sender/master/VersionHistory.txt"
Public Const LatestFileUrl As String = "https://github.com/Sven-Bo/excel-whatsapp-sender/raw/master/WhatsApp_BOT_BASIC.xlsm"

Public Sub CheckForUpdates()
                           
    Dim IsUpdateAvailable As Boolean
    Dim LatestVersion As String
    Dim ContentURLFile As String
    Dim VersionHistory As String
    
    'Read the textfile from the server
    'Textfile should be formatted as follows: "1.2 | YYYY/MM/DD Some Comments"
    'In case of an error it will return an empty string -> exit sub
    ContentURLFile = ReadURLFile(VersionControlfile)
    If ContentURLFile = vbNullString Then Exit Sub
    
    'Get only the version number
    'In case of an error it will return an empty string -> exit sub
    'Example Input "1.2 | YYYY/MM/DD Some Comments"
    'Example Output: "1.2"
    LatestVersion = GetVersionNumber(ContentURLFile)
    If LatestVersion = vbNullString Then Exit Sub

    VersionHistory = ReadURLFile(VersionHistoryfile)
    If VersionHistory = vbNullString Then Exit Sub

    IsUpdateAvailable = VersionCheck(VERSION_NUMBER, LatestVersion)
    If IsUpdateAvailable Then
        Select Case MsgBox("There is an update available. You are using version: " & VERSION_NUMBER & vbCrLf & vbCrLf & _
            "Would you like to download the latest version?" & vbCrLf & vbCrLf & _
            "Version History: " & vbCrLf & _
            VersionHistory, _
            vbYesNo, "Check for updates")
            Case vbYes
                Call DownloadFile(LatestFileUrl, LatestVersion)
            Case vbNo
                Exit Sub
        End Select
    Else
        MsgBox _
            "No update available. You are using the latest version [version: " & VERSION_NUMBER & "]" & vbCrLf & vbCrLf & _
            "Version History: " & vbCrLf & _
            VersionHistory, _
            vbOKOnly, "Check for updates"
    End If

End Sub

' ------------------------------------------------------
' Name:      DownloadFile
' Purpose:   Download a file from a server by using the Windows API
' ------------------------------------------------------
Private Sub DownloadFile(ByVal FileUrl As String, ByVal LatestVersion As String)

    On Error GoTo ErrorHandler
    Dim downloadStatus As Variant
    Dim destinationFile_local As String
    Dim fileName As String
    
    'Return only the file name
    'Example Input: https://server.com/example.csv, LatestVersion = "1.2"
    'Example Output: example_1.2.csv
    fileName = Mid(FileUrl, InStrRev(FileUrl, "/") + 1)
    fileName = Split(fileName, ".")(0) & "_" & LatestVersion & "." & Split(fileName, ".")(1)
    
    destinationFile_local = Application.ActiveWorkbook.path & "\" & fileName
    downloadStatus = URLDownloadToFile(0, FileUrl, destinationFile_local, 0, 0)
    
    If downloadStatus = 0 Then
        MsgBox _
            "Downloaded Successfully!" & vbCrLf & vbCrLf & _
            "Please find the file here:" & vbCrLf & _
            destinationFile_local, vbOKOnly, "Download File"
    Else
        MsgBox _
            "Downloaded failed!" & vbCrLf & vbCrLf & _
            "Please try to manually download the file:" & vbCrLf & _
            FileUrl, vbOKOnly, "Download File"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox _
        "The following error has occurred." & vbCrLf & vbCrLf & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Error Source: DownloadFile" & vbCrLf & _
        "Error Description: " & Err.Description, _
        vbCritical, "An Error has Occurred!"
End Sub
' ------------------------------------------------------
' Name:      ReadURLFile
' Purpose:   Read the content from a (text)file, which is stored on a server
' Returns:   VbNullString if an error is reported back from the server
'            otherwise returns the Response Text from the server
' ------------------------------------------------------
Private Function ReadURLFile(ByVal sFullURLWFile As String) As String

    On Error GoTo ErrorHandler
    Dim oHttp As Object
 
    Set oHttp = CreateObject("Microsoft.XMLHTTP")
 
    oHttp.Open "GET", sFullURLWFile, False
    oHttp.Send
    
    'Check for any errors reported by the server
    If oHttp.status >= 400 And oHttp.status <= 599 Then
        GoTo ErrorHandler
    Else
        ReadURLFile = oHttp.responseText
    End If
 
EndIt:
    Set oHttp = Nothing
    Exit Function
 
ErrorHandler:
    MsgBox _
        "The following error has occurred." & vbCrLf & vbCrLf & _
        "Module Name: ReadURLFile" & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Error Source: ReadURLFile" & vbCrLf & _
        "Error Description: " & Err.Description, _
        vbCritical, "An Error has Occurred!"
    ReadURLFile = ""
    GoTo EndIt
End Function
' ------------------------------------------------------
' Name:      GetVersionNumber
' Purpose:   Input Example: "1.2 | YYYY/MM/DD Some Comments"
' Returns:   VbNullString if an error occurs
'            Otherwise it will return the version number, e.g. "1.2"
' ------------------------------------------------------
Private Function GetVersionNumber(ByVal content As String) As String
    Dim VersionNumber As String
    On Error GoTo ErrorHandler
    VersionNumber = Split(content, "|")(0)
    VersionNumber = Replace(VersionNumber, " ", "")
    GetVersionNumber = VersionNumber

    On Error GoTo InvalidVersionNumber
    VersionNumber = CInt(VersionNumber)
    Exit Function

InvalidVersionNumber:
    MsgBox _
        "The following error has occurred." & vbCrLf & vbCrLf & _
        VersionNumber & " | is not a valid version number!" & vbCrLf & vbCrLf & _
        "Module Name: GetVersionNumber", vbOKOnly, "GetVersionNumber"
    GetVersionNumber = ""
    Exit Function

ErrorHandler:
    MsgBox _
        "The following error has occurred." & vbCrLf & vbCrLf & _
        "Module Name: GetVersionNumber" & _
        "Error Number: " & Err.Number & vbCrLf & _
        "Error Source: GetVersionNumber" & vbCrLf & _
        "Error Description: " & Err.Description, _
        vbCritical, "An Error has Occurred!"
    GetVersionNumber = ""
    Exit Function

End Function
' ------------------------------------------------------
' Name:      VersionCheck
' Purpose:   Compare two version numbers, e.g. "1.2" with "1.3"
' Returns:   False, if current version >= latest version
'            True, if current version < latest version
'            False, if an error occurs
' ------------------------------------------------------
Private Function VersionCheck(ByVal currVer As String, ByVal latestVer As String) As Boolean
    
    On Error GoTo ErrorHandler
    Dim currArr() As String
    Dim latestArr() As String
    currArr = Split(currVer, ".")
    latestArr = Split(latestVer, ".")
    
    'If versions are the same return False
    If currVer = latestVer Then
        VersionCheck = False
        Exit Function
    End If
    
    'Iterate through the version components
    Dim i As Integer
    For i = LBound(currArr) To UBound(currArr)
    
        'If the end of the latest cersion is reached, the current version must be up to greater
        'meaning it is up to date
        If i > UBound(latestArr) Then
            VersionCheck = False
            Exit Function
        End If
    
        'Cast the component to an integer
        Dim curr As Integer, latest As Integer
        curr = Int(currArr(i))
        latest = Int(latestArr(i))
    
        'Check which version component is greater in which case return a result
        If curr > latest Then
            VersionCheck = False
            Exit Function
        ElseIf curr < latest Then
            VersionCheck = True
            Exit Function
        End If
    
        'If the version components are equal, iterate to the next component
    Next
    
    'If there are remaining components in the latest version, return true
    If i < UBound(latestArr) Then
        VersionCheck = True
        Exit Function
    End If
    
ErrorHandler:
    VersionCheck = False
    Exit Function

End Function
' ------------------------------------------------------
' Name:      AddInfoButton
' Purpose:   Insert a 'Info Button' (shape with question mark) in selection.
'            Clicking the shape will trigger a MsgBox
' ------------------------------------------------------
Sub AddInfoButton()

    Dim clLeft As Double
    Dim clTop As Double
    Dim clLeftUpdate As Double
    Dim clTopUpdate As Double
    Dim cl As Range
    Dim shpInfoBtn As Shape
    Dim shpUpdateBtn As Shape
    Dim customShapes As New Collection
    Dim customShape As Variant

    Set cl = Range(Selection.Address)

    clLeft = cl.Left
    clTop = cl.Top
    
    clLeftUpdate = cl.Offset(0, -3).Left
    clTopUpdate = cl.Offset(0, -3).Top

    Set shpInfoBtn = ActiveSheet.Shapes.AddShape(msoShapeOval, clLeft, clTop, 30, 30)
    Set shpUpdateBtn = ActiveSheet.Shapes.AddShape(msoShapeRectangle, clLeftUpdate, clTopUpdate, 200, 30)

    'Insert Info Button
    With shpInfoBtn
        .Name = "info-button"
        .TextFrame2.TextRange.Characters = "?"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .ShapeStyle = msoShapeStylePreset11
        .OnAction = "'" & ActiveWorkbook.Name & "'!InfoBtnText"
    End With

    'Insert Check Updates Button
    With shpUpdateBtn
        .Name = "update-button"
        .TextFrame2.TextRange.Characters = "Check for Updates..."
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .ShapeStyle = msoShapeStylePreset11
        .OnAction = "'" & ActiveWorkbook.Name & "'!CheckForUpdates"
    End With
    
    customShapes.Add shpInfoBtn, "InfoBtn"
    customShapes.Add shpUpdateBtn, "UpdateBtn"
    
    'Style text within shapes
    For Each customShape In customShapes
        With customShape.TextFrame2.TextRange.Characters
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
            .Font.Fill.Visible = msoTrue
            .Font.size = 16
            .Font.Name = "Arial Black"
        End With
    Next customShape
    
    With customShapes("UpdateBtn").TextFrame2.TextRange.Characters
        .ParagraphFormat.Alignment = msoAlignLeft
        .Font.size = 9
    End With
    
End Sub
Sub InfoBtnText()
    MsgBox _
        "'Check for updates...' will check if you are using the latest version of this file." & vbCrLf _
        & "If there is an update available, you can download it directly from the server.", _
        vbOKOnly, "Check for updates..."
End Sub

