VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If BOT.Name <> "BOT" Then
    BOT.Name = "BOT"
End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    Dim keepFormat As Boolean
    Dim UndoString As String
    Dim UndoControlID As Long
    Dim ctrl As CommandBarControl
    Dim listCount As Long


    ' Check the toggle option.
    On Error Resume Next
    keepFormat = CBool(ThisWorkbook.Names("KeepOriginalFormatting").RefersToRange.Value)
    On Error GoTo 0
    If Not keepFormat Then GoTo CleanExit

    ' Get the Undo control.
    On Error Resume Next
    Set ctrl = Application.CommandBars("Standard").FindControl(ID:=128)
    On Error GoTo 0
    If ctrl Is Nothing Then GoTo CleanExit
    
    ' Safely get ListCount.
    On Error Resume Next
    listCount = ctrl.listCount
    On Error GoTo 0
    If listCount = 0 Then GoTo CleanExit

    ' Safely get the first item in the Undo list.
    On Error Resume Next
    UndoString = ctrl.List(1)
    If Err.Number <> 0 Then GoTo CleanExit
    On Error GoTo 0

    UndoControlID = GetControlID("Standard", "&" & UndoString)
    If UndoControlID <> 6002 Then GoTo CleanExit ' Only proceed if the action was Paste.

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Application.Undo
    PasteSpecialSafe Target

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub PasteSpecialSafe(Target As Range)
    On Error Resume Next
    Target.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    If Err.Number <> 0 Then
        Dim clipText As String
        clipText = Clipboard()
        If clipText <> "" Then
            Dim dataArr As Variant
            dataArr = ParseClipboardData(clipText)
            If IsArray(dataArr) Then
                Target.Resize(UBound(dataArr, 1) - LBound(dataArr, 1) + 1, _
                              UBound(dataArr, 2) - LBound(dataArr, 2) + 1).Value = dataArr
            Else
                Target.Value = clipText
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Private Function ParseClipboardData(ByVal text As String) As Variant
    On Error Resume Next
    Dim rows As Variant, r As Long, c As Long, cols As Variant
    ' Split text into rows.
    rows = Split(text, vbCrLf)
    If UBound(rows) = 0 Then rows = Split(text, vbLf)
    
    Dim maxCols As Long: maxCols = 0
    Dim i As Long
    ' Determine the maximum number of columns.
    For i = LBound(rows) To UBound(rows)
        cols = Split(rows(i), vbTab)
        If UBound(cols) > maxCols Then maxCols = UBound(cols)
    Next i
    
    Dim result() As Variant
    ReDim result(LBound(rows) To UBound(rows), 0 To maxCols)
    
    ' Populate the array.
    For r = LBound(rows) To UBound(rows)
        cols = Split(rows(r), vbTab)
        For c = 0 To maxCols
            If c <= UBound(cols) Then
                result(r, c) = cols(c)
            Else
                result(r, c) = ""
            End If
        Next c
    Next r
    ParseClipboardData = result
End Function

Private Function Clipboard(Optional StoreText As String) As String
    On Error Resume Next
    Dim X As Variant
    X = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(StoreText) > 0
                    .SetData "text", X
                Case Else
                    Clipboard = .GetData("text")
            End Select
        End With
    End With
End Function

Private Function GetControlID(ByVal cbName As String, ByVal sCaption As String) As Long
    On Error Resume Next
    Dim cbBar As CommandBar
    Dim cbCtl As CommandBarControl
    
    Set cbBar = Application.CommandBars(cbName)
    
    For Each cbCtl In cbBar.Controls
        If cbCtl.Caption = sCaption Then
            GetControlID = cbCtl.ID
            Exit Function
        End If
    Next cbCtl
    
    GetControlID = -1
End Function


