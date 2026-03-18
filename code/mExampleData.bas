Attribute VB_Name = "mExampleData"
'---------------------------------------------------------------------------------------
' Procedure : InsertExampleData
' Author    : Sven Bosau
' Website   : https://pythonandvba.com
'
' Purpose   : Overwrites existing data (starting at row 3) in the "BOT" sheet
'             with new example rows from the "Example_Data" sheet (columns A:C).
'             Asks the user to confirm before overwriting.
'             Skips blank rows in the example sheet.
'
' Usage     : Typically invoked via a button on the "BOT" sheet.
'             The "Example_Data" sheet can be hidden or very hidden.
'---------------------------------------------------------------------------------------
Option Explicit

Public Sub InsertExampleData()
    On Error GoTo ErrHandler
    
    ' =============== DECLARATIONS ===============
    Dim wsBot As Worksheet
    Dim wsExample As Worksheet
    Dim lastRowBot As Long
    Dim lastRowExample As Long
    Dim rngSource As Range
    Dim response As VbMsgBoxResult
    
    ' =============== STEP 1: REFERENCE SHEETS ===============
    ' The sheet where data will be inserted:
    Set wsBot = ThisWorkbook.Worksheets("BOT")
    
    ' The sheet that holds the example rows (can be hidden):
    Set wsExample = ThisWorkbook.Worksheets("Example_Data")
    
    ' =============== STEP 2: CHECK IF BOT ALREADY HAS DATA ===============
    ' Look at the "Number" column (wcNumber = Column C) to find the last used row.
    lastRowBot = wsBot.Cells(wsBot.rows.Count, BotColumn.wcNumber).End(xlUp).Row
    
    If lastRowBot >= FirstRow Then
        ' Prompt user about overwriting
        response = MsgBox( _
            "This will remove all existing data starting at row " & FirstRow & _
            " and replace it with example rows. Continue?", _
            vbYesNo + vbQuestion, _
            "Insert Example Data" _
            )
        If response = vbNo Then Exit Sub
    End If
    
    ' =============== STEP 3: CLEAR OLD DATA (COLUMNS B:D) ===============
    If lastRowBot >= FirstRow Then
        wsBot.Range( _
            wsBot.Cells(FirstRow, BotColumn.wcNumber), _
            wsBot.Cells(lastRowBot, BotColumn.wcStatus) _
            ).ClearContents
    End If
    
    ' =============== STEP 4: FIND EXAMPLE ROWS IN "Example_Data" ===============
    ' We assume row 2 onward are data rows in column A.
    lastRowExample = wsExample.Cells(wsExample.rows.Count, "A").End(xlUp).Row
    If lastRowExample < 2 Then
        MsgBox "No example rows found in 'Example_Data'! Check that row 2 and onward contain data.", _
            vbInformation, "Insert Example Data"
        Exit Sub
    End If
    
    ' =============== STEP 5: DEFINE THE SOURCE RANGE (A2:C...) ===============
    Set rngSource = wsExample.Range("A2:C" & lastRowExample)
    
    ' =============== STEP 6: COPY ROW BY ROW INTO BOT ===============
    Dim rowIndexBot As Long: rowIndexBot = FirstRow
    Dim rowIndexSrc As Long
    
    For rowIndexSrc = 1 To rngSource.rows.Count
        
        ' Column A -> wcNumber (col B)
        wsBot.Cells(rowIndexBot, BotColumn.wcNumber).Value = _
            rngSource.Cells(rowIndexSrc, 1).Value
        
        ' Column B -> wcText (col C)
        wsBot.Cells(rowIndexBot, BotColumn.wcText).Value = _
            rngSource.Cells(rowIndexSrc, 2).Value
        
        rowIndexBot = rowIndexBot + 1
        
    Next rowIndexSrc
    
    ' =============== STEP 7: DONE, INFORM USER ===============
    MsgBox "Example rows have been inserted successfully." & vbCrLf & _
        "You can now review or send them.", vbInformation, "Insert Example Data"
    
    Exit Sub

    ' ================= ERROR HANDLER =================
ErrHandler:
    DisplayError Err.Source, Err.Description, "mExampleData.InsertExampleData", Erl

End Sub

