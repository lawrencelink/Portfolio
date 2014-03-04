Private Sub Import_II_Data_2014()
        Application.Calculation = xlCalculationManual
        MsgBox "Please open the 'II Weekly Utilization w gtd 2014 MASTER'"
        Call SUPPLY
        
        Call Save_File
        'Save the new files
        Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub SUPPLY()
    
        Dim strTodaysDate As Date
        Dim intCountofRows As Integer
        Dim strFilename As String
        Dim strWorkbookName As String
        Dim i As Long, lngAreas, lngRows As Long
        Dim strSaveAsName As String
        Dim LDate As Date
        
        'SUPPLY
        
        Windows("II Weekly Utilization w gtd 2014 MASTER").Activate
        
        'Ask user to open the supply file
        MsgBox "Please open the data file"
        strFilename = Application.GetOpenFilename
        strWorkbookName = Right(strFilename, 40)
        'MsgBox Len(strWorkbookName)
        If strFilename = "False" Then
            MsgBox "Stopping because you did not select a file"
            Exit Sub
        Else
            Workbooks.OpenText Filename:=strFilename, DataType:=xlDelimited, Space:=True, ConsecutiveDelimiter:=True
        End If
        
        '2014
        'Go back to the data file
        'Windows(strWorkbookName).Activate
        
        Sheets("Supply 2014").Select
        Range("A1").Select
        Range(Selection, Selection.End(xlDown)).Select
        lngAreas = Selection.Areas.Count
     
        lngRows = 0
         
        For i = 1 To lngAreas
            lngRows = lngRows + Selection.Areas(i).Rows.Count
        Next i
        
        'Format the column to date
        Range("E1:E" & lngRows).Select
        'Columns("E:E").Select
        Selection.TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 3), TrailingMinusNumbers:=True
        
        'Select the data
        Range("A1:K1").Select
        Range(Selection, Selection.End(xlDown)).Select
        
        'copy the data
        Selection.Copy
        
        'Select the other workbook
        Windows("II Weekly Utilization w gtd 2014 MASTER").Activate
        
        Sheets("Total Supply").Visible = True
        Sheets("Total Supply").Select
        Range("A1").Select
        
        'Paste the data
        ActiveSheet.Paste
        
        'Fill in the formula
        'Fill in the formula
        Range("L2:N2").Select
        Selection.AutoFill Destination:=Range("L2", "N" & lngRows), Type:=xlFillDefault
        
        Sheets("Total Supply").Visible = False
        
        Call CONFIRMED(strWorkbookName)
        'Call CLOSE_FILE(strWorkbookName)
        
        'Application.Calculate
End Sub

Private Sub CONFIRMED(strWorkbookName)
        Dim strTodaysDate As Date
        Dim intCountofRows As Integer
        Dim strFilename As String
        'Dim strWorkbookName As String
        Dim i As Long, lngAreas, lngRows As Long
        Dim strSaveAsName As String
        Dim LDate As Date
        
        'CONFIRMED
        'Select the MASTER workbook
        Windows("II Weekly Utilization w gtd 2014 MASTER").Activate
        'Ask user to open the supply file
        MsgBox "Please open the confirmed file"
        strFilename = Application.GetOpenFilename
        strWorkbookName = Right(strFilename, 43)
        'MsgBox Len(strWorkbookName)
        If strFilename = "False" Then
            MsgBox "Stopping because you did not select a file"
            Exit Sub
        Else
            Workbooks.OpenText Filename:=strFilename, DataType:=xlDelimited, Space:=True, ConsecutiveDelimiter:=True
        End If
        
        '2014
        'Go back to the data file
        Windows(strWorkbookName).Activate
        
        Sheets("Confirmed 2014").Select
        Range("A1").Select
        Range(Selection, Selection.End(xlDown)).Select
        lngAreas = Selection.Areas.Count
     
        lngRows = 0
         
        For i = 1 To lngAreas
            lngRows = lngRows + Selection.Areas(i).Rows.Count
        Next i
        
        'Format the column to date
        Range("B1:B" & lngRows).Select
        Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 3), TrailingMinusNumbers:=True
        
        Range("C1:C" & lngRows).Select
        Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 3), TrailingMinusNumbers:=True
        
        'Select the data
        Range("B1:K1").Select
        Range(Selection, Selection.End(xlDown)).Select
        
        'copy the data
        Selection.Copy
        
        'Select the other workbook
        Windows("II Weekly Utilization w gtd 2014 MASTER").Activate
        
        Sheets("Total Confirmed").Visible = True
        Sheets("Total Confirmed").Select
        Range("A1").Select
        
        'Paste the data
        ActiveSheet.Paste
        
        'Fill in the formula
        'Fill in the formula
        Range("K2:M4").Select
        Selection.AutoFill Destination:=Range("K2", "M" & lngRows), Type:=xlFillDefault
        
        Sheets("Total Confirmed").Visible = False
        
        Call CLOSE_FILE(strWorkbookName)
        
        Application.Calculate
End Sub

Private Sub CLOSE_FILE(strWorkbookName)
    Windows(strWorkbookName).Activate
    
    If Len(strWorkbookName) = 48 Then
    
    
    Application.DisplayAlerts = False
    Windows(strWorkbookName).Close
    Application.DisplayAlerts = True
    End If

    If Len(strWorkbookName) = 51 Then
        
        Application.DisplayAlerts = False
        Windows(strWorkbookName).Close
        Application.DisplayAlerts = True
    End If
    If Len(strWorkbookName) = 40 Then
        
        Application.DisplayAlerts = False
        Windows(strWorkbookName).Close
        Application.DisplayAlerts = True
    End If
    If Len(strWorkbookName) = 43 Then
        
        Application.DisplayAlerts = False
        Windows(strWorkbookName).Close
        Application.DisplayAlerts = True
    End If
    If Len(strWorkbookName) = 33 Then
        
        Application.DisplayAlerts = False
        Windows(strWorkbookName).Close
        Application.DisplayAlerts = True
    End If

End Sub

Private Sub Save_File()
    
    'save the 2014 file
    Windows("II Weekly Utilization w gtd 2014 MASTER").Activate
    ChDir _
        "F:\II Utilization Reports\2014 Utilization"
        LDate = Date
        strSaveAsName = "F:\II Utilization Reports\2014 Utilization\II Weekly Utilization w gtd 2014 " & Format(LDate, "mm.dd.yy") & ".xlsx"
        'MsgBox strSaveAsName
        ActiveWorkbook.SaveAs Filename:=strSaveAsName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
End Sub

Private Sub Enter_the_Date()
    Dim LDate As Date
    Range("D3").Select
    LDate = Date
    FormulaR1C1 = "Updated" & LDate
End Sub