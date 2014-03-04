Public Sub RCI_2014()
       
    Dim strPage As String
    Dim Sht As Worksheet
    Dim strCurrentSheet As String
    Dim strRCIDownloadedSheet As String
    Dim sSheets() As Variant
    Dim sSheets2() As Variant
    Dim i As Integer
    Dim LDate As Date
    Dim strSaveAsName As String
    Dim iRet As Integer
    Dim strPrompt As String
    Dim strTitle As String
    Dim blnBoolean As Boolean
    Dim strFilename As String
    
    MsgBox "Please open the data file"
    strFilename = Application.GetOpenFilename
    strRCIDownloadedSheet = Right(strFilename, 45)
        
    If strFilename = "False" Then
        MsgBox "Stopping because you did not select a file"
        Exit Sub
    Else
        Workbooks.OpenText Filename:=strFilename, DataType:=xlDelimited, Space:=True, ConsecutiveDelimiter:=True
    End If
        
    LDate = Date
    
    strCurrentSheet = "2014 RCI Utilization " & Format(LDate, "mm.dd.yyyy") & ".xlsx"
    
    sSheets = Array("SBP", "SDO", "SDO-49", "SVR", "SVR-FTN", "VBC")
    sSheets2 = Array("Page1_7", "Page1_2", "Page1_3", "Page1_4", "Page1_5", "Page1_6")
        
    For i = 0 To 5
            
            Windows(strCurrentSheet).Activate
            Sheets(sSheets(i)).Select
            
            Range("A4:J4").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            
    Next
    
    For i = 0 To 5
            Windows(strRCIDownloadedSheet).Activate
            Sheets(sSheets2(i)).Select
            strPage = Range("A4").Value
            Range("A4:J4").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Select Case strPage
                
                Case "SHERATON BROADWAY PLANTATION"
                    Windows(strCurrentSheet).Activate
                    Sheets("SBP").Select
                    Range("A4").Select
                    ActiveSheet.Paste
                
                Case "SHERATON DESERT OASIS"
                    Windows(strCurrentSheet).Activate
                    Sheets("SDO").Select
                    Range("A4").Select
                    ActiveSheet.Paste
                    
                Case "SHERATON DESERT OASIS II"
                    Windows(strCurrentSheet).Activate
                    Sheets("SDO-49").Select
                    Range("A4").Select
                    ActiveSheet.Paste
                    
                Case "SHERATON DESERT OASIS II"
                    Windows(strCurrentSheet).Activate
                    Sheets("SDO-49").Select
                    Range("A4").Select
                    ActiveSheet.Paste
                    
                Case "SHERATON VISTANA RESORT"
                    Windows(strCurrentSheet).Activate
                    Sheets("SVR").Select
                    Range("A4").Select
                    ActiveSheet.Paste
                    
                Case "SHERATON VISTANA RESORT-FOUNTAINS"
                    Windows(strCurrentSheet).Activate
                    Sheets("SVR-FTN").Select
                    Range("A4").Select
                    ActiveSheet.Paste
                    
                Case "VISTANA'S BEACH CLUB"
                    Windows(strCurrentSheet).Activate
                    Sheets("VBC").Select
                    Range("A4").Select
                    ActiveSheet.Paste
            End Select
    Next
    
    For i = 0 To 5
            Windows(strCurrentSheet).Activate
            
            Sheets(sSheets(i)).Select
            Range("A4").Select
            Selection.End(xlDown).Select
            ActiveCell.Offset(0, 1).Select
            Selection.Clear
    Next
    
    Sheets("SUMMARY").Select
    Range("B2").Select
    Range("B2").Value = LDate
    
    Sheets("CHANGE FROM PRIOR WEEK").Select
    Range("B2").Select
    Range("B2").Value = LDate
    
    Windows(strRCIDownloadedSheet).Close
    'Windows(strCurrentSheet).Close
End Sub