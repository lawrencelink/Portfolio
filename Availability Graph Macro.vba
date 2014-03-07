Option Explicit

'*********************************************************************************************
'This macro imports the SUN Mixed Avail Report from the Availability folder
'It opens the file, copies and pastes the data into the SUN DATA tab
'*********************************************************************************************
Public Sub Open_SUN_File()
    Dim strWorkBookName As String
    strWorkBookName = ActiveWorkbook.Name
    Dim intResponse As Integer
    Dim strFileName As String
    Dim strFileClose
    
    Application.DisplayAlerts = False
    
    'Open the SUN Mixed Avail Report
    
    ChDrive "N:\av20\exchcomp\availability"
    
    'Show the open dialog and pass the selected _
    'file name to the String variable "strFileName
    strFileName = Application.GetOpenFilename
    
    'If they have cancelled
    
    If strFileName = "False" Then
    
        MsgBox "Stopping because you did not select a file"
    
        Exit Sub
    
    Else
    
        Workbooks.OpenText Filename:=strFileName, DataType:=xlDelimited, Space:=True, ConsecutiveDelimiter:=True
    
    End If
    
    strFileClose = Right(strFileName, 25)
    
    'Remove old data
    Windows(strWorkBookName).Activate
    Sheets("SUN DATA").Visible = True
    Sheets("SUN DATA").Select
    Range("A:L").Select
    Selection.Clear
    Range("N2:N16").Select
    Selection.Clear
    Windows(strFileClose).Activate
    
    Call SUN_EXTRACTOR(strWorkBookName)
    
    Call Update_Room_Types
    
    Call Replace_Room_Types
    
    Windows(strFileClose).Close
    
    Application.DisplayAlerts = True
    
    Windows(strWorkBookName).Activate
    
    Application.Calculate
    
End Sub

Private Sub SUN_EXTRACTOR(strWorkBookName As String)
    
    'REMOVE THE '
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Selection.Replace What:="'", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    'PASTE THE DATA INTO THE "SUN DATA SHEET" TAB
    Range("A:L").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    'Move the data
    Windows(strWorkBookName).Activate
    
    Sheets("SUN DATA").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Sheets("SUN DATA").Visible = False
    Sheets("GRAPH").Select

    'Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub Update_Room_Types()

    'Update the Room Types
    Dim CELL As Variant
    Dim rng As Range
    Dim strProperty As String
    Dim intRow As Integer
    strProperty = Sheets("SUN DATA").Range("A3").Value
    intRow = 2
    
    For Each CELL In Worksheets("REFERENCE TABLE").Range("H1:H113").Columns(1).Cells
        
        If CELL.Value = strProperty Then
            
            'Update the Property Number on the INFO tab
            Sheets("SUN DATA").Range("N" & intRow).Value = CELL.Offset(0, 3).Value
            intRow = intRow + 1
        End If
        
    Next

End Sub

Private Sub Replace_Room_Types()

'Update the Room Types
    Dim CELL As Variant
    Dim rng As Range
    Dim strOverallRoomType As Variant
    Dim strSUNRoomType As Variant
    Dim intRow As Integer
    Dim strProperty As String
    Dim intArrayCounter
    Dim intOverallCounter
    Dim intNumberofSUNRooms
    
    strSUNRoomType = Sheets("REFERENCE TABLE").Range("N1:O10").Value
    strProperty = Sheets("SUN DATA").Range("A3").Value
    'strOverallRoomType = Sheets("SUN DATA").Range("N2:N7").Value
    
    intArrayCounter = 1
    intOverallCounter = 1
    intRow = 2
    'Place the SUN room types into the strSUNRoomType array
    For Each CELL In Worksheets("REFERENCE TABLE").Range("A1:A113").Columns(1).Cells
        
        If CELL.Value = strProperty Then
            
            strSUNRoomType(intArrayCounter, 1) = CELL.Offset(0, 3).Value
            strSUNRoomType(intArrayCounter, 2) = CELL.Offset(0, 4).Value
            
            intArrayCounter = intArrayCounter + 1
            
        End If
        
    Next
        
    For intOverallCounter = 1 To 9
        
        'Replace the SUN Room Type with the overall room type in the SUN Data
        For Each CELL In Worksheets("SUN DATA").Range("B1:B2000").Columns(1).Cells
            
            If CELL.Value = strSUNRoomType(intOverallCounter, 2) Then
                            
                'Replace the SUN Room Type
                CELL.Value = strSUNRoomType(intOverallCounter, 1)
                
            End If
            
        Next
    
    Next intOverallCounter
    
End Sub
