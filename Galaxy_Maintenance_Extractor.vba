Public Sub Galaxy_Maintenance_Extractor()
'Extract the data from the Galaxy report
Application.ScreenUpdating = False

'Declare Exctractor variables
Dim sh As Worksheet
Dim rng As Range

'Declare string variables
Dim strRoomNumber As String
Dim strUnitType As String
Dim strStartDate As String
Dim strEndDate As String
Dim strLineAddress
Dim strCurrentCell As String
'Declare counter variables
Dim intCounter As Integer
Dim intcounterREN As Integer
Dim intCounterCON As Integer
'set the counter as positive so by the time it gets to A8 it will be zero
intCounter = 6
intcounterREN = 6
intCounterCON = 6
'Capture all the data
For Each rng In Range("A1:A4348")
    
    If IsNumeric(rng.Value) And Not rng.Value = "" Then
        'Find the row with needed data and activate it
        rng.Activate
        'Store the address of the first cell of that row
        strLineAddress = rng.Address
        'DILENEATE REN AND MW
        ActiveCell.Offset(0, 4).Select
        strCurrentCell = ActiveCell.Value
        Range(strLineAddress).Select
        
        Select Case strCurrentCell
            
            Case "MW"
        
                'copy the room number of that cell
                strRoomNumber = ActiveCell.Value
                'move over to the right
                ActiveCell.Offset(0, 2).Select
                'copy the room type
                strUnitType = ActiveCell.Value
                'move over to the right
                ActiveCell.Offset(0, 5).Select
                
                Do Until blnMoveOver = True
        
                    'strCurrentCellContents = ActiveCell.Value
                    'ActiveCell.NumberFormat
                    If InStr(ActiveCell.NumberFormat, "d") > 0 Then
                    'copy the start date
                    strStartDate = ActiveCell.Value
                    blnMoveOver = True
                    
                    Else
                    'Move over to the right
                    ActiveCell.Offset(0, 1).Select
                    intCounter2 = intCounter2 - 1
                    End If
                
                Loop
        
                blnMoveOver = False
                
                'Move over to the right
                ActiveCell.Offset(0, 1).Select
                'copy the enddate
                strEndDate = ActiveCell.Value
                'move over one to the right to start pasting
                ActiveCell.Offset(intCounter, 5).Select
                'paste the end date
                ActiveCell.Value = strRoomNumber
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strUnitType
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strStartDate
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strEndDate
                
                intcounterREN = intcounterREN - 1
                intCounterCON = intCounterCON - 1
        
        Case "MAIN"
        
                'copy the room number of that cell
                strRoomNumber = ActiveCell.Value
                'move over to the right
                ActiveCell.Offset(0, 2).Select
                'copy the room type
                strUnitType = ActiveCell.Value
                'move over to the right
                ActiveCell.Offset(0, 5).Select
                
                Do Until blnMoveOver = True
        
                    'strCurrentCellContents = ActiveCell.Value
                    'ActiveCell.NumberFormat
                    If InStr(ActiveCell.NumberFormat, "d") > 0 Then
                    'copy the start date
                    strStartDate = ActiveCell.Value
                    blnMoveOver = True
                    
                    Else
                    'Move over to the right
                    ActiveCell.Offset(0, 1).Select
                    intCounter2 = intCounter2 - 1
                    End If
                
                Loop
        
                blnMoveOver = False
                'Move over to the right
                ActiveCell.Offset(0, 1).Select
                'copy the enddate
                strEndDate = ActiveCell.Value
                'move over one to the right to start pasting
                ActiveCell.Offset(intCounter, 5).Select
                'paste the end date
                ActiveCell.Value = strRoomNumber
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strUnitType
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strStartDate
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strEndDate
                
                intcounterREN = intcounterREN - 1
                intCounterCON = intCounterCON - 1
        Case "RENO"
        
                'copy the room number of that cell
                strRoomNumber = ActiveCell.Value
                'move over to the right
                ActiveCell.Offset(0, 2).Select
                'copy the room type
                strUnitType = ActiveCell.Value
                'move over to the right
                ActiveCell.Offset(0, 5).Select
                'copy the room type
                
                Do Until blnMoveOver = True
        
                    'strCurrentCellContents = ActiveCell.Value
                    'ActiveCell.NumberFormat
                    If InStr(ActiveCell.NumberFormat, "d") > 0 Then
                    'copy the start date
                    strStartDate = ActiveCell.Value
                    blnMoveOver = True
                    
                    Else
                    'Move over to the right
                    ActiveCell.Offset(0, 1).Select
                    intCounter2 = intCounter2 - 1
                    End If
                
                Loop
        
                blnMoveOver = False
                
                ActiveCell.Offset(0, 1).Select
                'copy the enddate
                strEndDate = ActiveCell.Value
                'move over one to the right to start pasting
                ActiveCell.Offset(intcounterREN, 10).Select
                'paste the end date
                ActiveCell.Value = strRoomNumber
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strUnitType
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strStartDate
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strEndDate
                
                intCounter = intCounter - 1
        
        Case "REN"
        
                'copy the room number of that cell
                strRoomNumber = ActiveCell.Value
                'move over to the right
                ActiveCell.Offset(0, 2).Select
                'copy the room type
                strUnitType = ActiveCell.Value
                'move over to the right
                ActiveCell.Offset(0, 5).Select
                
                Do Until blnMoveOver = True
        
                    'strCurrentCellContents = ActiveCell.Value
                    'ActiveCell.NumberFormat
                    If InStr(ActiveCell.NumberFormat, "d") > 0 Then
                    'copy the start date
                    strStartDate = ActiveCell.Value
                    blnMoveOver = True
                    
                    Else
                    'Move over to the right
                    ActiveCell.Offset(0, 1).Select
                    intCounter2 = intCounter2 - 1
                    End If
                
                Loop
        
                blnMoveOver = False
                
                'Move over to the right
                ActiveCell.Offset(0, 1).Select
                'copy the enddate
                strEndDate = ActiveCell.Value
                'move over one to the right to start pasting
                ActiveCell.Offset(intCounter, 10).Select
                'paste the end date
                ActiveCell.Value = strRoomNumber
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strUnitType
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strStartDate
                'move over one to the right
                ActiveCell.Offset(0, 1).Select
                'paste the end date
                ActiveCell.Value = strEndDate
                
                intcounterREN = intcounterREN - 1
                intCounterCON = intCounterCON - 1
        
        Case "OTHR"
            intCounter = intCounter - 1
            intcounterREN = intcounterREN - 1
            intCounterCON = intCounterCON - 1
        End Select
    
    Else
        
        intCounter = intCounter - 1
        intcounterREN = intcounterREN - 1
        intCounterCON = intCounterCON - 1
    End If

Next

Application.ScreenUpdating = True

End Sub
