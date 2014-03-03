Sub Supply_Suggestions()
    
    Call Clear_Suggestion
    
    Dim intSupply() As Variant
    'varHierarchy is an array that holds the supply and availability for each room type
    Dim varHierarchy() As Variant
    Dim varAvailability(1 To 20) As Variant
    Dim varRoomType As Variant
    Dim intNONSupply As Integer
    Dim intSVNSupply As Integer
    Dim intSupplyVariance As Integer
    'intDistribution is the amount that will be added or subtracted from each room type as it loops
    Dim intDistribution As Integer
    Dim intSupplyRemainder As Integer
    Dim intAllocation As Integer
    Dim intSupplyChange As Integer
    Dim intNumberSUNRoomTypes As Integer
    Dim intRow As Integer
    'intAVL is the number of columns to the right of H that the procedure must move to gather the availability
    Dim intAVL As Integer
    Dim intMaxKeyCount As Integer
    Dim intSuggestedSupply As Integer
    Dim i As Integer
    Dim i2 As Integer
    Dim i3 As Integer
    'sometimes the varHierarchy goes past the last room type
    Dim i4 As Integer
    Dim intSUNColumn As Integer
    Dim intSuggestedColumn As Integer
    'intIndex is the number of rows in the varHierarchy array that will be affected next in the loop
    Dim intIndex As Integer
    Dim intIndex2 As Integer
    'intCurrentTally is the supply for a certain room type plus the additional space that was allocated to it
    Dim intCurrentTally As Integer
    Dim intCurrentTallyTemp As Integer
    Dim intCurrentTallyNON As Integer
    Dim intCurrentTallyAVL As Integer
    Dim Nrows As Long, Ncols As Integer
    Dim intNumberofSUNRoomTypes As Integer
    Dim intLastSuggestionRow As Integer
    Dim intFITSUPPLY As Integer
    Dim intTotalSupply As Integer
    Dim strSUNColumn As String
    Dim strSupply As String
    Dim strToday As Date
    Dim datLastDayofYear As Date
    
    intNumberSUNRoomTypes = Sheets("INFO").Range("A9").Value
    varHierarchy = Sheets("INFO").Range("C2:H21").Value
    intNumberofSUNRoomTypes = Sheets("INFO").Range("A9").Value
    
    blnBoolean = False
    strToday = Date
    datLastDayofYear = "12/31/14"
    intLastSuggestionRow = datLastDayofYear - strToday
        
    Sheets("INFO").Visible = True
    Sheets("SUN").Select
    
    'For each variance of FIT to SUN for a full year
    For intRow = 2 To intLastSuggestionRow
        
        intSupplyVariance = Range("H" & intRow).Value
        intFITSUPPLY = Sheets("SUN").Range("E" & intRow).Value
        intTotalSupply = intFITSUPPLY - intSupplyVariance
        'Reset the variables
        intAVL = 11
        intIndex = 1
        i2 = 1
        i3 = 0
        
        If intSupplyVariance <> 0 Then
            
            'Populate the Availability for each room type on the INFO tab column E
            For i = 1 To intNumberSUNRoomTypes
                
                i2 = i2 + 1
                varHierarchy(i, 3) = Range(Cells(intRow, intAVL - 1), Cells(intRow, intAVL - 1)).Value
                varHierarchy(i, 5) = Range(Cells(intRow, intAVL + 1), Cells(intRow, intAVL + 1)).Value
                varHierarchy(i, 6) = Range(Cells(intRow, intAVL + 2), Cells(intRow, intAVL + 2)).Value
                varAvailability(i) = Range(Cells(intRow, intAVL), Cells(intRow, intAVL)).Value
                intAVL = intAVL + 5
                
            Next i
                        
            'Increase or decrease supply in each room type
            If intSupplyVariance > 0 Then
                
                'This block of code is executed if the inventory in SUN needs to be increased
                Sheets("SUN").Select
                
                Do Until intTotalSupply >= intFITSUPPLY
                    
                    'intDistribution is set up here
                    
                    intDistribution = WorksheetFunction.Min(varAvailability) + Round(WorksheetFunction.Average(varAvailability))
                    intDistribution = Abs(intDistribution)
                    If i3 = 0 Then
                        intIndex = Application.Match(WorksheetFunction.Min(varAvailability), varAvailability, 0)
                    Else
                        If intIndex = 16 Then
                            intIndex = 1
                        End If
                        intIndex = intIndex + 1
                        If i3 > 16 Then
                            intDistribution = 1
                        End If
                    End If
                    'Make sure that the intTotalSupply is not increased above the intFITSupply
                    
                    If (intFITSUPPLY < intTotalSupply + intDistribution) Then
                        intDistribution = intFITSUPPLY - intTotalSupply
                    
                    ElseIf intDistribution = 0 Then
                        intDistribution = 1
                    
                    End If
                                        
                    intSuggestedColumn = varHierarchy(intIndex, 4)
                    intSUNColumn = varHierarchy(intIndex, 4) + 1
                        If intIndex > intNumberofSUNRoomTypes Then
                            intSUNColumn = varHierarchy(1, 4) + 1
                            
                            If i4 > 2 Then
                                intSUNColumn = varHierarchy(2, 4) + 1
                            End If
                            i4 = i4 + 1
                            
                        End If
                    intSuggestedSupply = Sheets("SUN").Range(Cells(intRow, intSUNColumn), Cells(intRow, intSUNColumn)).Value + intDistribution
                    intNONSupply = Sheets("SUN").Range(Cells(intRow, intSUNColumn + 2), Cells(intRow, intSUNColumn + 2)).Value
                    intSVNSupply = Sheets("SUN").Range(Cells(intRow, intSUNColumn + 3), Cells(intRow, intSUNColumn + 3)).Value + intDistribution
                    intCurrentTallyTemp = varHierarchy(intIndex, 3) + intDistribution
                    intMaxKeyCount = varHierarchy(intIndex, 2)
                    
                    If intSuggestedSupply <= intMaxKeyCount And intCurrentTallyTemp <= intMaxKeyCount Then
                    
                        intTotalSupply = intTotalSupply + intDistribution
                        intCurrentTally = varHierarchy(intIndex, 3) + intDistribution
                        intCurrentTallyNON = varHierarchy(intIndex, 5)
                        intCurrentTallyAVL = varHierarchy(intIndex, 6) + intDistribution
                        'Ensure that room type has ability to add supply
                        
                            If Range(Cells(intRow, intSuggestedColumn), Cells(intRow, intSuggestedColumn)).Value = 0 Then
                                'Enter the suggested supply in the cell on the spreadsheet
                                strSupply = "(" & intSuggestedSupply & "," & intNONSupply & "," & intSVNSupply & ")"
                                
                                Range(Cells(intRow, intSuggestedColumn), Cells(intRow, intSuggestedColumn)).Value = strSupply
                                
                            Else
                                'Enter the suggested supply in the cell on the spreadsheet
                                Range(Cells(intRow, intSuggestedColumn), Cells(intRow, intSuggestedColumn)).Value = "(" & intCurrentTally & "," & intCurrentTallyNON & "," & intCurrentTallyAVL & ")"
                            End If
                        
                        'increase the supply for that room type
                        varHierarchy(intIndex, 3) = varHierarchy(intIndex, 3) + intDistribution
                        varHierarchy(intIndex, 5) = varHierarchy(intIndex, 5)
                        varHierarchy(intIndex, 6) = varHierarchy(intIndex, 6) + intDistribution
                        'increase the availability for that room type
                        varAvailability(intIndex) = varAvailability(intIndex) + intDistribution
                        i3 = 0
                    Else
                        i3 = i3 + 1
                    End If
                Loop
                
            Else
                'This block of code is executed if the inventory in SUN needs to be reduced
                Sheets("SUN").Select
                
                Do Until intTotalSupply <= intFITSUPPLY
                    
                    'intDistribution is set up here
                    intDistribution = WorksheetFunction.Max(varAvailability) - Round(WorksheetFunction.Average(varAvailability))
                    intIndex = Application.Match(WorksheetFunction.Max(varAvailability), varAvailability, 0)
                    '''''''MsgBox WorksheetFunction.Max(varAvailability)
                    'Make sure that the intTotalSupply is not reduced below the intFITSupply
                    If Abs(intSupplyVariance) <= intDistribution Then
                        'intDistribution = Abs(intSupplyVariance)
                        intDistribution = intTotalSupply - intFITSUPPLY
             
                    'make sure that the FITSUPPLY matches the suggested column
                    ElseIf (intFITSUPPLY > intTotalSupply - intDistribution) Then
                        intDistribution = intTotalSupply - intFITSUPPLY
                    
                    ElseIf intDistribution = 0 Then
                        intDistribution = 1
                    
                    End If
                    
                    intTotalSupply = intTotalSupply - intDistribution
                    
                    intAVL = varAvailability(intIndex)
                    intSuggestedColumn = varHierarchy(intIndex, 4)
                    intSUNColumn = varHierarchy(intIndex, 4) + 1
                    
                    intSuggestedSupply = Sheets("SUN").Range(Cells(intRow, intSUNColumn), Cells(intRow, intSUNColumn)).Value - intDistribution
                    intNONSupply = Sheets("SUN").Range(Cells(intRow, intSUNColumn + 2), Cells(intRow, intSUNColumn + 2)).Value
                    intSVNSupply = Sheets("SUN").Range(Cells(intRow, intSUNColumn + 3), Cells(intRow, intSUNColumn + 3)).Value - intDistribution
                    
                    If intAVL > 0 Then
                        
                        intCurrentTally = varHierarchy(intIndex, 3) - intDistribution
                        intCurrentTallyNON = varHierarchy(intIndex, 5)
                        intCurrentTallyAVL = varHierarchy(intIndex, 6) - intDistribution
                        
                        If Range(Cells(intRow, intSuggestedColumn), Cells(intRow, intSuggestedColumn)).Value = 0 Then
                            'Enter the suggested supply in the cell on the spreadsheet
                            Range(Cells(intRow, intSuggestedColumn), Cells(intRow, intSuggestedColumn)).Value = "(" & intSuggestedSupply & "," & intNONSupply & "," & intSVNSupply & ")"
                            
                        Else
                            'Enter the suggested supply in the cell on the spreadsheet
                            Range(Cells(intRow, intSuggestedColumn), Cells(intRow, intSuggestedColumn)).Value = "(" & intCurrentTally & "," & intCurrentTallyNON & "," & intCurrentTallyAVL & ")"
                        End If
                        
                        'Decrease the supply for that room type
                        varHierarchy(intIndex, 3) = varHierarchy(intIndex, 3) - intDistribution
                        varHierarchy(intIndex, 5) = varHierarchy(intIndex, 5)
                        varHierarchy(intIndex, 6) = varHierarchy(intIndex, 6) - intDistribution
                        'Decrease the availability for that room type
                        varAvailability(intIndex) = varAvailability(intIndex) - intDistribution
                           
                    End If
                    
                Loop
            End If
            
        End If
        
    Next intRow
    
    ''''Sheets("TEMP").Visible = False
End Sub