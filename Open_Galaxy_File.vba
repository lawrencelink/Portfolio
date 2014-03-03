Option Explicit

Public Sub Open_Galaxy_File(strMatcherName As String, intNumberofRooms, ByRef strGXRoomType As Variant, strProperty)

Application.Calculation = xlCalculationManual

Dim strFileName As String
Dim strWorkBookName As String
Dim strFileClose

strWorkBookName = ActiveWorkbook.Name

'Show the open dialog and pass the selected _
'file name to the String variable "strFileName
strFileName = Application.GetOpenFilename

'If they have cancelled
If strFileName = "False" Then

MsgBox "Stopping because you did not select a file"

Exit Sub

Else
    If strProperty = "WSJ" Or strProperty = "SDO" Then
        Workbooks.OpenText Filename:=strFileName, Origin:=437, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, Comma:=False, _
        Space:=True, Other:=True, OtherChar:="-", FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
        , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1)), _
        TrailingMinusNumbers:=True
    Else
        Workbooks.OpenText Filename:=strFileName, DataType:=xlDelimited, Space:=True, ConsecutiveDelimiter:=True
    End If
End If

Call Galaxy_Extractor_Mini(intNumberofRooms, strGXRoomType)

Call Find_Date

Call P2FILE_MOVER(strWorkBookName)

'MsgBox Len(strFileName)
If Len(strFileName) = 27 Then
    
    strFileClose = Mid(strFileName, 13, 8)
    Application.DisplayAlerts = False
    Windows(strFileClose).Close
    Application.DisplayAlerts = True
    'MsgBox strFileClose & "27"
End If

    If Len(strFileName) = 28 Then
        strFileClose = Mid(strFileName, 14, 8)
        Application.DisplayAlerts = False
        Windows(strFileClose).Close
        Application.DisplayAlerts = True
    End If
    
    If Len(strFileName) = 26 Then
        strFileClose = Mid(strFileName, 12, 8)
        Application.DisplayAlerts = False
        Windows(strFileClose).Close
        Application.DisplayAlerts = True
    End If
If Len(strFileName) = 40 Then
        strFileClose = Mid(strFileName, 12, 8)
        Application.DisplayAlerts = False
        Windows(strFileClose).Close
        Application.DisplayAlerts = True
    End If

Windows(strWorkBookName).Activate

End Sub

Private Sub Galaxy_Extractor_Mini(intNumberofRooms, ByRef strGXRoomType As Variant)

Dim sh As Worksheet
Dim rng As Range
Dim strcurrentcell As String

Dim intbscounter As Integer
Dim intBLCounter As Integer
           
Dim strtargetNET As String
Dim strtargetAVL As String

Dim strcolumnNET As String
Dim strcolumnAVL As String
Dim introwNET As Integer
Dim introwAVL As Integer
Dim blnBoolean As Boolean

introwNET = 1
introwAVL = 1
intbscounter = 1
intBLCounter = 1

If intNumberofRooms > 1 Then

    For Each rng In Range("C1:C5000")
    
        Select Case rng
            
            Case strGXRoomType(intbscounter, 1)
                
                rng.Activate
                            
                strcurrentcell = rng.Address
                      
                Call column(intbscounter, strcolumnNET, strcolumnAVL)
                            
                strtargetNET = strcolumnNET & introwNET
                strtargetAVL = strcolumnAVL & introwAVL
                
                Call Move_It(strcurrentcell, strtargetNET, strtargetAVL)
                
                If intbscounter = intNumberofRooms Then
                    blnBoolean = True
                    
                    '************************
                    introwNET = introwNET + 20
                    introwAVL = introwAVL + 20
                    '************************

                End If
                
                If blnBoolean = True Then
                    intbscounter = 1
                Else
                    intbscounter = intbscounter + 1
                End If
                    blnBoolean = False
        End Select
    Next
 
Else


    For Each rng In Range("C1:C5000")
    
        Select Case rng
            
            Case strGXRoomType(1, 1)
                
                rng.Activate
                            
                strcurrentcell = rng.Address
                      
                Call column(intbscounter, strcolumnNET, strcolumnAVL)
                            
                strtargetNET = strcolumnNET & introwNET
                strtargetAVL = strcolumnAVL & introwAVL
                
                Call Move_It(strcurrentcell, strtargetNET, strtargetAVL)
                
                If intbscounter = intNumberofRooms Then
                    blnBoolean = True
                    
                    '************************
                    introwNET = introwNET + 20
                    introwAVL = introwAVL + 20
                    '************************
                    
                End If
                
                If blnBoolean = True Then
                    intbscounter = 1
                Else
                    intbscounter = intbscounter + 1
                End If
                    blnBoolean = False
        End Select
    Next

End If
 
End Sub

Function Move_It(strcurrentcell As String, strtargetNET, strtargetAVL)
                    
                        Range(strcurrentcell).Select
                        
                        ActiveCell.Offset(2, 0).Select
                        
                        Range(ActiveCell, ActiveCell.Offset(0, 19)).Copy
                        
                        Range(strtargetNET).Select
                                                             
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=True
                        
                        
                        'RETURN FOR AVAILABILITY
                        Range(strcurrentcell).Select
                        
                        ActiveCell.Offset(6, 0).Select
                        
                        Range(ActiveCell, ActiveCell.Offset(0, 19)).Copy
                        
                        Range(strtargetAVL).Select
                                      
                        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                        False, Transpose:=True
                        

End Function

Private Sub Find_Date()
    Dim strMonth As String
    Dim intDate As Integer
    Dim strYear As String
    Dim myMonth As Integer
    
    Range("B5").Select
    strMonth = ActiveCell.Value
    
    ActiveCell.Offset(2, 0).Select
    intDate = ActiveCell.Value
    
    Call Galaxy_Get_the_Date(strMonth, myMonth)
    
    Range("E2").Select
    
    If Right(ActiveCell.Value, 4) = "2014" Or Right(ActiveCell.Value, 4) = "2013" Then
        'End
    Else
        Range("F2").Select
        
        If Right(ActiveCell.Value, 4) = "2014" Or Right(ActiveCell.Value, 4) = "2013" Then
            'End
            
        Else
            Range("G2").Select
            
            If Right(ActiveCell.Value, 4) = "2014" Or Right(ActiveCell.Value, 4) = "2013" Then
                'MsgBox Right(ActiveCell.Value, 4)
                'End
            Else
                Range("H2").Select
            End If
        End If
    End If
    
    strYear = Right(ActiveCell.Value, 4)
    
    Range("AA1").Select
    ActiveCell.Value = myMonth & "/" & intDate & "/" & strYear
    
    Selection.DataSeries Rowcol:=xlRows, Type:=xlChronological, Date:=xlDay, _
        Step:=1, Trend:=False
    Selection.AutoFill Destination:=Range("AA1:AA400"), Type:=xlFillDefault
    
End Sub

Function Galaxy_Get_the_Date(strMonth, myMonth)

Select Case strMonth
    
    Case "JAN"
        myMonth = 1
    
    Case "FEB"
        myMonth = 2
    
    Case "MAR"
        myMonth = 3
    
    Case "APR"
        myMonth = 4

     
    Case "MAY"
        myMonth = 5


    Case "JUN"
        myMonth = 6

        
    Case "JUL"
        myMonth = 7
 
    Case "AUG"
        myMonth = 8
 
    Case "SEP"
        myMonth = 9
        
    Case "OCT"
        myMonth = 10
  
    Case "NOV"
        myMonth = 11

    Case "DEC"
        myMonth = 12

End Select

End Function

Private Sub column(ByVal intbscounter, ByRef strcolumnNET, ByRef strcolumnAVL)

            If intbscounter = 1 Then
                    strcolumnNET = "AB"
                    strcolumnAVL = "AC"
            End If
            If intbscounter = 2 Then
                    strcolumnNET = "AD"
                    strcolumnAVL = "AE"
            End If
            If intbscounter = 3 Then
                    strcolumnNET = "AF"
                    strcolumnAVL = "AG"
            End If
            If intbscounter = 4 Then
                    strcolumnNET = "AH"
                    strcolumnAVL = "AI"
            End If
            If intbscounter = 5 Then
                    strcolumnNET = "AJ"
                    strcolumnAVL = "AK"
            End If
            If intbscounter = 6 Then
                    strcolumnNET = "AL"
                    strcolumnAVL = "AM"
            End If
            If intbscounter = 7 Then
                    strcolumnNET = "AN"
                    strcolumnAVL = "AO"
            End If
            If intbscounter = 8 Then
                    strcolumnNET = "AP"
                    strcolumnAVL = "AQ"
            End If
            If intbscounter = 9 Then
                    strcolumnNET = "AR"
                    strcolumnAVL = "AS"
            End If
            If intbscounter = 10 Then
                    strcolumnNET = "AT"
                    strcolumnAVL = "AU"
            End If
            If intbscounter = 11 Then
                    strcolumnNET = "AV"
                    strcolumnAVL = "AW"
            End If
            If intbscounter = 12 Then
                    strcolumnNET = "AX"
                    strcolumnAVL = "AY"
            End If
            If intbscounter = 13 Then
                    strcolumnNET = "AZ"
                    strcolumnAVL = "BA"
            End If
            If intbscounter = 14 Then
                    strcolumnNET = "BB"
                    strcolumnAVL = "BC"
            End If
            If intbscounter = 15 Then
                    strcolumnNET = "BD"
                    strcolumnAVL = "BE"
            End If
            If intbscounter = 16 Then
                    strcolumnNET = "BF"
                    strcolumnAVL = "BG"
            End If
            If intbscounter = 17 Then
                    strcolumnNET = "BH"
                    strcolumnAVL = "BI"
            End If
            If intbscounter = 18 Then
                    strcolumnNET = "BJ"
                    strcolumnAVL = "BK"
            End If
            If intbscounter = 19 Then
                    strcolumnNET = "BL"
                    strcolumnAVL = "BM"
            End If
            If intbscounter = 20 Then
                    strcolumnNET = "BN"
                    strcolumnAVL = "BO"
            End If
            If intbscounter = 21 Then
                    strcolumnNET = "BP"
                    strcolumnAVL = "BQ"
            End If
            
End Sub

Public Sub P2FILE_MOVER(strWorkBookName)
        
    'PASTE THE DATA INTO THE "SUN DATA SHEET" TAB
    Range("AA1").Select
    Range(Selection, ActiveCell.Offset(0, 30)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveCell.Offset(0, 1).Select
    'Move the data
    Windows(strWorkBookName).Activate
    Sheets("GALAXY DATA").Visible = True
    Sheets("GALAXY DATA").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("GALAXY DATA").Visible = False
    Sheets("GALAXY").Select

End Sub


