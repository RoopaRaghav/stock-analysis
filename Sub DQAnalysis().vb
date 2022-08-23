Sub MacroCheck()

    Dim testMessage As String

    testMessage = "Hello World!"

    MsgBox (testMessage)
    
    For i = 1 To 10

    Cells(1, i).Value = i * i

Next i

End Sub

Sub DQAnalysis()


    Worksheets("DQAnalysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row.
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    
    'start analysis
    
    Worksheets("2018").Activate
    
    
    Dim rowStart As Long
    
    Dim rowEnd As Long
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    
        
    
    rowStart = 3
    
    'rowEnd = MsgBox(vba_count_rows_with_data(rowEnd))
    
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    totalVolume = 0
    
    'iterate through thr 2018 worksheet to calculate the total volume for ticker DQ
    
    For i = rowStart To rowEnd
    
     If Cells(i, 1).Value = "DQ" Then
    
      totalVolume = totalVolume + Cells(i, 8).Value
      
     End If
      
      
      If (Cells(i, 1).Value = "DQ") Then
      If (Cells(i - 1, 1).Value <> "DQ") Then
      
       openPrice = Cells(i, 6).Value
    
      End If
      
      End If
      
       If (Cells(i, 1).Value = "DQ") Then
       If (Cells(i + 1, 1).Value <> "DQ") Then
      
       closePrice = Cells(i, 6).Value
    End If
      
      End If
      
      
    Next i
    
    
    
   
    
    
      
    'Update the total volume in the DQ Analysis worksheet
  
    Worksheets("DQAnalysis").Activate
    
    Cells(4, 1).Value = 2018
    
    Cells(4, 2).Value = totalVolume


    'Percentage increase or decrease in price from the beginning of the year to the end of the year

    Cells(4, 3).Value = (closePrice / openPrice) - 1
   
   


End Sub


Function vba_count_rows_with_data(counter As Long)


Dim iRange As Range

Worksheets("2018").Activate


With ActiveSheet.UsedRange

    'loop through each row from the used range
    For Each iRange In .Rows

        'check if the row contains a cell with a value
        If Application.CountA(iRange) > 0 Then
        
            'If (Cells(iRange, 1).vaue = "DQ") Then
            'counts the number of rows non-empty Cells
            counter = counter + 1
            'End If
        
        End If

    Next

End With

'MsgBox "Number of used rows is " & counter

End Function





