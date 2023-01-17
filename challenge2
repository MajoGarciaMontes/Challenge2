Attribute VB_Name = "Module1"
Sub StockChallenge2()

For Each ws In Worksheets
    
    Dim WorksheetName As String
        
    Dim Ticker As String
    Dim Volticker As Double
    Dim Row1 As Integer
    Dim Row2 As Integer
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim PercentageChange As Double
    Dim GreatestIncrease As Double
    Dim GreatesDecrease As Double
    Dim GreatesVolume As Double

  WorksheetName = ws.Name
  
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    
    Row1 = 2
    
    Volticker = 0
    
    openingPrice = ws.Cells(2, 3).Value

    For i = 2 To LastRow

       
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            Ticker = ws.Cells(i, 1).Value
            closingPrice = ws.Cells(i, 6).Value
            yearlyChange = (closingPrice - openingPrice)

        
            If openingPrice = 0 Then
                PercentageChange = 0
            Else
                PercentageChange = yearlyChange / openingPrice
            End If

    
            ws.Range("I" & Row1).Value = Ticker
            ws.Range("J" & Row1).Value = yearlyChange
            ws.Range("K" & Row1).Value = PercentageChange
            ws.Range("K" & Row1).NumberFormat = "0.00%"
            ws.Range("L" & Row1).Value = Volticker

           
            Row1 = Row1 + 1

           
            Volticker = 0
            openingPrice = ws.Cells(i + 1, 3)

        Else
            
            Volticker = Volticker + ws.Cells(i, 7).Value
        End If
    
    Next i

   
    LastRow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row

    
    For i = 2 To LastRow_summary_table
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    
    Next i

        
        LastRowSumtable = ws.Cells(Rows.Count, 9).End(xlUp).Row
             
            For i = 2 To LastRowSumtable
                If ws.Cells(i, 12).Value > GreatesVolume Then
                GreatesVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatesVolume = GreatesVolume
                
                End If
                
                If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestIncrease = GreatestIncrease
                
                End If
                
        
                If ws.Cells(i, 11).Value < GreatesDecrease Then
                GreatesDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatesDecrease = GreatesDecrease
                
                End If

            ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatesDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GreatesVolume, "Scientific")
            
            Next i
            
 Next ws

End Sub

