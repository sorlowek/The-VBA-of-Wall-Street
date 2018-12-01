Private Sub easy()

    Dim ticker As String
    Dim vol As Double

    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double

    Cells(1, 9).Value = "ticker"
    Cells(1, 10).Value = "Yearly_change"
    Cells(1, 12).Value = "Total Stock Vol"
    Cells(1, 11).Value = "Percent_change"

    Summary_Table_Row = 2

Dim lastrow As Long
    With ActiveSheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    End With
    
    For i = 2 To lastrow

        ticker = Cells(i, 1).Value

      If Cells(i - 1, 1).Value = Cells(i, 1).Value And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          year_close = Cells(i, 6).Value
          
          year_open = Cells(i, 3).Value
          
          yearly_change = year_close - year_open
          
          percent_change = yearly_change / year_open
          
          Range("j" & Summary_Table_Row).Value = yearly_change

          Range("I" & Summary_Table_Row).Value = ticker

          Range("K" & Summary_Table_Row).Value = percent_change

          Range("L" & Summary_Table_Row).Value = vol

          Summary_Table_Row = Summary_Table_Row + 1

          vol = 0


      Else

          vol = vol + Cells(i, 7).Value


      End If
      
      Cells(i, 10).Value = yearly_change
      Cells(i, 11).Value = percent_change
      
For x = 2 To lastrow

    If Cells(x, 11) > 0 Then
    
    Cells(x, 11).Interior.ColorIndex = 4
    
    Else
    
    Cells(x, 11).Interior.ColorIndex = 3
    
    End If
    
    Next x
    Next i
    
End Sub

