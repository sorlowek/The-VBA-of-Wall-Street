Private Sub easy()

    Dim ticker As String
    Dim vol As Double
    Dim Summary_Table_Row As Integer

    Cells(1, 9).Value = "ticker"
    Cells(1, 10).Value = "Total Stock Vol"

    Summary_Table_Row = 2

Dim lastrow As Long
    With ActiveSheet
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    End With
    
    For i = 2 To lastrow
        ticker = Cells(i, 1).Value

      If Cells(i - 1, 1).Value = Cells(i, 1).Value And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          
          Range("I" & Summary_Table_Row).Value = ticker

          Range("J" & Summary_Table_Row).Value = vol

          Summary_Table_Row = Summary_Table_Row + 1

          vol = 0


      Else

          vol = vol + Cells(i, 7).Value


      End If
      
      
       Next i
       
    
End Sub

