
Sub Module2():

     
      For Each ws In Worksheets
 
    
      Dim ticker As String
      Dim Open_Price As Double
      
      Dim Close_Price As Double
      
      Dim Yearly_Change As Double
      Dim Percent_Change As Double
      
      Dim Greatest_Increase As Double
      Dim Greatest_Decrease As Double
      Dim Greatest_Total As Double
      Dim Greatest_Increase_Ticker As String
      Dim Greatest_Decrease_Ticker As String
      Dim Greatest_Total_Ticker As String
      
     
      Dim Price_Row As Long
      Price_Row = 2
      
    
      Total = 0
      
     
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
      
   
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
      
   
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
     
      For i = 2 To LastRow:
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
               
               ticker = ws.Cells(i, 1).Value
               
              
               Total = Total + ws.Range("G" & i).Value
               
              
               ws.Range("I" & Summary_Table_Row).Value = ticker
               
               
               ws.Range("L" & Summary_Table_Row).Value = Total
               
               
               Open_Price = ws.Range("C" & Price_Row).Value
               Close_Price = ws.Range("F" & i).Value
               Yearly_Change = Close_Price - Open_Price
               
                  If Open_Price = 0 Then
                      Percent_Change = 0
                     Else
                         Percent_Change = Yearly_Change / Open_Price
                  End If
                 
                
                  ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                  ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                  ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                  
                        
                        If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        Else
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
                  
                  
                  Summary_Table_Row = Summary_Table_Row + 1
                  Price_Row = i + 1
               
                  
                  Total = 0
            Else
              Total = Total + ws.Range("G" & i).Value
                 
            End If
                      
        Next i
        
        Greatest_Increase = ws.Range("K2").Value
        Greatest_Decrease = ws.Range("K2").Value
        Greatest_Total = ws.Range("L2").Value
        
        
        Lastrow_Ticker = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
       
         For r = 2 To Lastrow_Ticker:
               If ws.Range("K" & r + 1).Value > Greatest_Increase Then
                  Greatest_Increase = ws.Range("K" & r + 1).Value
                  Greatest_Increase_Ticker = ws.Range("I" & r + 1).Value
               ElseIf ws.Range("K" & r + 1).Value < Greatest_Decrease Then
                  Greatest_Decrease = ws.Range("K" & r + 1).Value
                  Greatest_Decrease_Ticker = ws.Range("I" & r + 1).Value
                ElseIf ws.Range("L" & r + 1).Value > Greatest_Total Then
                  Greatest_Total = ws.Range("L" & r + 1).Value
                  Greatest_Total_Ticker = ws.Range("I" & r + 1).Value
                End If
            Next r
            
            
            ws.Range("P2").Value = Greatest_Increase_Ticker
            ws.Range("P3").Value = Greatest_Decrease_Ticker
            ws.Range("P4").Value = Greatest_Total_Ticker
            ws.Range("Q2").Value = Greatest_Increase
            ws.Range("Q3").Value = Greatest_Decrease
            ws.Range("Q4").Value = Greatest_Total
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            
    Next ws
End Sub


