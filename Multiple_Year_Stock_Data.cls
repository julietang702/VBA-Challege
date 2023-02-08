
Sub MYSData():
  Dim Source As Worksheet
    For Each WS In Worksheets
    WorksheetName = WS.Name


        
        'Create column headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"


        'Functionality/Summary headers
        Range(" O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        
        

                    
End Sub

