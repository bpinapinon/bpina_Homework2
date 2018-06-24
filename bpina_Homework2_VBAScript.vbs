Attribute VB_Name = "Module1"
Sub BP_Macro()
    For Each W In Worksheets
    
        'Declare variable to hold Ticket and Total Volumne for each ticker
        Dim vchTicker As String
        Dim dVolumneTotal As Double
        dVolumneTotal = 0
        
        'Set location for each Ticker in the summary table
        Range("J1").Value = "Ticker"
        Range("K1").Value = "TotalVolume"
        Dim iSummaryRow As Integer
        iSummaryRow = 2 'account for header row
        
        'Count the number of rows in data set
        iLastRowCount = Cells(Rows.Count, 1).End(xlUp).Row 'Count the number of records
    
        'Loop through all Ticker rows
        For r = 2 To iLastRowCount
        
            'Check if still within the same Ticker
            If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then 'Next cell <> current cell
               
               'Set Ticker Name
               vchTickerName = Cells(r, 1).Value
               
               'Add to the Volume total
               dVolumeTotal = dVolumeTotal + Cells(r, 7).Value
               
               'Add Ticker Name and Total Volume to Summary Table
               Range("J" & iSummaryRow).Value = vchTickerName
               Range("K" & iSummaryRow).Value = Format(dVolumeTotal, "Currency") 'Format Currency
               
               'Add one to Summary Row counter
               iSummaryRow = iSummaryRow + 1
               
               'Reset Card Total
               dVolumeTotal = 0
               
            Else 'If following cell is the same card
               dVolumeTotal = dVolumeTotal + Cells(r, 7).Value
               
            End If
        Next r
    
    Next W
    
End Sub


