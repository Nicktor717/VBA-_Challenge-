Attribute VB_Name = "Module2"
Sub RunCodeOnAllWorksheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ' Your code here
        
             Dim Ticker As String
    Dim Change As Double
    Dim Percent As Double
    Dim Tolvolume As Double
    Dim Row As Integer
    Dim OpenP As Double
    Dim CloseP As Double
    Row = 2
    Tolvolume = 0
    
    
    
    
    

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = " Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total volume stouck"
        
        ',,,,,,,,,,,,,,,
        
        
        ws.Cells(1, 16).Value = " Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decreas"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        
        
        OpenP = Range("C3").Value
        
        
     
     For i = 2 To LastRow
         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         ws.Range("I" & Row).Value = ws.Cells(i, 1).Value
         Tolvolume = Tolvolume + ws.Cells(i, 7).Value
         ws.Range("L" & Row).Value = Tolvolume
         Tolvolume = 0
          CloseP = ws.Cells(i, 6).Value
          Change = CloseP - OpenP
          ws.Range("J" & Row).Value = Change
            
          ws.Range("K" & Row).Value = (CloseP - OpenP) / OpenP
          Row = Row + 1
          OpenP = ws.Cells(i + 1, 3).Value
            
            Else
            Tolvolume = Tolvolume + ws.Cells(i, 7).Value
            CloseP = ws.Cells(i + 1, 6).Value
            
           If ws.Range("J" & Row) >= 0 Then
            ws.Range("J" & Row).Interior.ColorIndex = 4
            
           Else
           ws.Range("J" & Row).Interior.ColorIndex = 3
            
           End If
            
            
        End If
     Next i
     Next ws
        

End Sub
