# VBA_Challenge
This challenge utilizes VBA skills learned during the lesson. 


Sub stock()
'Define Variables

Dim total_volume As Double
Dim counter As Integer
Dim yearly_change As Double
Dim start_value As Long
Dim percent_change As Double
Dim temp As Long

Dim wb As Workbook
Set wb = ActiveWorkbook

For Each ws In wb.Sheets



    counter = 2
    start_value = 2
    temp = 2
    
   ws.Range("I1").Value = "Ticker"
   ws.Range("L1").Value = "Total Stock Volume"
   ws.Range("J1").Value = "yearly_change"
   ws.Range("K1").Value = "percent_change"
    
    Row_Count = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For i = 2 To Row_Count
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            If ws.Cells(start_value, 3) = 0 Then
                For firstnonzero = start_value To i
                    If ws.Cells(firstnonzero, 3).Value <> 0 Then
                        start_value = firstnonzero
                        Exit For
                   End If
               Next firstnonzero
           End If
           
            yearly_change = ws.Cells(i, 6) - ws.Cells(temp, 3)
            
            percent_change = (yearly_change / ws.Cells(start_value, 3)) * 100
            percent_change = Round(percent_change, 2)
            
            If yearly_change > 0 Then
            ws.Range("J" & counter).Interior.ColorIndex = 4
            Else
            ws.Range("J" & counter).Interior.ColorIndex = 3
            
            End If
            
            ws.Range("I" & counter).Value = Cells(i, 1).Value
            ws.Range("J" & counter).Value = yearly_change
            ws.Range("K" & counter).Value = percent_change
            ws.Range("L" & counter).Value = total_volume
            
            total_volume = 0
            yearly_change = 0
            counter = counter + 1
            
            temp = i + 1
            
       Else:
       
            total_volume = total_volume + ws.Cells(i, 7).Value
            
        End If
        
      

        Next i
        
    Next ws
    
    
    End Sub

