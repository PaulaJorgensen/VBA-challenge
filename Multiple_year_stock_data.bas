Attribute VB_Name = "Module1"
Sub TickerCount()

' Set CurrentWs as a worksheet object variable.
Dim CurrentWs As Worksheet
Dim Need_Summary_Table_Header As Boolean
    
' Loop through all of the worksheets in the active workbook.
For Each CurrentWs In Worksheets
    
' Set initial variables
Dim Ticker_Name As String
    Ticker_Name = " "
        
Dim Total_Ticker_Volume As Double
    Total_Ticker_Volume = 0
        
Dim Open_Price As Double
    Open_Price = 0
    
Dim Close_Price As Double
    Close_Price = 0
    
Dim Delta_Price As Double
    Delta_Price = 0
    
Dim Delta_Percent As Double
    Delta_Percent = 0
    
Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
        
Dim Lastrow As Long
    Dim i As Long
        
Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

' Set Titles for the Summary Table
CurrentWs.Range("I1").Value = "Ticker"
CurrentWs.Range("J1").Value = "Yearly Change"
CurrentWs.Range("K1").Value = "Percent Change"
CurrentWs.Range("L1").Value = "Total Stock Volume"
        
'Set opening price
Open_Price = CurrentWs.Cells(2, 3).Value
        
For i = 2 To Lastrow
        
If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
    Ticker_Name = CurrentWs.Cells(i, 1).Value
    Close_Price = CurrentWs.Cells(i, 6).Value
    Delta_Price = Close_Price - Open_Price
        If Open_Price <> 0 Then
         Delta_Percent = (Delta_Price / Open_Price) * 100
        End If
                
    Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
    
    CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
    CurrentWs.Range("J" & Summary_Table_Row).Value = Delta_Price
    
        If (Delta_Price > 0) Then
            CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf (Delta_Price <= 0) Then
            CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
                
    CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")
    CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
    
    'Reset variables
    Summary_Table_Row = Summary_Table_Row + 1
    Delta_Price = 0
    Close_Price = 0
    Open_Price = CurrentWs.Cells(i + 1, 3).Value
    Delta_Percent = 0
    Total_Ticker_Volume = 0
                
Else
    Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
End If
      
Next i
        
Next CurrentWs

End Sub

