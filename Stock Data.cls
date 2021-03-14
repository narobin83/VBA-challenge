VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stockmarket():

      ' Define variables
      Dim i As Long
      Dim Ticker As String
      Dim Total_Stock_Volume As Double
      Dim Summary_Table_Row As Double
      Dim ws As Worksheet
      Dim Last_Row As Long
      Dim Opening_Price As Double
      Dim Closing_Price As Double
      Dim Yearly_Change As Double
      Dim Percent_Change As Double
       
    
      ' Loop through all worksheets
      For Each ws In ActiveWorkbook.Worksheets
      
      'Create Summary_Table
      ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
      
      ' Set starting points
      Total_Stock_Volume = 0
      Yearly_Change = 0
      Summary_Table_Row = 1
Next

      ' Detemine Last_Row
           Last_Row = ws.Cells(Rows.Count, 1).End(x1Up).Row
           
      ' Loop through Dataset
            For i = 2 To Last_Row
            
                ' Check for change in Ticker and set new starting values
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                      ' Add one to Summary_table_row for next entry
                      Summary_Table_Row = Summary_Table_Row + 1
                      Ticker = ws.Cells(i, 1).Value
                      Opening_Price = ws.Cells(i, 3).Value
                      ws.Range("I" & Summary_Table_Row).Value = Ticker
                      Total_Stock_Volume = 0
                      
                ' Set Closing_Price
                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                      Total_Stock_Volume = Total_Table_Row + ws.Cells(i, 7).Value
                      Closing_Price = ws.Cells(i, 6).Value
                      Yearly_Change = Closing_Price - Opening_Price
                      
                      If Yearly_Change = 0 Then
                          Percent_Change = 0
                      Else
                           Percent_Change = (Yearly_Change / Opening_Price)
                      End If
                      
                      ' Percent_Change casted to Double and multiply by 100
                      ws.Range("K" & Summary_Table_Row).Value = CDbl(Percent_Range) * 100
                      
                      ' Print Yearly_Change and Percent_change in summary table
                      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                      ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                      ' Format Percent_Change
                      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                Else
                      ' Add volumes if ticker is the same and print in summary table
                      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                      ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                End If
Next
   
                      If Range("J" & Summary_Table_Row).Value < 0 Then
                      ' Set the Cell color to red
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                      Else
                      ' Set the cell color to green
                      Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                      End If


                    
          

End Sub
