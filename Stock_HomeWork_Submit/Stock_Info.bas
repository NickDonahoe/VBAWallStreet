Attribute VB_Name = "Module1"
Sub Stock_Info()
  
  ' Initialize variables to loop through each sheet and the rows on each sheet
  Dim WS_Count As Integer
  Dim j As Integer
  Dim i As Long
  
  ' Initialize 2 variables to store the stock name and the next stock name
  Dim rCell As String
  Dim bCell As String

  'Set WS_Count equal to the number of worksheets in the active workbook
  WS_Count = ActiveWorkbook.Worksheets.Count
  
  ' Set an initial variable for holding the stock name
  Dim Stock_Name As String

  ' Set an initial variable for holding the total stock volume per stock
  Dim Vol_Total As Double
  Vol_Total = 0
  
  ' Set an initial variable for holding the opening and closing stock price for that year
  Dim Open_Stock As Double
  Dim Close_Stock As Double

  ' Set an initial variable for holding the yearly stock price change and percentage change
  Dim Yearly_Delta As Double
  Dim Yearly_Percent_Delta As Double
  Yearly_Delta = 0
  Yearly_Percent_Delta = 0

  ' Variable used to keep track of each stock's calculated values
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ' Variable used to loop through every row on the sheet
  Dim Lastrow As Long
  Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  


  ' Begin the loop for each sheet
  For j = 1 To WS_Count
    
    'New Column Headers for each calculated variable
    Worksheets(j).Range("I1").Value = "Ticker"
    Worksheets(j).Range("J1").Value = "Total Stock Volume"
    Worksheets(j).Range("K1").Value = "Yearly Change"
    Worksheets(j).Range("L1").Value = "Percent Change"

    ' Resets the summary table counter back to initial value so each summary table starts at same position per sheet
    Summary_Table_Row = 2
     
      ' Loop through every stock per sheet
      For i = 2 To Lastrow
      
        ' Sets the stock names to variables
        rCell = Worksheets(j).Cells(i, 1).Value
        bCell = Worksheets(j).Cells(i + 1, 1).Value
        
        ' Check if the stock ticker names match
        If rCell = bCell Then
          
          ' Sums and stores the stock volume per stock
          Vol_Total = Vol_Total + Worksheets(j).Cells(i, 7).Value
         
            ' Checks if there is a value for open stock value
            If Open_Stock = 0 Then
              ' Add to the opening stock value if it is blank
              Open_Stock = Worksheets(j).Cells(i, 3).Value
            End If
        
        ' Checks if the stock ticker names do not match
        ElseIf rCell <> bCell Then
        
          ' Set the Stock name
          Stock_Name = rCell

          ' Adds final stock volume to the Volume Total
          Vol_Total = Vol_Total + Worksheets(j).Cells(i, 7).Value

          ' Stores the closing stock value for the year
          Close_Stock = Worksheets(j).Cells(i, 6).Value

            ' Checks if the open_stock value is 0 to prevent errors when calculating the percentage
            If Open_Stock = "0" Then
              Yearly_Percent_Delta = 0
            Else
              'Calucates the stock price change and percent change
              Yearly_Delta = Close_Stock - Open_Stock
              Yearly_Percent_Delta = Yearly_Delta / Open_Stock
            End If
          
          ' Prints the stock name, volume total, yearly change, and percent change on the summary table per worksheet
          Worksheets(j).Range("I" & Summary_Table_Row).Value = Stock_Name
          Worksheets(j).Range("J" & Summary_Table_Row).Value = Vol_Total
          Worksheets(j).Range("K" & Summary_Table_Row).Value = Yearly_Delta
          Worksheets(j).Range("L" & Summary_Table_Row).Value = Yearly_Percent_Delta
          
          ' Add one to the summary table row for next stocks calcualted volume
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the variables used to calculate values
          Vol_Total = 0
          Open_Stock = 0
          Close_Stock = 0
          Yearly_Delta = 0
          Yearly_Percent_Delta = 0


        End If
         
      Next i
      
  'Declare rg as range object
  Dim rg As Range
  
  ' Stores the amont of rows on summary length variable
  Dim Summary_Table_Len As Long
  Summary_Table_Len = Worksheets(j).Cells(Rows.Count, "I").End(xlUp).Row
  
  ' Sets 3 conditions as format condition objects
  Dim cond1 As FormatCondition, cond2 As FormatCondition, cond3 As FormatCondition
  'Sets the range for each worksheet summary to apply conditional formatting over
  Set rg = Worksheets(j).Range("K2" & ":" & "K" & Summary_Table_Len)
 
  'Clear any existing conditional formatting
  rg.FormatConditions.Delete
 
  ' Define the rule for each conditional format
  Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "=0")
  Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")
  Set cond3 = rg.FormatConditions.Add(xlCellValue, xlEqual, "=0")
 
  ' Define the format applied for each conditional format
  With cond1
  .Interior.Color = vbGreen
  .Font.Color = vbBlack
  End With
 
  With cond2
  .Interior.Color = vbRed
  .Font.Color = vbBlack
  End With
 
  With cond3
  .Interior.Color = vbGreen
  .Font.Color = vbBlack
  End With
  
  'Define Variables to calculate largerst and smallest percentage chage in stock value as well as largest stock value per year
  Dim Max_Percent_Inc As Double
  Dim Max_Percent_Dec As Double
  Dim Max_Vol As Double
  Dim Tick As String
  Dim Addr As String
  
  ' Initialize row and coloumn headers
  Worksheets(j).Range("O2").Value = "Max % Increase"
  Worksheets(j).Range("O3").Value = "Greatest % Decrease"
  Worksheets(j).Range("O4").Value = "Greatest Total Volume"
  Worksheets(j).Range("P1").Value = "Ticker"
  Worksheets(j).Range("Q1").Value = "Value"
  
  'Calculates max and min percentage changes as well as the largest stock volume total
  Max_Percent_Inc = Worksheets(j).Application.WorksheetFunction.Max(Worksheets(j).Range("L2" & ":" & "L" & Summary_Table_Len))
  Max_Percent_Dec = Worksheets(j).Application.WorksheetFunction.Min(Worksheets(j).Range("L2" & ":" & "L" & Summary_Table_Len))
  Max_Vol = Worksheets(j).Application.WorksheetFunction.Max(Worksheets(j).Range("J2" & ":" & "J" & Summary_Table_Len))
  
  ' Loop used to find the associated ticker name to calculated values
  For Each c In Worksheets(j).Range("J2" & ":" & "L" & Summary_Table_Len)

        If c.Value = Max_Percent_Inc Then
            Worksheets(j).Range("Q2").Value = c.Value
            Worksheets(j).Range("P2").Value = c.Offset(0, -3).Value
        ElseIf c.Value = Max_Percent_Dec Then
            Worksheets(j).Range("Q3").Value = c.Value
            Worksheets(j).Range("P3").Value = c.Offset(0, -3).Value
        ElseIf c.Value = Max_Vol Then
            Worksheets(j).Range("Q4").Value = c.Value
            Worksheets(j).Range("P4").Value = c.Offset(0, -1).Value
        End If
  
  Next c
  
  
  'Formats each percentage column to percentage values
  Worksheets(j).Range("L:L").NumberFormat = "0.00%"
  Worksheets(j).Range("Q2:Q3").NumberFormat = "0.00%"
  
  Next j
 
  

End Sub
