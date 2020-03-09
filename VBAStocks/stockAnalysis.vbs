Sub stockAnalysis()
'---------------Program Purpose - Stock Analysis for each sheet and generate Summary Report-------------------------------
'Create a script that will loop through all the stocks for one year and output the following information.
' - The ticker symbol.
' - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' - The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' - The total stock volume of the stock.
' - You should also have conditional formatting that will highlight positive change in green and negative change in red.
'Note: Summary headers and formatting is done programmatically. Focus is set back to the first worksheet.
'--------------------------------------------------------------------------------------------------------------------------
  'Variable declaration
    Dim I As Double
    Dim lastRow As Double
    Dim workSheetName As String
    Dim firstSheetName As String
    Dim tickerSymbol As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearChange As Double
    Dim stockVolume As Double
    Dim Summary_Table_Row As Double
    Dim grtPercentageIncrease As Double
    Dim grtPercentageDecrease As Double
    Dim grtVolumne As Double
    Dim bBeginingTicker As Boolean
    Dim bFirstSheet As Boolean
    
  'Declare constants
    Const COLOR_GREEN As Integer = 4
    Const COLOR_RED As Integer = 3

   'Intialize variables
     bFirstSheet = True

      'Loop through all worksheets or tabs
        For Each Sheet In Worksheets
          
          'Initiate variables for the sheet
            lastRow = 0
            bBeginingTicker = True
            Summary_Table_Row = 2
                
          'Active worksheet name
            workSheetName = Sheet.Name
            
          'Get the first sheet name - this will be used at the end of processing to activate so focus goes back to first sheet
            If bFirstSheet = True Then
                firstSheetName = Sheet.Name
            End If
            
          'Preparare column headings for first and second summary tables and columm cell formatting for % and rounding cell values
            initSummarySheet (workSheetName)
            
          'Activate current worksheet
            Worksheets(workSheetName).Activate
            
          'Get last row of the active worksheet
            lastRow = Cells(Rows.Count, "A").End(xlUp).Row
              
          'Loop through ticker symbols and get unique values populated in column I
            For I = 2 To lastRow
            
              'Get the stock open price for first counter
                If bBeginingTicker = True Then
                    openPrice = Cells(I, 3).Value
                End If
                
              'Check if row value is changed
                If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
  
                    'Get sticker symbol
                      tickerSymbol = Cells(I, 1).Value
                    
                    'Get the stock close price
                      closePrice = Cells(I, 6).Value
                      
                    'Calculate Yearly Change
                      yearChange = closePrice - openPrice
                     
                    'Calculate total stock volume
                      stockVolume = stockVolume + Cells(I, 7).Value
                      
                    '------ Beggining to populate first summary table ---------------------------------------------------------
                        
                        'Populate ticker symbol
                          Cells(Summary_Table_Row, 9).Value = tickerSymbol
                       
                        'Populate yearly change
                          Cells(Summary_Table_Row, 10).Value = yearChange
                        
                         'Populate percentage change
                           If (yearChange = 0 Or openPrice = 0) Then
                              Cells(Summary_Table_Row, 11).Value = 0
                           Else
                              Cells(Summary_Table_Row, 11).Value = yearChange / openPrice
                           End If
                      
                         'Populate total stock volume
                           Cells(Summary_Table_Row, 12).Value = stockVolume
                        
                         'Change the cell color
                           If yearChange >= 0 Then
                             'Fill cell with green color
                               Cells(Summary_Table_Row, 10).Interior.ColorIndex = COLOR_GREEN
                           Else
                             'Fill cell with red color
                               Cells(Summary_Table_Row, 10).Interior.ColorIndex = COLOR_RED
                           End If
                           
                       '------------end of populating first summary table for the iteration ---------------------------------------------
                    
                   'Increment summary table row value
                     Summary_Table_Row = Summary_Table_Row + 1
                     
                   'Reset openPrice, ClosePrice,stockVolume and biginning ticket flag
                      openPrice = 0
                      closePrice = 0
                      yearChange = 0
                      stockVolume = 0
                      bBeginingTicker = True
           
                Else
                    'Calculate total stock volumne
                      stockVolume = stockVolume + Cells(I, 7).Value
                    
                    'Reset biginning stock flag
                      bBeginingTicker = False
                End If
                
          'Continue interartion
            Next I
          
          '----------- At this time, program is completed processing data for the first summary table -----------------------
        
        'Its time to populate second summary table
        
        'Populate greatest % increase
            'Get greates % increase - call to function getMinMaxSummaryValues with
            ' parameters-sheet name, column name and criteria="Max"
              grtPercentageIncrease = getMinMaxSummaryValues(workSheetName, "K", "Max")
            
            'Get ticker Symbol for the greatest % increase value - call to function getStockTicker with
            ' parameters-sheet name, column name and greatest % increase value
              tickerSymbol = getStockTicker(workSheetName, "K", grtPercentageIncrease)

            'Populate worksheet columns for greatest % increase and Stock Ticker values
              Range("P2").Value = tickerSymbol
              Range("Q2").Value = grtPercentageIncrease
         
            'Get greates % decrease - call to function getMinMaxSummaryValues with
            ' parameters-sheet name, column name and criteria="Min"
              grtPercentageDecrease = getMinMaxSummaryValues(workSheetName, "K", "Min")
           
            'Get ticker Symbol for the greatest % decrease value - call to function getStockTicker with
            ' parameters-sheet name, column name and greatest % decrease value
              tickerSymbol = getStockTicker(workSheetName, "K", grtPercentageDecrease)
            
            'Populate worksheet columns for greatest % decrease and Stock Ticker values
              Range("P3").Value = tickerSymbol
              Range("Q3").Value = grtPercentageDecrease
         
            'Get greates stock volume - call to function getMinMaxSummaryValues with
            ' parameters-sheet name, column name and criteria="Max"
              grtVolume = getMinMaxSummaryValues(workSheetName, "L", "Max")
            
            'Get ticker Symbol for the greatest volume - call to function getStockTicker with
            ' parameters-sheet name, column name and greatest volume
              tickerSymbol = getStockTicker(workSheetName, "L", grtVolume)
            
            'Populate worksheet columns for greatest stock volume and Stock Ticker values
              Range("P4").Value = tickerSymbol
              Range("Q4").Value = grtVolume
              
              'Set the first sheet flag as false - proceding to second worksheet
                bFirstSheet = False
  
        Next Sheet
 
 'Bring back focus to first sheet
    'Activate first worksheet
      Worksheets(firstSheetName).Select
      
 'Processing complete - Happy Programming
   MsgBox ("Thank you for using Stock Analysis Program - Happy Programming!")
   
 
End Sub

'-----------------------------------------------------------------------------------------------------------------------------
' Function Name: getStockTicker
' Function Parameters: Excel sheet name: pSheetName as string, Excel Column: iColName as string, Cell Value: iValue as double
' Function Purpose: This function will match the cell value and return the corresponding ticker symbol
'------------------------------------------------------------------------------------------------------------------------------
Function initSummarySheet(pSheetName) As String
  
  'variable declaration
   Dim ws As Worksheet
   
   'Set current worksheet and activate it
     Set ws = ThisWorkbook.Sheets(pSheetName)
     Worksheets(pSheetName).Activate

        'Populate first summary sheet column names
          Range("I1").Value = "Ticker"
          Range("J1").Value = "Yearly Change"
          Range("K1").Value = "Percentage Change"
          Range("L1").Value = "Total Stock Volume"
                 
        'Format first-summary table columns
          Range("J:J").NumberFormat = "0.00"
          Range("K:K").NumberFormat = "0.00%"
          Range("L:L").NumberFormat = "0,00"
          
                
        'Populate second summary sheet column names
          Range("P1").Value = "Ticker"
          Range("Q1").Value = "Value"
          Range("O2").Value = "Greatest % Increase"
          Range("O3").Value = "Greatest % Decrease"
          Range("O4").Value = "Greatest Total Volume"
                
        'Format second summary table
          Range("Q2").NumberFormat = "0.00%"
          Range("Q3").NumberFormat = "0.00%"
          Range("Q4").NumberFormat = "0.0000E+00" 'My preferance - "0,00"
          
        'Auto-fit columns
          Range("I:I").EntireColumn.AutoFit
          Range("J:J").EntireColumn.AutoFit
          Range("K:K").EntireColumn.AutoFit
          Range("L:L").EntireColumn.AutoFit
          Range("O:O").EntireColumn.AutoFit
          Range("P:P").EntireColumn.AutoFit
          Range("Q:Q").EntireColumn.AutoFit
          
End Function

'-----------------------------------------------------------------------------------------------------------------------------
' Function Name: getMinMaxSummaryValues
' Function Parameters: Excel sheet name: pSheetName as string, Excel Column: iColName as string, Criteria (Min or Max): iCriteria as String
' Function Purpose: This function will return minimum or maximum value for the given column range values.
'------------------------------------------------------------------------------------------------------------------------------

Function getMinMaxSummaryValues(pSheetName, iColName, iCriteria As String) As String
    
    ' Variable declaration
        Dim rCount As Integer
        Dim tmpValue As Double
        Dim currentValue As Double
        Dim ws As Worksheet
     
     'Set current worksheet and activate it
        Set ws = ThisWorkbook.Sheets(pSheetName)
        Worksheets(pSheetName).Activate

     'Get the column row count
        rCount = Range("K" & Rows.Count).End(xlUp).Row

      'Get the first value of cell
        tmpValue = Range(iColName & "2")
    
      'Check conditioncriteria  - Max or Min
        If iCriteria = "Max" Then
        
            'Loop through fist summary table to get max value from the give column range
                For I = 3 To rCount
                    
                    'Get the cell value to compare
                      currentValue = Range(iColName & I).Value
                    
                    'Comapre two values and reset value of tmpValue variable to max value
                      If (currentValue > tmpValue) Then
                        tmpValue = currentValue
                      End If
                
                'Go to next row
                Next I
        
        ElseIf iCriteria = "Min" Then

              'Loop through fist summary table to get min value from the give column range
                For I = 3 To rCount
                   
                   'Get the cell value to compare
                     currentValue = Range(iColName & I).Value
                    
                   'Comapre two values and reset value of tmpValue variable to min value
                       If (currentValue < tmpValue) Then
                        tmpValue = currentValue
                    End If
                
                Next I
        Else
                'Msg - wrong criteria passed to the function
                 MsgBox ("Only Max or Min value is allowed for the fucntion criteria")
                
                'Set tem value as black as there is no Max or Min value that can be found
                 tmpValue = ""
   End If
   
   'Return Max or Min value for the criteria
    getMinMaxSummaryValues = tmpValue

End Function
'-----------------------------------------------------------------------------------------------------------------------------
' Function Name: getStockTicker
' Function Parameters: Excel sheet name: pSheetName as string, Excel Column: iColName as string, Cell Value: iValue as double
' Function Purpose: This function will match the cell value and return the corresponding ticker symbol
'------------------------------------------------------------------------------------------------------------------------------
Function getStockTicker(pSheetName, iColName, iValue) As String
  
  'Variable declaration
   Dim rCount As Long
   Dim ws As Worksheet
   Dim tmpStickerValue As String
   
   'Set current worksheet and activate it
     Set ws = ThisWorkbook.Sheets(pSheetName)
     Worksheets(pSheetName).Activate

    'Get the row count for the given column from the parameters
     rCount = Range(iColName & Rows.Count).End(xlUp).Row
    
    'Loop through fist-summary table to check cell value and get corresponding ticker symbol
        For I = 2 To rCount
                       
            'Check if current cell value and parameter value is same. Round the values for compare
              If (Round(Range(iColName & I).Value, 2) = Round(iValue, 2)) Then
                 'Set ticker symbol value based on the cell value
                  tmpStickerValue = Range("I" & I).Value
                  
                  'Exit for loop as ticker symbol value is captured
                   Exit For
              End If
              
        'Go to next row
        Next I
        
 'Set the ticker value to return to function call
   getStockTicker = tmpStickerValue

End Function
