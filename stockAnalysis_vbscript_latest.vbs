' Function that performs following calculation
' Calculates total stock volume by ticker
' Calculates Yearlychange (Close - open), % Percent change for each ticker
' Finds out greatest % increase, greatest % decrease and Greatest total stock volume
' 2D arrays will be used to peform all calculations and stored in the result range
'**********************************************************************************************'
Sub Main()
    
    Debug.Print Now()
    
    '--------------------------------------------------------------------------------
    ' BEGIN - Variable Declaration
    '--------------------------------------------------------------------------------
    'Declare worksheet variables
    Dim wb As Workbook
    Dim aws As Excel.Worksheet
    Dim last_row, last_Col As String
    
    'Delcare data variables
    Dim next_tkr As String
    Dim i, j, cntr, open_date, close_date As Long
    Dim volSum, open_val, close_val, percent_chng As Double
    Dim set_open_val As Boolean
    
    ' Declare arrays
   ' Dim source_data, tmp, results, min_max As Variant
    
    '--------------------------------------------------------------------------------
    ' END - Variable Declaration
    '--------------------------------------------------------------------------------
        
    '--------------------------------------------------------------------------------
    ' BEGIN - Analysis of each worksheet in the workbook
    '--------------------------------------------------------------------------------
    For Each aws In ThisWorkbook.Worksheets
    
        'Set aws = ThisWorkbook.ActiveSheet ' Only for testing
        
        ' Declare arrays -- !!!! IMPORTANT: TO ENABLE RE-INITIALIZATION !!!!
        Dim source_data, tmp, results, min_max As Variant
    
        '--------------------------------------------------------------------------------
        ' BEGIN - Variable Initialization & default value setting
        '--------------------------------------------------------------------------------
        
        'initialize counters and default variables
        
        cntr = 1 ' as 0 is for header
        volSum = 0 ' initialize it to zero
        percent_chng = 0
        set_open_val = True ' Indicator to indicate that open value of a ticker needs to be captured
        next_tkr = ""
        
        
        ' initialize min-max array.
        ' This array will store greatest % increase, % decrease and total stock volume
            
        min_max = Array(Array("", "Ticker", "Value"), _
                        Array("Greatest % Increase", 0, 0), _
                        Array("Greatest % Decrease", 0, 0), _
                        Array("Greatest Total Volume", 0, 0))
                        
        With WorksheetFunction
            min_max = .Transpose(.Transpose(min_max))
        End With
        
        '--------------------------------------------------------------------------------
        ' END - Variable Initialization & default value setting
        '--------------------------------------------------------------------------------
                
        '--------------------------------------------------------------------------------
        ' BEGIN - Sort & Copy source data range into a 2-D array
        '--------------------------------------------------------------------------------
            With aws
                .Activate
                
                With .Range("A1").CurrentRegion
                    
                    .Select ' Activate current source data region
                    
                    'get last row and col from current worksheet
                    
                     last_row = .Rows.Count
                     last_Col = .Columns.Count
                     
                    ' sort the range data by ticker and data in ascending order
                    .Sort key1:="<ticker>", order1:=xlAscending, Key2:="<date>", order2:=xlAscending, _
                                    Header:=xlYes
                    
                    'get the unformatted data (use value2) from the source range into an array
                    source_data = .Value2
                    
               End With
               
            End With
            
        '--------------------------------------------------------------------------------
        ' END - Sort & Copy source data range into a 2-D array
        '--------------------------------------------------------------------------------
        
        
        '------------------------------------------------------------------
        ' BEGIN : Calculate
        '   a. Total volume by ticker
        '   b. Yearly change from what the stock opened the year at to what the closing price was.
        '   c. The percent change from the what it opened the year at to what it closed.
        '
        ' column index to header row for ref:
        ' <ticker>    <date>  <open>  <high>  <low>   <close> <vol>
        '    1          2       3       4       5       6       7
        '------------------------------------------------------------------
    
        For i = 2 To UBound(source_data, 1) ' ignore row 1 as it is header row
                    
            If (i = UBound(source_data, 1)) Then
                next_tkr = ""
            Else
                next_tkr = source_data(i + 1, 1)
            End If
            
            ' Check if current row and next row refer to same ticker name
            ' If YES, just add sum up the total stock volume
            ' if NO, calculate summary stats for the current ticker
                   
            If (source_data(i, 1) <> next_tkr) Then
                
                ' Calculate total stock volume
                volSum = volSum + source_data(i, 7)
                
                ' get close date value to calculate yearly change and % change
                close_val = source_data(i, 6)
                    
                
                ' add the ticker, vol sum, yrly_chng (close - open), %change [(close - open)/open]
                ' to another summary array
                
                ' redimension the results array based on empty or not empty
                If (IsEmpty(results)) Then
                    
                    ' if array is empty, add the headers
                     ReDim results(0 To 3, 0 To 1)
                     results(0, 0) = "Ticker"
                     results(1, 0) = "Yearly Change"
                     results(2, 0) = "Percent Change"
                     results(3, 0) = "Total Stock Volume"
                     
                Else
                
                    ' if not empty redimension results arr with new row with preserve data option
                    ReDim Preserve results(UBound(results, 1), LBound(results, 2) To (UBound(results, 2) + 1))
                
                End If
                
                If (open_val = 0) Then
                    percent_chng = 0
                Else
                    percent_chng = ((close_val - open_val) / open_val)  ' calc % change
                End If
                
                results(0, cntr) = source_data(i, 1)                ' ticker value
                results(1, cntr) = close_val - open_val             ' yearly change
                results(2, cntr) = percent_chng                     ' % change
                results(3, cntr) = volSum                           'Total stock volume
                
                                
                '**** Begin - calculation for Greatest % increase, %decrease & total stock vol table **********
                
                If (min_max(4, 3) < volSum) Then
                    min_max(4, 3) = volSum             ' replace the greatest total stock volume
                    min_max(4, 2) = source_data(i, 1)   'replace ticker name too
                End If
                
                If (min_max(3, 3) > percent_chng) Then
                    ' this means current % change is lowest, so replace it as greatest % decrease
                    min_max(3, 3) = percent_chng
                    min_max(3, 2) = source_data(i, 1)
                End If
                
                If (min_max(2, 3) < percent_chng) Then
                    ' this means current % change is lowest, so replace it as greatest % increase
                    min_max(2, 3) = percent_chng
                    min_max(2, 2) = source_data(i, 1)
                End If
                
              
                '**** End - calculation for Greatest % increase, %decrease & total stock vol table *********
                
                ' reset all cntrs, values and indicators
                volSum = 0
                cntr = cntr + 1
                open_val = 0
                close_val = 0
                set_open_val = True
                
            Else
                If (set_open_val And source_data(i, 3) <> 0) Then
                    open_val = source_data(i, 3)    ' capture the stock open value in a variable (can be anytime in the yr)
                    set_open_val = False            ' Set indicator to false to ignore other days stock value
                    
                End If
                
                volSum = volSum + source_data(i, 7)
            End If
            
            
        Next i
        
        '------------------------------------------------------------------
        ' END : ************ All calculation done ***********
        '------------------------------------------------------------------
        
        '---------------------------------------------------------------------------------------------
        ' BEGIN - Add the summary of Yrly change, % change, Total stock volume by Ticker to Active sheet
        '---------------------------------------------------------------------------------------------
        
        'transpose the array to add to the range for better viewing in excel
        ReDim tmp(UBound(results, 2), UBound(results, 1))
        For i = 0 To UBound(results, 2)
            For j = 0 To UBound(results, 1)
                tmp(i, j) = results(j, i)
            Next j
        Next i
        
        ' Select the range of cells that need to be populated dynamically
        With aws.Range(Cells(1, last_Col + 2), Cells(UBound(tmp, 1) + 1, last_Col + UBound(tmp, 2) + 2))
            
            ' set range active and copy array to the range without any formatting
            .Select
            .Value2 = tmp ' copy
            
            ' BEGIN - Formatting
            
            ' add the conditional formating to Yearly Change column
            With .Columns(2)
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                        Formula1:="0.00"
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                        Formula1:="0.00"
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                        Formula1:="=0.00"
               
                .FormatConditions(1).Interior.ColorIndex = 3 			' Red for values < 0
                .FormatConditions(1).StopIfTrue = False
                .FormatConditions(2).Interior.ColorIndex = 4			' Green for values > 0
                .FormatConditions(2).StopIfTrue = False
                .FormatConditions(3).Interior.ColorIndex = 4			' Green for values = 0
                .FormatConditions(3).StopIfTrue = False
            
            End With

            .Columns(2).NumberFormat = "0.00000000#" 	' number format for yrly Change cells
            .Columns(3).NumberFormat = "0.00%" 			' % format for % Change cells
            
            ' Clear any formats set to the header by conditional formating
            .Rows(1).ClearFormats
            
            .Columns.AutoFit               ' Autofit for clear presentation
        
        End With
        '---------------------------------------------------------------------------------------------
        ' END - Add the summary of Yrly change, % change, Total stock volume by Ticker to Active sheet
        '---------------------------------------------------------------------------------------------
        
        '---------------------------------------------------------------------------------------------
        ' BEGIN - Copy > % Increase, >% decrease, > total stock volume to worksheet
        '---------------------------------------------------------------------------------------------
        
        With aws.Range(Cells(1, last_Col + UBound(tmp, 2) + 4), _
                    Cells(UBound(min_max, 1), last_Col + UBound(tmp, 2) + UBound(min_max, 2) + 3))
                    
            ' Select copying range and copy array to selected range
            .Select
            .Value2 = min_max
            
            ' Apply formatting as required, finally autofit for better presentation
            .Columns(3).NumberFormat = "0.00%"
            .Cells(4, 3).NumberFormat = "0"
            .Rows.AutoFit
            .Columns.AutoFit
                    
        End With
        
        '---------------------------------------------------------------------------------------------
        ' END - Copy > % Increase, >% decrease, > total stock volume to woorksheet
        '---------------------------------------------------------------------------------------------
        
        'clear & re-initialize all arrays
        Erase source_data
        Erase tmp
        results = Empty
        Erase min_max
       
    
    Next aws
    
    MsgBox "!!! Analysis Complete !!!"
            
    Debug.Print Now()
    
End Sub
