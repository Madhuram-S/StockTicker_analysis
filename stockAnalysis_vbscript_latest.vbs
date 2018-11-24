' Function that performs following calculation
' Calculates total stock volume by ticker
' Calculates Yearlychange (Close - open),
' % Percent change for each ticker ((Close - Open) / Open) *100
' Finds out greatest % increase, greatest % decrease and Greatest total stock volume
' 2D arrays will be used to peform all calculations and then stored in the result range
'**********************************************************************************************'
Sub Main_AnalyzeStock()
    
    '--------------------------------------------------------------------------------
    ' BEGIN - Variable Declaration
    '--------------------------------------------------------------------------------
        Dim aws As Excel.Worksheet
        Dim next_tkr As String
        Dim i, cntr As Long
        Dim volSum, open_val, close_val, percent_chng As Double
        Dim set_open_val As Boolean
    '--------------------------------------------------------------------------------
    ' END - Variable Declaration
    '--------------------------------------------------------------------------------
        
    '--------------------------------------------------------------------------------
    ' BEGIN - Analysis of each worksheet in the workbook
    '--------------------------------------------------------------------------------
    For Each aws In ThisWorkbook.Worksheets
    
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

            'Initialize results array
            ReDim results(0 To 3, 0 To 1)
             results(0, 0) = "Ticker"
             results(1, 0) = "Yearly Change"
             results(2, 0) = "Percent Change"
             results(3, 0) = "Total Stock Volume"
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
                    
                    ' sort the range data by ticker and data in ascending order.
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
                
                If (open_val <> 0) Then
                    ' calc % change. Note : don't use *100. formatting will take care of it later
                    percent_chng = ((close_val - open_val) / open_val)  
                End If
                
                ' add the ticker, vol sum, yrly_chng (close - open), %change [(close - open)/open]
                ' to results array
                
                'redimension results arr with new row with preserve data option
                    ReDim Preserve results(UBound(results, 1), LBound(results, 2) To (UBound(results, 2) + 1))

                    results(0, cntr) = source_data(i, 1)                ' ticker value
                    results(1, cntr) = close_val - open_val             ' yearly change
                    results(2, cntr) = percent_chng                     ' % change
                    results(3, cntr) = volSum                           'Total stock volume
                
                ' reset all cntrs, values and indicators
                volSum = 0
                cntr = cntr + 1
                open_val = 0
                close_val = 0
                set_open_val = True
                percent_chng = 0
                
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
            tmp = WorksheetFunction.Transpose(results)
        
            ' The result range is stored in I:L range. select the range of cells that needs to be populated 
            With aws.Range("I1:L" & UBound(tmp, 1))
                
                ' set range active and copy array to the range without any formatting
                .Select
                .Value2 = tmp ' copy array to Range in active sheet
                
                ' BEGIN - Formatting
                ' add the conditional formating to Yearly Change column 
                ' Green color if value >= 0.00 and Red for value < 0.00

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

                '---------------------------------------------------------------------------------------------
                ' BEGIN - Copy > % Increase, >% decrease, > total stock volume to worksheet
                '---------------------------------------------------------------------------------------------
                    ' initialize min-max array.
                    ' This array will store greatest % increase, % decrease and total stock volume
                        
                    min_max = Array(Array("", "Ticker", "Value"), _
                                    Array("Greatest % Increase", "", 0), _
                                    Array("Greatest % Decrease", "", 0), _
                                    Array("Greatest Total Volume", "", 0))
                                    
                    With WorksheetFunction
                        min_max = .Transpose(.Transpose(min_max))
                    End With
                    
                    'Sort the results by % change desc. The first row will be %increase and last row will %decrease

                    .Sort key1:="Percent Change", order1:=xlDescending, Header:=xlYes
                    min_max(2, 3) = .Cells(2, 3).Value
                    min_max(2, 2) = .Cells(2, 1).Value
                    min_max(3, 3) = .Cells(.End(xlDown).Row, 3).Value
                    min_max(3, 2) = .Cells(.End(xlDown).Row, 1).Value
                    
                    'min_max(3, 3) = .Cells(UBound(tmp, 1) - 1, 3).Value
                    'min_max(3, 2) = .Cells(UBound(tmp, 1) - 1, 1).Value

                    .Sort key1:="Total Stock volume", order1:=xlDescending, Header:=xlYes
                    min_max(4, 3) = .Cells(2, 4).Value
                    min_max(4, 2) = .Cells(2, 1).Value
                    
                    'Reverse results range to sorted by Ticker asc.
                    .Sort key1:="Ticker", order1:=xlAscending, Header:=xlYes

                    'Select the range for adding % Increase, decrease and total vol (Range is N:P)
                    With aws.Range("N1:P" & UBound(min_max, 1))
                    
                        ' Select copying range and copy array to selected range
                        .Select
                        .Value2 = min_max
                        
                        ' Apply formatting as required, finally autofit for better presentation
                        .Columns(3).NumberFormat = "0.00%"
                        .Cells(4, 3).NumberFormat = "0"
                        .Columns.AutoFit
                            
                    End With
                '---------------------------------------------------------------------------------------------
                ' END - Complete calcuation and Copy > % Increase, >% decrease, > total stock volume to worksheet
                '---------------------------------------------------------------------------------------------
           
            End With
        '---------------------------------------------------------------------------------------------
        ' END - Add the summary of Yrly change, % change, Total stock volume by Ticker to Active sheet
        '---------------------------------------------------------------------------------------------
        
        'clear & re-initialize all arrays
        Erase source_data, tmp, min_max
        results = Empty
        
    Next aws
    
    MsgBox "!!! Analysis Complete !!!"
       
End Sub
