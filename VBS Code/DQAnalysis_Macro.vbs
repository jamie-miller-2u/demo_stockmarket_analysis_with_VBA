Sub DQAnalysis()

    'set DQAnalysis as the active worksheet
    Worksheets("DQ Analysis").Activate
    
    ' assign a value to cell A1
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    ' assign additional values for row 3 columns 1, 2, 3
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    ' make assigned values bold
    Range("A1:C3").Select
    Selection.Font.Bold = True
    
    ' resize columns A to C
    Columns("A:C").EntireColumn.AutoFit
    
    ' ensure sheet 2018 is active
    Worksheets("2018").Activate
    
    ' inialize variables
    rowStart = 2
    
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
    totalVolume = 0
    
    ' iterate for ticker = "DQ"
    For i = rowStart To rowEnd
 
        ' Use conditional to check if current row ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
                    
            ' increase totalVolume for ticker "DQ"
            totalVolume = totalVolume + Cells(i, 8).Value
        
        End If

    Next i
    
    ' format
    Range("B4").Select
    Selection.NumberFormat = "#,##0"

        
        ' increase totalVolume
        totalVolume = totalVolume + Cells(i, 8).Value

        
        ' display resulting totalVolume on DQ_Analysis
        Worksheets("DQ Analysis").Activate
        
        Cells(4, 1).Value = 2018
        Cells(4, 2).Value = totalVolume
 
    
End Sub