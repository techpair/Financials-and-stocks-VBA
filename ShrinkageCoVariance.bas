Option Base 1
Function Shrinkage(ASSET_RETURNS As Range, CHOICE1, CHOICE2, LAMBDA As Single _
    , Optional MKT_RET = 0)
    
    Dim q5 As Worksheet, strFormula1 As String, strFormula2 As String, strFinalFormula As String, _
        plotRng As Range
    
    Set q5 = ThisWorkbook.Sheets("Q5")
    
    
 With q5

            n = ASSET_RETURNS.Columns.Count
            
            
            Dim matrix() As Double, tableStartRow As Long, table2StartRow As Long, table1RowCtr As Long, _
                table2RowCtr As Long, tableColCtr As Long, tableStartCol As Long
     
            
             Select Case CHOICE1
                Case "sample"
                    tableStartRow = 94
                Case "SIM"
                    tableStartRow = 82
                Case "constant-corr"
                    tableStartRow = 107
            End Select
            
            Select Case CHOICE2
                Case "sample"
                    table2StartRow = 94
                Case "SIM"
                    table2StartRow = 82
                Case "constant-corr"
                    table2StartRow = 107
            End Select
            
            tableStartCol = 14
            
            ReDim matrix(n, n)
            
                For i = 1 To n
                    For j = 1 To n
                    '=(N94*$C$79) + ((1-$C$79)*N82)
                    
                    matrix(i, j) = (.Cells(tableStartRow + table1RowCtr, tableStartCol + tableColCtr).Value * LAMBDA) + _
                        ((1 - LAMBDA) * .Cells(table2StartRow + table2RowCtr, tableStartCol + tableColCtr).Value)
                    tableColCtr = tableColCtr + 1
                    
                    Next j
                    
                    tableColCtr = 0
                    table1RowCtr = table1RowCtr + 1
                    table2RowCtr = table2RowCtr + 1
                Next i
            
            Shrinkage = matrix
                 
    End With
End Function

Function VarCovar(rng As Range) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim numcols As Integer
    numcols = rng.Columns.Count
    numrows = rng.Rows.Count
    Dim matrix() As Double
    ReDim matrix(numcols, numcols)

    For i = 1 To numcols
        For j = 1 To numcols
            matrix(i, j) = Application.WorksheetFunction.Covariance_S(rng.Columns(i), rng.Columns(j))
        Next j
    Next i
    VarCovar = matrix
End Function

Sub PlotTablesQ5()
    
        Dim startRow As Long, lastRow As Long, q5 As Worksheet, _
        strFormula1 As String, strFormula2 As String, strFinalFormula As String, plotRng As Range
    
    Set q5 = ThisWorkbook.Sheets("Q5")
    
    
     With q5
        
        startRow = 4
        lastRow = .Cells(4, 2).End(xlDown).Row
    
    'update SIM table for calculations
        .Range("N82:R86").FormulaArray = "=sim(C" & startRow & ":G" & lastRow & ",H" & startRow & ":H" & lastRow & ")"
    
        'Update Sample for calculations
        .Range("N94:R98").FormulaArray = "=VarCovar(C" & startRow & ":G" & lastRow & ")"
    
        'Update const-corr for calculations
         .Range("N107:R111").FormulaArray = "=ConstantRhoHist(C" & startRow & ":G" & lastRow & ")"
    End With
End Sub

Function SIM(assetdata As Range, marketdata _
As Range) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim numcols As Integer
    numcols = assetdata.Columns.Count
    Dim matrix() As Double
    ReDim matrix(numcols, numcols)
    
    For i = 1 To numcols
    For j = 1 To numcols
        If i = j Then
     
        
        matrix(i, j) = Application. _
        WorksheetFunction.Var_S(assetdata.Columns(i))
        
        Else
        
        matrix(i, j) = _
        Application.WorksheetFunction. _
        Slope(assetdata.Columns(i), marketdata) * _
        Application.WorksheetFunction. _
        Slope(assetdata.Columns(j), marketdata) * _
        Application.WorksheetFunction. _
        Var_S(marketdata)
        
       
        
    End If
    Next j
    Next i
    SIM = matrix
End Function
