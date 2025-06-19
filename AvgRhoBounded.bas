

Function AvgRhoBounded(RET, LB, UB)
    
    
    Dim q3  As Worksheet, totalColumns As Integer, columnCtr As Integer, secColumnCtr As Integer, _
        rjo_ij As Double, total_rho As Double, n_rho As Double, lastRowData As Long

    Set q3 = ThisWorkbook.Sheets("Q3")

    
    With q3
         
         'get last row
         lastRowData = .Cells(3, 2).End(xlDown).Row
         
         'get the number of assets
        
        totalColumns = q3.UsedRange.Columns.Count
        
         'add up all the correlations
        total_rho = 0
        n_rho = 0
        
        For columnCtr = 3 To totalColumns
            For secColumnCtr = columnCtr + 1 To totalColumns
                
                If .Cells(3, secColumnCtr).Value <> "" Then
                
                    rho_ij = Application.WorksheetFunction.Correl(.Range(.Cells(4, columnCtr), _
                                                                .Cells(lastRowData, columnCtr)), _
                                                                .Range(.Cells(4, secColumnCtr), _
                                                                .Cells(lastRowData, secColumnCtr)))
                    
                    
                    total_rho = total_rho + rho_ij
                    n_rho = n_rho + 1
                
                End If
                
                Next secColumnCtr
            Next columnCtr
    
    End With
    
         'return average correlations
   AvgRhoBounded = total_rho / n_rho
   
   
 
End Function

