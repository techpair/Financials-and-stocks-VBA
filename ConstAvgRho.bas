Option Base 1

Function ConstantRhoHist(data As Range)

 rho = AvgRho(data)

    ' get number of assets
    n = data.Columns.Count
    
    ' declare var-cov matrix
    Dim vcov() As Double
    ReDim vcov(n, n)
    
    ' calculate var-cov matrix
    
    For i = 1 To n
        For j = 1 To n
        
        If i = j Then
            vcov(i, j) = Application.WorksheetFunction.Var_S(data.Columns(i))
            Else
                vcov(i, j) = rho _
            * Application.WorksheetFunction.StDev_S(data.Columns(i)) _
            * Application.WorksheetFunction.StDev_S(data.Columns(j))
        End If
        
        Next j
    Next i
    
    ConstantRhoHist = vcov
    
    End Function
    
    Function AvgRho(data As Range) As Double

     'get the number of assets
    n = data.Columns.Count
    
     'add up all the correlations
    total_rho = 0
    n_rho = 0
    
    For i = 1 To n
        For j = i + 1 To n
            
            rho_ij = Application.WorksheetFunction.Correl(data.Columns(i), data.Columns(j))
            total_rho = total_rho + rho_ij
            n_rho = n_rho + 1
            
            Next j
        Next i
        
         'return average correlations
         AvgRho = total_rho / n_rho
         
      End Function
         
    



