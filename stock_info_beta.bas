Option Base 1
Function Stock_Info(STOCK_RET As Range, RF As Single, CHOICE, Optional MKT_RET)

Dim q1 As Worksheet

Set q1 = ThisWorkbook.Sheets("Q1")

With q1
    

    Select Case CHOICE
    
    Case "Sharpe"
           
           'Sharpe Ratio Calculation
            Dim avg As Double, std As Double
            
            avg = Application.WorksheetFunction.Average(STOCK_RET)
            std = Application.WorksheetFunction.StDev(STOCK_RET)
            
            Stock_Info = (avg - RF) / std  'formula for Sharpe
    
    Case "Beta"
        
           Stock_Info = Application.WorksheetFunction.Slope(STOCK_RET, MKT_RET)
       
    Case "Alpha"
        
            Stock_Info = Application.WorksheetFunction.Intercept(STOCK_RET, MKT_RET)
       
    End Select

End With

End Function


Function Beta(NTFX As Range, MKT_RET As Range)



Beta = Application.WorksheetFunction.Slope(NTFX, MKT_RET)



End Function

