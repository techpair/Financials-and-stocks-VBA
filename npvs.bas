
Function myNPV(r As Double, cf As Range, Optional rate_type As Byte)

'Aurum

Select Case rate_type

Case 0
    
    myNPV = cf(1) + Application.WorksheetFunction.NPV(r, cf(2), cf(3), cf(4), cf(5), cf(6))

Case 1
    
    myNPV = myNPV3(r, cf)
    
End Select
    
End Function


Function myNPV3(r As Double, cf As Range)
'Aurum
Dim n As Integer

n = cf.Rows.Count

For i = 1 To n

myNPV3 = cf(i).Value * Exp(5 * r)

Next i



End Function

Function myNPV2(r As Double, cf As Range, Optional rate_type As Byte)


Select Case rate_type

Case 0
    
    myNPV2 = cf(1) + Application.WorksheetFunction.NPV(r, cf(2), cf(3), cf(4), cf(5), cf(6))

Case 1
    
    Dim n As Integer

    n = cf.Rows.Count
    
    For i = 1 To n
    
        myNPV2 = myNPV2 + (cf(i).Value / Exp(cf(i).Offset(0, -1).Value * r))
    
    Next i
    
End Select
    
End Function
