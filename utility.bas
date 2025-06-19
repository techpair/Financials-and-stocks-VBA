Function getformula(r As Range) As String
   Application.Volatile
   If r.HasArray Then
   getformula = "<-- " & " {" & r.FormulaArray & "}"
   Else
   getformula = "<-- " & " " & r.FormulaArray
   End If
End Function
Function ggetformula(r As Range) As String
   Application.Volatile
   If r.HasArray Then
   ggetformula = " {" & r.FormulaArray & "}"
   Else
   ggetformula = r.FormulaArray
   End If
End Function

