Option Compare Database

Public Function CalcularVencimento(dia As Integer, Optional MES As Integer, Optional ANO As Integer) As Date

If Month(Now) = 2 Then
    If dia = 29 Or dia = 30 Or dia = 31 Then
        dia = 1
        MES = MES + 1
    End If
End If

If MES > 0 And ANO > 0 Then
    CalcularVencimento = Format((DateSerial(ANO, MES, dia)), "dd/mm/yyyy")
ElseIf MES = 0 And ANO > 0 Then
    CalcularVencimento = Format((DateSerial(ANO, Month(Now), dia)), "dd/mm/yyyy")
ElseIf MES = 0 And ANO = 0 Then
    CalcularVencimento = Format((DateSerial(Year(Now), Month(Now), dia)), "dd/mm/yyyy")
End If

End Function
