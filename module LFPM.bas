'Loan Fee Per Month
Function LFPM(Credit As Double, Procent As Double, Period As Double) As Double
    Dim p
    
    Application.Volatile
    
    If Procent < 1 Then
        LFPM = "Ошибка: процент должен быть больше единицы."
        
        Exit Function
    ElseIf Credit <= 0 Then
        LFPM = "Ошибка: номинальный кредит не может быть меньше или равен нулю."

        Exit Function
    ElseIf Period <= 0 Then
        LFPM = "Ошибка: срок кредита не может быть меньше или равен нулю."
    
        Exit Function
    Else
        p = Procent / (100 * 12)
    
        LFPM = Credit * (p + (p / ((1 + p) ^ (Period * 12) - 1)))
    End If
End Function
