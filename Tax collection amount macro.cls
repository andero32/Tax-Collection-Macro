Sub CollectTax()
    ' Declare variables
    Dim TotalIncome As Double
    Dim TaxRate As Double
    Dim TotalTax As Double
    
    ' Get total income from user
    TotalIncome = InputBox("Enter total income:")
    
    ' Determine tax rate based on income
    If TotalIncome < 10000 Then
        TaxRate = 0
    ElseIf TotalIncome < 50000 Then
        TaxRate = 0.1
    ElseIf TotalIncome < 100000 Then
        TaxRate = 0.2
    Else
        TaxRate = 0.3
    End If
    
    ' Calculate total tax
    TotalTax = TotalIncome * TaxRate
    
    ' Display total tax
    MsgBox "Total tax: " & TotalTax
End Sub

