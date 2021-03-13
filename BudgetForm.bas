Private NumBoxCollection As Collection
Public Sheet As Worksheet

Private Sub OkButton_Click()
    findAndSetValueOnSheet "АУП", val(AUPRateBox.value) / 100#
    findAndSetValueOnSheet "НР", val(NaklRateBox.value) / 100#
    findAndSetValueOnSheet "НДС к уплате в бюджет", val(NDSRateBox.value) / 100#
    findAndSetValueOnSheet "Налог на прибыль", val(NalogRateBox.value) / 100#
    findAndSetValueOnSheet "Чистая прибыль", val(ProfitRateBox.value) / 100#
    
    BudgetForm.Hide
End Sub

Private Sub findAndSetValueOnSheet(name As String, value)
    For i = 1 To 9
        If Sheet.Cells(i, 1).Text = name Then
            Sheet.Cells(i, 2).value = value
            Exit Sub
        End If
    Next i
End Sub

Private Sub UserForm_Activate()
    AUPRateBox.value = Sheet.Cells(1, 2).value * 100#
    NaklRateBox.value = Sheet.Cells(2, 2).value * 100#
    NDSRateBox.value = Sheet.Cells(3, 2).value * 100#
    NalogRateBox.value = Sheet.Cells(4, 2).value * 100#
    ProfitRateBox.value = Sheet.Cells(5, 2).value * 100#
End Sub

Private Sub UserForm_Initialize()
    Set NumBoxCollection = New Collection
    
    Dim TBox As Control
    For Each TBox In BudgetForm.Controls
        ' If this is a textbox
        If TypeOf TBox Is MSForms.TextBox Then
            Dim NBox As NumericTextBox
            Set NBox = New NumericTextBox
            Set NBox.TextBox = TBox
            NumBoxCollection.Add NBox
        End If
    Next TBox
End Sub
