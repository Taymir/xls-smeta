Private WithEvents NumTB As MSForms.TextBox

Public Property Set TextBox(ByVal TB As MSForms.TextBox)
    Set NumTB = TB
End Property

Private Sub NumTB_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If InStr(NumTB.name, "Date") > 0 Then
        ' Hotkey for date fields - ctrl + ;
        If (KeyCode = 186 Or KeyCode = Asc("J")) And Shift = 2 Then
            NumTB.value = Format(Date, "dd/mm/yyyy")
        End If
    End If
End Sub

Private Sub NumTB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57
            ' Allow digits - do nothing
        Case 44, 46
            ' Replace comma with dot
            If InStr(NumTB.name, "Date") = 0 And InStr(NumTB.value, ".") > 0 Then
                KeyAscii = 0
            Else
            KeyAscii = 46
            End If
        Case 47
            ' Allow slash only for dates
            If InStr(NumTB.name, "Date") = 0 Then
                KeyAscii = 0
            End If
        Case Else
            ' Block others
            KeyAscii = 0
    End Select
End Sub
