Sub openSmetaAndTransform_click()
    Filename = Application.GetOpenFilename(FileFilter:="Excel files,*.xl*;*.xm*")
    If Filename <> False Then
        Workbooks.Open Filename:=Filename
        transformSmeta
    End If
End Sub

Sub openSmeta_click()
    Filename = Application.GetOpenFilename(FileFilter:="Excel files,*.xl*;*.xm*")
    If Filename <> False Then
        Workbooks.Open Filename:=Filename
        Dim btn As Button
        Set btn = ActiveSheet.Buttons.Add(0, 12.75, 160, 25.5)
        With btn
            .OnAction = "transformSmeta"
            .Caption = "Преобразовать в смету"
            .name = "transBtn"
        End With
    End If
End Sub

