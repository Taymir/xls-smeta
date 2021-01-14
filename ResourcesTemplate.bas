Private nWB As Workbook
Private MiMWB As Worksheet
Private MRWB As Worksheet

Public Sub createBook()
    Workbooks.Add
    Set nWB = ActiveWorkbook
    setBook nWB
End Sub

Public Sub setBook(WB As Workbook)
    Set nWB = WB
    
    Set MiMWB = nWB.Sheets.Add(After:= _
        nWB.Sheets(nWB.Sheets.Count))
    MiMWB.name = "МиМ"
    
    Set MRWB = nWB.Sheets.Add(After:= _
        nWB.Sheets(nWB.Sheets.Count))
    MRWB.name = "МР"
End Sub

Public Sub fill_MiM(nameRange As Range, unitRange As Range, amountRange As Range, priceRange As Range)
    CopyRange nameRange, MiMWB.Range("B3")
    CopyRange unitRange, MiMWB.Range("C3")
    CopyRange amountRange, MiMWB.Range("D3")
    CopyRange priceRange, MiMWB.Range("E3")
End Sub

Public Sub render_MiM(Optional MiM)

    lastrow = get_last_row(MiMWB) + 1 ' (для прочего)
    
    With MiMWB
        ' нумерация
        For i = 3 To lastrow
            .Range(.Cells(i, 1), .Cells(i, 1)).Value = i - 2
        Next i
        .Range(.Cells(lastrow, 2), .Cells(lastrow, 2)).Value = "Прочее"
        
        .Names.Add name:="MiMOther", RefersTo:=.Range(.Cells(lastrow, 6), .Cells(lastrow, 6))
        .Names("MiMOther").Comment = "Машины и механизмы - Прочее"
        .Range("MiMOther").Value = 0
        
    
    
        ' Форматирование
        .Range("A1:F2").HorizontalAlignment = xlCenter
        .Range("A1:F2").VerticalAlignment = xlBottom
        
        .Range("A1:A2").Merge
        .Range("A1:A2").FormulaR1C1 = "N П/П"
        
        .Columns("B:B").ColumnWidth = 38
        .Range("B1:B2").Merge
        .Range("B1:B2").FormulaR1C1 = "Наименование"
        
    
        .Range("C1:C2").Merge
        .Range("C1:C2").FormulaR1C1 = "ед. изм."
        
        .Range("D1:F1").Merge
        .Range("D1:F1").FormulaR1C1 = "Сметное (планируемое)"
        
        .Range("D2").FormulaR1C1 = "Кол-во"
        .Range("E2").FormulaR1C1 = "Цена за ед."
        .Range("F2").FormulaR1C1 = "Итого"
        
        .Columns("A:A").ColumnWidth = 7.71
        .Columns("C:C").ColumnWidth = 8.71
        .Columns("C:C").HorizontalAlignment = xlCenter
        .Columns("C:C").VerticalAlignment = xlCenter
        .Columns("D:D").ColumnWidth = 9
        .Columns("E:E").ColumnWidth = 11
        .Columns("F:F").ColumnWidth = 11
        
        .Columns("C:F").NumberFormat = "#,##0.00"
    
    
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlEdgeLeft).Weight = xlThin
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlEdgeTop).Weight = xlThin
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlEdgeRight).Weight = xlThin
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlInsideVertical).Weight = xlThin
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(lastrow, 6)).Borders(xlInsideHorizontal).Weight = xlThin
        .Range(.Cells(3, 6), .Cells(lastrow, 6)).FormulaR1C1 = "=RC[-1]*RC[-2]"
        
        ' Итого
        .Names.Add name:="MiMTotal", RefersTo:=.Range(.Cells(lastrow + 2, 6), .Cells(lastrow + 2, 6))
        .Names("MiMTotal").Comment = "Машины и механизмы - Итого"
        
        .Range(.Cells(lastrow + 2, 5), .Cells(lastrow + 2, 5)).Value = "Итого"
        .Range(.Cells(lastrow + 3, 5), .Cells(lastrow + 3, 5)).Value = "НДС"
        .Range(.Cells(lastrow + 4, 5), .Cells(lastrow + 4, 5)).Value = "Всего с НДС"
        
        .Range(.Cells(lastrow + 2, 6), .Cells(lastrow + 2, 6)).Formula = "=SUM(F3:F" & lastrow & ")"
        .Range(.Cells(lastrow + 3, 6), .Cells(lastrow + 3, 6)).FormulaR1C1 = "=R[-1]C*0.2"
        .Range(.Cells(lastrow + 4, 6), .Cells(lastrow + 4, 6)).FormulaR1C1 = "=R[-2]C+R[-1]C"
        
        .Range(.Cells(lastrow + 2, 5), .Cells(lastrow + 4, 5)).Font.Bold = True
        .Range(.Cells(lastrow + 2, 5), .Cells(lastrow + 2, 6)).Font.Bold = True
        
        If Not IsMissing(MiM) Then
            .Range("MiMOther").Value = MiM - .Range("MiMTotal").Value
        End If
    End With
    
    render_MiM_Additional_form

End Sub

Private Sub render_MiM_Additional_form()
   ' форматирование
    With MiMWB
        row = get_last_row(MiMWB) + 5
    
    
    
        ' форматирование
        .Rows(row).RowHeight = 37.5
        .Range(.Cells(row, 1), .Cells(row + 18, 5)).Borders(xlEdgeLeft).LineStyle = xlDouble
        .Range(.Cells(row, 1), .Cells(row + 18, 5)).Borders(xlEdgeLeft).Weight = xlThick
        .Range(.Cells(row, 1), .Cells(row + 18, 5)).Borders(xlEdgeTop).LineStyle = xlDouble
        .Range(.Cells(row, 1), .Cells(row + 18, 5)).Borders(xlEdgeTop).Weight = xlThick
        .Range(.Cells(row, 1), .Cells(row + 18, 5)).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(.Cells(row, 1), .Cells(row + 18, 5)).Borders(xlEdgeBottom).Weight = xlThick
        .Range(.Cells(row, 1), .Cells(row + 18, 5)).Borders(xlEdgeRight).LineStyle = xlDouble
        .Range(.Cells(row, 1), .Cells(row + 18, 5)).Borders(xlEdgeRight).Weight = xlThick
        
        .Range(.Cells(row, 1), .Cells(row, 5)).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(.Cells(row, 1), .Cells(row, 5)).Borders(xlEdgeBottom).Weight = xlThick
        
        .Range(.Cells(row, 1), .Cells(row, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(.Cells(row, 1), .Cells(row, 5)).Borders(xlInsideVertical).Weight = xlThin
        
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 5)).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 5)).Borders(xlEdgeBottom).Weight = xlThick
        
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 4)).Borders(xlEdgeRight).LineStyle = xlDouble
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 4)).Borders(xlEdgeRight).Weight = xlThick
        
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlEdgeLeft).LineStyle = xlDouble
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlEdgeLeft).Weight = xlThick
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlEdgeTop).LineStyle = xlDouble
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlEdgeTop).Weight = xlThick
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlEdgeBottom).Weight = xlThick
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlEdgeRight).LineStyle = xlDouble
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlEdgeRight).Weight = xlThick
        
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlInsideVertical).Weight = xlThin
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range(.Cells(row + 2, 1), .Cells(row + 10, 4)).Borders(xlInsideHorizontal).Weight = xlThin
        
        .Range(.Cells(row + 12, 1), .Cells(row + 17, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(.Cells(row + 12, 1), .Cells(row + 17, 4)).Borders(xlInsideVertical).Weight = xlThin
        .Range(.Cells(row + 12, 1), .Cells(row + 17, 4)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range(.Cells(row + 12, 1), .Cells(row + 17, 4)).Borders(xlInsideHorizontal).Weight = xlThin
        
        .Range(.Cells(row + 11, 1), .Cells(row + 11, 4)).Borders(xlEdgeTop).LineStyle = xlDouble
        .Range(.Cells(row + 11, 1), .Cells(row + 11, 4)).Borders(xlEdgeTop).Weight = xlThick
        .Range(.Cells(row + 11, 1), .Cells(row + 11, 4)).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(.Cells(row + 11, 1), .Cells(row + 11, 4)).Borders(xlEdgeBottom).Weight = xlThick
        .Range(.Cells(row + 11, 1), .Cells(row + 11, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(.Cells(row + 11, 1), .Cells(row + 11, 4)).Borders(xlInsideVertical).Weight = xlThin
        
        .Range(.Cells(row + 2, 5), .Cells(row + 9, 5)).Borders(xlEdgeTop).LineStyle = xlDouble
        .Range(.Cells(row + 2, 5), .Cells(row + 9, 5)).Borders(xlEdgeTop).Weight = xlThick
        .Range(.Cells(row + 2, 5), .Cells(row + 9, 5)).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(.Cells(row + 2, 5), .Cells(row + 9, 5)).Borders(xlEdgeBottom).Weight = xlThick
        .Range(.Cells(row + 2, 5), .Cells(row + 9, 5)).Borders(xlInsideVertical).LineStyle = xlDouble
        .Range(.Cells(row + 2, 5), .Cells(row + 9, 5)).Borders(xlInsideVertical).Weight = xlThick
        
        .Range(.Cells(row + 18, 1), .Cells(row + 18, 5)).Borders(xlEdgeTop).LineStyle = xlDouble
        .Range(.Cells(row + 18, 1), .Cells(row + 18, 5)).Borders(xlEdgeTop).Weight = xlThick
        
        .Range(.Cells(row + 18, 1), .Cells(row + 18, 4)).Borders(xlEdgeTop).LineStyle = xlDouble
        .Range(.Cells(row + 18, 1), .Cells(row + 18, 4)).Borders(xlEdgeTop).Weight = xlThick
        .Range(.Cells(row + 18, 1), .Cells(row + 18, 4)).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(.Cells(row + 18, 1), .Cells(row + 18, 4)).Borders(xlEdgeBottom).Weight = xlThick
        .Range(.Cells(row + 18, 1), .Cells(row + 18, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(.Cells(row + 18, 1), .Cells(row + 18, 4)).Borders(xlInsideVertical).Weight = xlThin
        
        .Range(.Cells(row + 12, 1), .Cells(row + 12, 4)).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(.Cells(row + 12, 1), .Cells(row + 12, 4)).Borders(xlEdgeBottom).Weight = xlThick
        
        .Range(.Cells(row + 11, 5), .Cells(row + 16, 5)).Borders(xlEdgeTop).LineStyle = xlDouble
        .Range(.Cells(row + 11, 5), .Cells(row + 16, 5)).Borders(xlEdgeTop).Weight = xlThick
        .Range(.Cells(row + 11, 5), .Cells(row + 16, 5)).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(.Cells(row + 11, 5), .Cells(row + 16, 5)).Borders(xlEdgeBottom).Weight = xlThick
        .Range(.Cells(row + 11, 5), .Cells(row + 18, 5)).Borders(xlEdgeLeft).LineStyle = xlDouble
        .Range(.Cells(row + 11, 5), .Cells(row + 18, 5)).Borders(xlEdgeLeft).Weight = xlThick
        .Range(.Cells(row + 11, 5), .Cells(row + 11, 5)).Borders(xlEdgeBottom).LineStyle = xlDouble
        .Range(.Cells(row + 11, 5), .Cells(row + 11, 5)).Borders(xlEdgeBottom).Weight = xlThick
        
        
        .Columns("C:C").ColumnWidth = 13.14
        .Columns("D:D").ColumnWidth = 13
        .Columns("E:E").ColumnWidth = 17.57
        
        .Range(.Cells(row, 1), .Cells(row + 1, 5)).VerticalAlignment = xlCenter
        .Range(.Cells(row, 1), .Cells(row + 1, 5)).HorizontalAlignment = xlCenter
        .Range(.Cells(row + 12, 1), .Cells(row + 12, 4)).HorizontalAlignment = xlCenter
        .Range(.Cells(row + 1, 5), .Cells(row + 1, 5)).HorizontalAlignment = xlCenter
        
        .Range(.Cells(row, 1), .Cells(row + 1, 5)).Font.Bold = True
        .Range(.Cells(row + 12, 1), .Cells(row + 12, 1)).Font.Bold = True
        .Range(.Cells(row + 11, 1), .Cells(row + 11, 1)).Font.Bold = True
        .Range(.Cells(row + 18, 1), .Cells(row + 18, 1)).Font.Bold = True
        .Range(.Cells(row, 5), .Cells(row, 5)).Font.Color = -16776961
        
        .Range(.Cells(row, 1), .Cells(row, 1)).FormulaR1C1 = "N"
        .Range(.Cells(row, 2), .Cells(row, 2)).FormulaR1C1 = "Наименование"
        .Range(.Cells(row, 3), .Cells(row, 3)).FormulaR1C1 = "Стоимость с учетом НДС"
        .Range(.Cells(row, 4), .Cells(row, 4)).FormulaR1C1 = "Стоимость без ндс"
        .Range(.Cells(row, 5), .Cells(row, 5)).FormulaR1C1 = "Кол-во смен крана"
        
        .Range(.Cells(row, 3), .Cells(row, 3)).WrapText = True
        .Range(.Cells(row, 4), .Cells(row, 4)).WrapText = True
        
        
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 4)).Merge
        .Range(.Cells(row + 12, 1), .Cells(row + 12, 4)).Merge
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 4)).FormulaR1C1 = "Машины и механизмы"
        .Range(.Cells(row + 12, 1), .Cells(row + 12, 4)).FormulaR1C1 = "Зарплата машинистов"
        
        .Range(.Cells(row + 1, 5), .Cells(row + 1, 5)).Interior.Pattern = xlSolid
        .Range(.Cells(row + 1, 5), .Cells(row + 1, 5)).Interior.ThemeColor = xlThemeColorAccent2
        .Range(.Cells(row + 1, 5), .Cells(row + 1, 5)).Interior.TintAndShade = 0.799981688894314
        
        .Range(.Cells(row + 2, 5), .Cells(row + 9, 5)).Interior.ThemeColor = xlThemeColorDark1
        .Range(.Cells(row + 12, 5), .Cells(row + 16, 5)).Interior.ThemeColor = xlThemeColorDark1
        
        .Range(.Cells(row + 2, 1), .Cells(row + 2, 1)).FormulaR1C1 = "1"
        .Range(.Cells(row + 3, 1), .Cells(row + 3, 1)).FormulaR1C1 = "2"
        .Range(.Cells(row + 4, 1), .Cells(row + 4, 1)).FormulaR1C1 = "3"
        .Range(.Cells(row + 5, 1), .Cells(row + 5, 1)).FormulaR1C1 = "4"
        .Range(.Cells(row + 6, 1), .Cells(row + 6, 1)).FormulaR1C1 = "5"
        .Range(.Cells(row + 7, 1), .Cells(row + 7, 1)).FormulaR1C1 = "6"
        .Range(.Cells(row + 8, 1), .Cells(row + 8, 1)).FormulaR1C1 = "7"
        .Range(.Cells(row + 9, 1), .Cells(row + 9, 1)).FormulaR1C1 = "8"
        .Range(.Cells(row + 10, 1), .Cells(row + 10, 1)).FormulaR1C1 = "9"
        .Range(.Cells(row + 1, 5), .Cells(row + 1, 5)).FormulaR1C1 = "0"
        
        .Range(.Cells(row + 11, 2), .Cells(row + 11, 2)).FormulaR1C1 = "Итого"
        .Range(.Cells(row + 11, 2), .Cells(row + 11, 2)).Font.Color = -16776961
        .Range(.Cells(row + 11, 2), .Cells(row + 11, 2)).HorizontalAlignment = xlCenter
        
        .Range(.Cells(row + 18, 2), .Cells(row + 18, 2)).FormulaR1C1 = "Итого"
        .Range(.Cells(row + 18, 2), .Cells(row + 18, 2)).Font.Color = -16776961
        .Range(.Cells(row + 18, 2), .Cells(row + 18, 2)).HorizontalAlignment = xlCenter
        
        .Range(.Cells(row + 10, 5), .Cells(row + 10, 5)).FormulaR1C1 = "НДС 20%"
        .Range(.Cells(row + 10, 5), .Cells(row + 10, 5)).Font.Color = -16776961
        .Range(.Cells(row + 17, 5), .Cells(row + 17, 5)).FormulaR1C1 = "НДФЛ 13%"
        .Range(.Cells(row + 17, 5), .Cells(row + 17, 5)).Font.Color = -16776961
        
        .Range(.Cells(row + 13, 1), .Cells(row + 13, 1)).FormulaR1C1 = "1"
        .Range(.Cells(row + 14, 1), .Cells(row + 14, 1)).FormulaR1C1 = "2"
        .Range(.Cells(row + 15, 1), .Cells(row + 15, 1)).FormulaR1C1 = "3"
        .Range(.Cells(row + 16, 1), .Cells(row + 16, 1)).FormulaR1C1 = "4"
        .Range(.Cells(row + 17, 1), .Cells(row + 17, 1)).FormulaR1C1 = "5"
        
        .Range(.Cells(row + 10, 2), .Cells(row + 10, 2)).FormulaR1C1 = "Оплата по безналичному расчету"
        
        .Range(.Cells(row + 11, 3), .Cells(row + 11, 3)).FormulaR1C1 = "=SUM(R[-9]C:R[-1]C)"
        .Range(.Cells(row + 11, 4), .Cells(row + 11, 4)).FormulaR1C1 = "=SUM(R[-9]C:R[-1]C)"
        
        .Range(.Cells(row + 18, 3), .Cells(row + 18, 3)).FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
        .Range(.Cells(row + 18, 4), .Cells(row + 18, 4)).FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
        
        .Range(.Cells(row + 11, 5), .Cells(row + 11, 5)).FormulaR1C1 = "=RC[-2]-RC[-1]"
        .Range(.Cells(row + 18, 5), .Cells(row + 18, 5)).FormulaR1C1 = "=RC[-1]-(RC[-1]*0.87)"
    
    
    End With
End Sub

Private Function get_last_row(ws As Worksheet) As Integer
    'get_last_row = nWS.Cells(nWS.Cells.Rows.Count, "A").End(xlUp).row
    get_last_row = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).row ' учитывает пустые колонки
    If get_last_row < 2 Then
        get_last_row = 2
    End If

End Function

Public Sub fill_MR(nameRange As Range, unitRange As Range, amountRange As Range, priceRange As Range)
    ' заполнение данных
    CopyRange nameRange, MRWB.Range("C3")
    CopyRange unitRange, MRWB.Range("D3")
    CopyRange amountRange, MRWB.Range("E3")
    CopyRange priceRange, MRWB.Range("F3")
End Sub

Public Sub render_MR(Optional MR)
    
    lastrow = get_last_row(MRWB) + 1 ' для прочего
    
    ' нумерация
    With MRWB
        For i = 3 To lastrow
            .Range(.Cells(i, 1), .Cells(i, 1)).Value = i - 2
        Next i
        .Range(.Cells(lastrow, 3), .Cells(lastrow, 3)).Value = "Прочее"
        
        .Names.Add name:="MROther", RefersTo:=.Range(.Cells(lastrow, 7), .Cells(lastrow, 7))
        .Names("MROther").Comment = "Материальные ресурсы - Прочее"
        .Range("MROther").Value = 0


        ' Форматирование
        .Range("A1:Z2").Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range("A1:Z2").Borders(xlEdgeLeft).Weight = xlMedium
        .Range("A1:Z2").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range("A1:Z2").Borders(xlEdgeTop).Weight = xlMedium
        .Range("A1:Z2").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A1:Z2").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("A1:Z2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("A1:Z2").Borders(xlEdgeRight).Weight = xlMedium
        
        .Range("A1:Z2").HorizontalAlignment = xlCenter
        .Range("A1:Z2").VerticalAlignment = xlCenter
        
        .Range("A1:A2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("A1:A2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("A1:A2").Merge
        .Range("A1:A2").Value = "N п/п"
        
        .Range("B1:B2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("B1:B2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("B1:B2").Merge
        
        .Columns("C:C").ColumnWidth = 32.29
        
        .Range("C1:C2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("C1:C2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("C1:C2").Merge
        .Range("C1:C2").FormulaR1C1 = "Наименование МР"
        
        .Range("D1:D2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("D1:D2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("D1:D2").Merge
        .Range("D1:D2").FormulaR1C1 = "Ед.изм"
        
        .Range("E1:G1").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("E1:G1").Borders(xlEdgeRight).Weight = xlMedium
        .Range("E1:G1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("E1:G1").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("E1:G1").Merge
        .Range("E1:G1").FormulaR1C1 = "Сметное (планируемое)"
        
        .Range("E2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("E2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("E2").FormulaR1C1 = "Кол-во"
        
        .Range("F2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("F2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("F2").FormulaR1C1 = "Цена за ед."
        
        .Columns("G:G").ColumnWidth = 16.86
        
        .Range("G2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("G2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("G2").FormulaR1C1 = "Итого"
        
        .Columns("F:F").ColumnWidth = 11.43
        .Columns("E:E").ColumnWidth = 9.86
        .Columns("H:H").ColumnWidth = 37.57
        
        .Range("H1:H2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("H1:H2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("H1:H2").Merge
        .Range("H1:H2").FormulaR1C1 = "Наименование МР"
        
        .Range("I1:I2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("I1:I2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("I1:I2").Merge
        .Range("I1:I2").FormulaR1C1 = "Ед. изм"
        
        
        
        .Columns("L:L").ColumnWidth = 14.43
        .Columns("K:K").ColumnWidth = 10.71
        .Columns("J:J").ColumnWidth = 11.43
        .Columns("L:L").ColumnWidth = 13.43
        
        .Range("J1:L1").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("J1:L1").Borders(xlEdgeRight).Weight = xlMedium
        .Range("J1:L1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("J1:L1").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("J1:L1").Merge
        .Range("J1:L1").FormulaR1C1 = "Коммерческая смета"
        
        .Range("J2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("J2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("J2").FormulaR1C1 = "Кол-во"
        
        .Range("K2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("K2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("K2").FormulaR1C1 = "Цена за ед."
        
        .Range("L2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("L2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("L2").FormulaR1C1 = "Итого"
        
        
        
        .Columns("M:M").ColumnWidth = 11.29
        .Columns("N:N").ColumnWidth = 10.86
        .Columns("O:O").ColumnWidth = 13.14
        
        .Range("M1:O1").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("M1:O1").Borders(xlEdgeRight).Weight = xlMedium
        .Range("M1:O1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("M1:O1").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("M1:O1").Merge
        .Range("M1:O1").FormulaR1C1 = "Заявки"
        
        .Range("M2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("M2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("M2").FormulaR1C1 = "N"
        
        .Range("N2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("N2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("N2").FormulaR1C1 = "Ед.изм"
        
        .Range("O2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("O2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("O2").FormulaR1C1 = "Кол-во"
        
        
        
        .Columns("P:P").ColumnWidth = 11.14
        .Columns("Q:Q").ColumnWidth = 11
        .Columns("R:R").ColumnWidth = 11.43
        .Columns("S:S").ColumnWidth = 12.43
        .Columns("T:T").ColumnWidth = 11.57
        
        .Range("P1:T1").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("P1:T1").Borders(xlEdgeRight).Weight = xlMedium
        .Range("P1:T1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("P1:T1").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("P1:T1").Merge
        .Range("P1:T1").FormulaR1C1 = "Счета"
        
        .Range("Q2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("Q2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("Q2").FormulaR1C1 = "N"
        
        .Range("R2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("R2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("R2").FormulaR1C1 = "Ед.изм"
        
        .Range("O2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("O2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("O2").FormulaR1C1 = "Кол-во"
        
        .Range("S2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("S2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("S2").FormulaR1C1 = "Цена за ед."
        
        .Range("T2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("T2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("T2").FormulaR1C1 = "Стоимость"
        
        
        
        .Columns("U:U").ColumnWidth = 12.71
        .Columns("V:V").ColumnWidth = 12.86
        
        .Range("U1:W1").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("U1:W1").Borders(xlEdgeRight).Weight = xlMedium
        .Range("U1:W1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("U1:W1").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("U1:W1").Merge
        .Range("U1:W1").FormulaR1C1 = "Фактическое"
        
        .Range("U2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("U2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("U2").FormulaR1C1 = "N"
        
        .Range("V2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("V2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("V2").FormulaR1C1 = "Цена за ед"
        
        .Range("W2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("W2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("W2").FormulaR1C1 = "Кол-во"
        
        
        
        .Columns("X:X").ColumnWidth = 11.14
        .Columns("Y:Y").ColumnWidth = 10.57
        
        .Range("X1:Z1").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("X1:Z1").Borders(xlEdgeRight).Weight = xlMedium
        .Range("X1:Z1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("X1:Z1").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("X1:Z1").Merge
        .Range("X1:Z1").FormulaR1C1 = "Списанное"
        
        .Range("X2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("X2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("X2").FormulaR1C1 = "Кол-во"
        
        .Range("Y2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("Y2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("Y2").FormulaR1C1 = "Цена за ед"
        
        .Range("Z2").Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range("Z2").Borders(xlEdgeRight).Weight = xlMedium
        .Range("Z2").FormulaR1C1 = "Итого"
    
        .Range(.Cells(3, 1), .Cells(lastrow, 26)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(lastrow, 26)).Borders(xlEdgeLeft).Weight = xlThin
        .Range(.Cells(3, 1), .Cells(lastrow, 26)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(lastrow, 26)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(3, 1), .Cells(lastrow, 26)).Borders(xlEdgeRight).LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(lastrow, 26)).Borders(xlEdgeRight).Weight = xlThin
        .Range(.Cells(3, 1), .Cells(lastrow, 26)).Borders(xlInsideVertical).LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(lastrow, 26)).Borders(xlInsideVertical).Weight = xlThin
        .Range(.Cells(3, 1), .Cells(lastrow, 26)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(lastrow, 26)).Borders(xlInsideHorizontal).Weight = xlThin
            
        .Range(.Cells(1, 13), .Cells(lastrow, 15)).Interior.Pattern = xlSolid
        .Range(.Cells(1, 13), .Cells(lastrow, 15)).Interior.ThemeColor = xlThemeColorAccent4
        .Range(.Cells(1, 13), .Cells(lastrow, 15)).Interior.TintAndShade = 0.799981688894314
    
        .Range(.Cells(1, 16), .Cells(lastrow, 20)).Interior.Pattern = xlSolid
        .Range(.Cells(1, 16), .Cells(lastrow, 20)).Interior.ThemeColor = xlThemeColorAccent2
        .Range(.Cells(1, 16), .Cells(lastrow, 20)).Interior.TintAndShade = 0.799981688894314
            
        .Range(.Cells(3, 7), .Cells(lastrow, 7)).FormulaR1C1 = "=RC[-1]*RC[-2]"
        .Range(.Cells(3, 5), .Cells(lastrow + 5, 7)).NumberFormat = "#,##0.00"
        
        ' Итого
        .Range(.Cells(lastrow + 2, 6), .Cells(lastrow + 2, 6)).Value = "Итого"
        .Range(.Cells(lastrow + 3, 6), .Cells(lastrow + 3, 6)).Value = "НДС"
        .Range(.Cells(lastrow + 4, 6), .Cells(lastrow + 4, 6)).Value = "Всего с НДС"
        
        .Range(.Cells(lastrow + 2, 7), .Cells(lastrow + 2, 7)).FormulaR1C1 = _
        "=SUM(INDIRECT(ADDRESS(MATCH(""Итого"",C,0)+1,COLUMN())):INDIRECT(ADDRESS(ROW()-1,COLUMN())))"
        .Range(.Cells(lastrow + 3, 7), .Cells(lastrow + 3, 7)).FormulaR1C1 = "=R[-1]C*0.2"
        .Range(.Cells(lastrow + 4, 7), .Cells(lastrow + 4, 7)).FormulaR1C1 = "=R[-2]C+R[-1]C"
        
        .Range(.Cells(lastrow + 2, 6), .Cells(lastrow + 4, 6)).Font.Bold = True
        .Range(.Cells(lastrow + 2, 6), .Cells(lastrow + 2, 7)).Font.Bold = True
        
        .Names.Add name:="MRTotal", RefersTo:=.Range(.Cells(lastrow + 2, 7), .Cells(lastrow + 2, 7))
        .Names("MRTotal").Comment = "Материальные ресурсы - Итого"
        
        
        
        
        .Range(.Cells(lastrow + 2, 11), .Cells(lastrow + 2, 11)).Value = "Итого"
        .Range(.Cells(lastrow + 3, 11), .Cells(lastrow + 3, 11)).Value = "НДС"
        .Range(.Cells(lastrow + 4, 11), .Cells(lastrow + 4, 11)).Value = "Всего с НДС"
        
        .Range(.Cells(lastrow + 2, 12), .Cells(lastrow + 2, 12)).FormulaR1C1 = _
        "=SUM(INDIRECT(ADDRESS(MATCH(""Итого"",C,0)+1,COLUMN())):INDIRECT(ADDRESS(ROW()-1,COLUMN())))"
        .Range(.Cells(lastrow + 3, 12), .Cells(lastrow + 3, 12)).FormulaR1C1 = "=R[-1]C*0.2"
        .Range(.Cells(lastrow + 4, 12), .Cells(lastrow + 4, 12)).FormulaR1C1 = "=R[-2]C+R[-1]C"
        
        .Range(.Cells(lastrow + 2, 11), .Cells(lastrow + 4, 11)).Font.Bold = True
        .Range(.Cells(lastrow + 2, 11), .Cells(lastrow + 2, 12)).Font.Bold = True
        
        
        
        .Range(.Cells(lastrow + 2, 19), .Cells(lastrow + 2, 19)).Value = "Итого"
        .Range(.Cells(lastrow + 3, 19), .Cells(lastrow + 3, 19)).Value = "НДС"
        .Range(.Cells(lastrow + 4, 19), .Cells(lastrow + 4, 19)).Value = "Всего с НДС"
        
        .Range(.Cells(lastrow + 2, 20), .Cells(lastrow + 2, 20)).FormulaR1C1 = _
        "=SUM(INDIRECT(ADDRESS(MATCH(""Стоимость"",C,0)+1,COLUMN())):INDIRECT(ADDRESS(ROW()-1,COLUMN())))"
        .Range(.Cells(lastrow + 3, 20), .Cells(lastrow + 3, 20)).FormulaR1C1 = "=R[-1]C*0.2"
        .Range(.Cells(lastrow + 4, 20), .Cells(lastrow + 4, 20)).FormulaR1C1 = "=R[-2]C+R[-1]C"
        
        .Range(.Cells(lastrow + 2, 19), .Cells(lastrow + 4, 19)).Font.Bold = True
        .Range(.Cells(lastrow + 2, 19), .Cells(lastrow + 2, 20)).Font.Bold = True
        
        If Not IsMissing(MR) Then
            .Range("MROther").Value = MR - .Range("MRTotal").Value
        End If
        
    End With


    

    
End Sub



Private Sub CopyRange(from, too)
    from.Copy
    too.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
End Sub


