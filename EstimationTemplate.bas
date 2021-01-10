' Class EstimationTemplate
' Содержит шаблоны для рендеринга результата-таблицы сметы

Private nWB As Workbook
Private nWS As Worksheet

Public Sub createBook()
    Workbooks.Add
    Set nWB = ActiveWorkbook
    Set nWS = nWB.Worksheets(1)
    nWS.name = "КомСм"
    
    'Set createBook = nWS
End Sub

Public Sub renderHeader(object, smetaName)
    ' Ширина первых строк
    nWS.Rows(1).RowHeight = 27.25
    nWS.Rows(2).RowHeight = 42.75
    
    ' Ширина столбцов
    nWS.Columns(1).ColumnWidth = 3.45
    nWS.Columns(2).ColumnWidth = 16.2
    nWS.Columns(3).ColumnWidth = 40.3
    nWS.Columns(4).ColumnWidth = 12.15
    nWS.Columns(5).ColumnWidth = 9.75
    nWS.Columns(6).ColumnWidth = 11.88
    nWS.Columns(7).ColumnWidth = 17.6
    nWS.Columns(8).ColumnWidth = 14.6
    nWS.Columns(9).ColumnWidth = 16
    nWS.Columns(10).ColumnWidth = 16.75
    nWS.Columns(11).ColumnWidth = 14.7
    nWS.Columns(12).ColumnWidth = 16.7
    nWS.Columns(13).ColumnWidth = 18.3
    
    ' Объединение, расположение текста
    nWS.Range("A1:J2").Merge
    nWS.Range("A1:J2").HorizontalAlignment = xlCenter
    nWS.Range("A1:J2").VerticalAlignment = xlCenter
    nWS.Range("A1:J2").Borders.LineStyle = True
    
    ' Название объекта
    nWS.Cells(1, 1).Value = object
    
    ' Статья Бюджета
    nWS.Cells(1, 1).Font.name = "Arial"
    nWS.Cells(1, 1).Font.Size = 14
    nWS.Cells(1, 1).WrapText = True
    nWS.Range("K1:M1").Merge
    nWS.Range("K1:M2").HorizontalAlignment = xlCenter
    nWS.Range("K1:M2").VerticalAlignment = xlCenter
    nWS.Range("K1:M2").Borders.LineStyle = True
    nWS.Range("K1:M2").Interior.Color = 3243501 '5296274 - условное форматирование
    nWS.Cells(1, 11).Value = "Статья Бюджета"
    nWS.Cells(1, 11).Font.name = "Arial"
    nWS.Cells(1, 11).Font.Size = 12
    nWS.Cells(1, 11).Font.Bold = True
    nWS.Cells(1, 11).Font.Italic = True
    nWS.Cells(2, 11).Font.name = "Arial"
    nWS.Cells(2, 11).Font.Size = 10
    nWS.Cells(2, 11).Font.Bold = True
    'nWS.Cells(2, 11).Formula = ""
    'nWS.Cells(2, 11).FormulaR1C1 = _
    '    "=IF(LEFT(R[2]C[-10],5)=""02-01"",""Корпус А"",IF(LEFT(R[2]C[-10],5)=""02-02"",""Корпус Б"",""НЕТ ДАННЫХ""))"
    nWS.Cells(2, 12).Font.name = "Arial"
    nWS.Cells(2, 12).Font.Size = 10
    nWS.Cells(2, 12).Font.Bold = True
    'nWS.Cells(2, 12).Formula = ""
    'nWS.Cells(2, 12).FormulaR1C1 = _
    '    "=IF(RC[-1]=""Корпус А"",INDEX(БЮДЖЕТ!R[2]C[-11]:R[12]C[-10],MATCH(RC[1],БЮДЖЕТ!R[2]C[-10]:R[12]C[-10],0),1),IF(RC[-1]=""Корпус Б"",INDEX(БЮДЖЕТ!R[14]C[-11]:R[36]C[-10],MATCH(RC[1],БЮДЖЕТ!R[14]C[-10]:R[36]C[-10],0),1),""НЕТ ДАННЫХ""))"
    
    ' Смета
    nWS.Cells(2, 13).Font.name = "Arial"
    nWS.Cells(2, 13).Font.Size = 10
    nWS.Cells(2, 13).Font.Bold = True
    nWS.Cells(2, 13).WrapText = True
    'nWS.Cells(2, 13).Formula = ""
    nWS.Range("A3:M3").Merge
    nWS.Range("A3:M3").HorizontalAlignment = xlCenter
    nWS.Range("A3:M3").VerticalAlignment = xlCenter
    nWS.Cells(3, 1).Value = "Согласование коммерческих расценок на выполнение работ для физических лиц"
    nWS.Cells(3, 1).Font.name = "Arial"
    nWS.Cells(3, 1).Font.Size = 16
    nWS.Cells(3, 1).Font.Bold = True
    nWS.Range("A4:M4").Merge
    nWS.Range("A4:M4").HorizontalAlignment = xlCenter
    nWS.Range("A4:M4").VerticalAlignment = xlCenter
    nWS.Cells(4, 1).Value = smetaName
    
    ' Заголовок таблицы
    nWS.Cells(4, 1).Font.name = "Arial"
    nWS.Cells(4, 1).Font.Size = 14
    nWS.Cells(4, 1).Font.Italic = True
    nWS.Cells(4, 1).Font.Bold = True
    nWS.Cells(4, 1).Font.Color = 8421504
    nWS.Cells(4, 1).WrapText = True
    nWS.Rows(4).RowHeight = 18.75
    nWS.Rows(5).RowHeight = 13.5
    nWS.Rows(6).RowHeight = 12.75
    nWS.Rows(7).RowHeight = 39
    nWS.Rows(8).RowHeight = 13.5
    nWS.Range("A6:A7").Merge
    nWS.Range("A6:M8").HorizontalAlignment = xlCenter
    nWS.Range("A6:M8").VerticalAlignment = xlCenter
    nWS.Range("A6:M8").WrapText = True
    nWS.Range("A6:M8").Borders.LineStyle = True
    nWS.Range("A6:M8").Borders.Weight = xlMedium
    nWS.Range("B6:B7").Merge
    nWS.Range("C6:C7").Merge
    nWS.Range("D6:D7").Merge
    nWS.Range("E6:E7").Merge
    nWS.Range("F6:H6").Merge
    nWS.Range("D6:D7").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("D6:D7").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("I6:K6").Merge
    nWS.Range("L6:M7").Merge
    nWS.Range("G7").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("G7").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("F7:H7").Borders(xlEdgeTop).Weight = xlThin
    nWS.Range("J7").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("J7").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("I7:K7").Borders(xlEdgeTop).Weight = xlThin
    'nWS.Cells(6, 1).Font.Name = "Arial"
    'nWS.Cells(6, 1).Font.Size = 10
    'nWS.Cells(6, 1).Font.Bold = True
    nWS.Range("A6:M8").Font.name = "Arial"
    nWS.Range("A6:M8").Font.Size = 10
    nWS.Range("A6:M8").Font.Bold = True
    nWS.Range("D8").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("D8").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("G8").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("G8").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("J8").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("J8").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("M8").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Cells(6, 1).Value = "№ п/п"
    nWS.Cells(6, 2).Value = "Шифр расценки"
    nWS.Cells(6, 3).Value = "Наименование работ"
    nWS.Cells(6, 4).Value = "Ед. измерения"
    nWS.Cells(6, 5).Value = "Кол-во"
    nWS.Cells(6, 6).Value = "Локальная смета"
    nWS.Cells(7, 6).Value = "Стоимость за ед."
    nWS.Cells(7, 7).Value = "ИТОГО"
    nWS.Cells(7, 8).Value = "% в общей сумме затрат в смете"
    nWS.Cells(6, 9).Value = "Коммерческая смета"
    nWS.Cells(7, 9).Value = "Стоимость за ед."
    nWS.Cells(7, 10).Value = "ИТОГО"
    nWS.Cells(7, 11).Value = "% в общей сумме затрат в смете"
    nWS.Cells(6, 12).Value = "Финансовый результат"
    nWS.Cells(8, 1).Value = 1
    nWS.Cells(8, 2).Value = 2
    nWS.Cells(8, 3).Value = 3
    nWS.Cells(8, 4).Value = 4
    nWS.Cells(8, 5).Value = 5
    nWS.Cells(8, 6).Value = 6
    nWS.Cells(8, 7).Value = 7
    nWS.Cells(8, 8).Value = 8
    nWS.Cells(8, 9).Value = 9
    nWS.Cells(8, 10).Value = 10
    nWS.Cells(8, 11).Value = 11
    nWS.Cells(8, 12).Value = 12
    nWS.Cells(8, 13).Value = 13
    ' заголовок отрисован
End Sub

Private Function get_last_row() As Integer
    'get_last_row = nWS.Cells(nWS.Cells.Rows.Count, "A").End(xlUp).row
    get_last_row = nWS.UsedRange.Rows(nWS.UsedRange.Rows.Count).row ' учитывает пустые колонки

End Function

Public Sub render_section(name)
    With nWS
        row = get_last_row + 1
    
        .Range(.Cells(row, 1), .Cells(row, 13)).Merge
        
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.name = "Arial"
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Size = 14
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Bold = True
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Italic = True
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders.LineStyle = True
        .Range(.Cells(row, 1), .Cells(row, 13)).HorizontalAlignment = xlCenter
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders.Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row, 13)).Value = "Раздел: " & name
    End With
End Sub

Public Sub render_subsection(name)
    render_section (name) ' тот же шаблон, что и для основного раздела
End Sub

Public Sub render_item(num, code, name, unit, amount, total, total_fot)
    With nWS
        row = get_last_row + 1
        
        ' Шрифт
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Font.name = "Arial"
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Font.Size = 11
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Font.Bold = True
        .Range(.Cells(row, 6), .Cells(row, 6)).Font.Bold = False
        .Range(.Cells(row, 8), .Cells(row, 9)).Font.Bold = False
        .Range(.Cells(row, 13), .Cells(row, 13)).Font.Bold = False
        .Range(.Cells(row + 1, 2), .Cells(row + 1, 13)).Font.Bold = False
        ' В т.ч. Фот
        .Range(.Cells(row + 1, 3), .Cells(row + 1, 3)).Font.Italic = True
        .Range(.Cells(row + 1, 3), .Cells(row + 1, 3)).Font.Size = 10
        
        
        ' Границы
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Borders.LineStyle = True
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Borders(xlEdgeLeft).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders(xlEdgeBottom).Weight = xlHairline
        
        ' Ориентация
        .Range(.Cells(row, 1), .Cells(row, 13)).VerticalAlignment = xlCenter
        .Range(.Cells(row, 1), .Cells(row, 2)).HorizontalAlignment = xlCenter
        .Range(.Cells(row, 3), .Cells(row, 3)).WrapText = True
        .Range(.Cells(row, 4), .Cells(row, 5)).HorizontalAlignment = xlCenter
        .Range(.Cells(row + 1, 3), .Cells(row + 1, 3)).HorizontalAlignment = xlRight
        
        
        ' Формат данных
        .Range(.Cells(row, 5), .Cells(row + 1, 5)).NumberFormat = "#,##0.00"
        .Range(.Cells(row, 6), .Cells(row + 1, 7)).NumberFormat = "#,##0"
        .Range(.Cells(row, 10), .Cells(row + 1, 12)).NumberFormat = "#,##0.00"
        .Range(.Cells(row, 9), .Cells(row + 1, 9)).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

        .Range(.Cells(row, 8), .Cells(row + 1, 8)).NumberFormat = "0.0%"
        .Range(.Cells(row, 11), .Cells(row + 1, 11)).NumberFormat = "0.0%"
        .Range(.Cells(row, 13), .Cells(row + 1, 13)).NumberFormat = "0.0%"
        
        
        ' Заполнение данных
        .Range(.Cells(row, 1), .Cells(row, 1)).Value = num       ' П/П
        .Range(.Cells(row, 2), .Cells(row, 2)).Value = code      ' Шифр расценки
        .Range(.Cells(row, 3), .Cells(row, 3)).Value = name      ' Наименование работ
        .Range(.Cells(row + 1, 3), .Cells(row + 1, 3)).Value = "в т.ч. ФОТ" 'todo форматирование
        .Range(.Cells(row, 4), .Cells(row, 4)).Value = unit      ' Ед. измерения
        .Range(.Cells(row, 5), .Cells(row, 5)).Value = amount    ' кол-во
        .Range(.Cells(row, 7), .Cells(row, 7)).Value = total     ' ИТОГО
        .Range(.Cells(row + 1, 7), .Cells(row + 1, 7)).Value = total_fot ' Итого - ФОТ
        
        ' Вставка формул
        ' 6: F10=G10/E10, F11=G11/E10
        .Range(.Cells(row, 6), .Cells(row, 6)).FormulaR1C1 = "=RC[1]/RC[-1]"
        .Range(.Cells(row + 1, 6), .Cells(row + 1, 6)).FormulaR1C1 = "=RC[1]/R[-1]C[-1]"
        
        ' 8: H10=G10/GrandTotal, H11=G11/GrandTotal
        .Range(.Cells(row, 8), .Cells(row + 1, 8)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
        ' 9: I11=I10
        .Range(.Cells(row + 1, 9), .Cells(row + 1, 9)).FormulaR1C1 = "=R[-1]C"
        
        ' 10: J10=E10*ОКРУГЛ(I10;2), J11=E10*ОКРУГЛ(I11;2)
        .Range(.Cells(row, 10), .Cells(row, 10)).FormulaR1C1 = "=RC[-5]*ROUND(RC[-1],2)"
        .Range(.Cells(row + 1, 10), .Cells(row + 1, 10)).FormulaR1C1 = "=R[-1]C[-5]*ROUND(RC[-1],2)"
        
        ' 11: K10=J10/GrandTotal, K11=J11/GrandTotal
        .Range(.Cells(row, 11), .Cells(row + 1, 11)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
        ' 12 L10=G10-J10, L11=G11-J11
        .Range(.Cells(row, 12), .Cells(row + 1, 12)).FormulaR1C1 = "=RC[-5]-RC[-2]"
        
        ' 13: M10=L10/GrandTotal, M11=L11/GrandTotal
        .Range(.Cells(row, 13), .Cells(row + 1, 13)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
        ' условное форматирование
        Set f = .Range(.Cells(row, 12), .Cells(row + 1, 13)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0")
        f.Font.ColorIndex = 30
        
    
        
        ' Ширина строки
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).EntireRow.AutoFit
        .Rows(row).RowHeight = .Rows(row).RowHeight + 15 ' Увеличение автовысоты на 15
        
    End With
End Sub

Public Sub render_footer(MR, MiM, ZPmas, NR, SP)
    
    render_footer1
    render_footer2 MR, MiM, ZPmas
    render_footer3 NR, SP
    render_footer4
    render_footer5
    
    ' tmp Add Grand Total named Range
    row = get_last_row + 2
    
    
    nWS.Cells(row, 7).Value = 42285286.22
    
    nWS.Names.Add name:="GrandTotal", RefersTo:=nWS.Range(nWS.Cells(row, 7), nWS.Cells(row, 7))
    nWS.Names("GrandTotal").Comment = "ВСЕГО ЗАТРАТ В СМЕТЕ"
End Sub

' Итого по смете
Private Sub render_footer1()
    With nWS
        row = get_last_row + 1
        'пустая строка
        .Range(.Cells(row, 1), .Cells(row, 13)).Merge
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders.LineStyle = True
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders.Weight = xlMedium
        row = row + 1
        
        ' Итого по смете
        .Range(.Cells(row, 1), .Cells(row, 6)).Merge
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders.Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders.LineStyle = True
        .Rows(row).RowHeight = 13.5
        
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Borders.LineStyle = True
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Font.name = "Arial"
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Size = 11
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Font.Size = 10
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Bold = True
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Borders(xlEdgeLeft).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 1, 13)).Borders(xlEdgeBottom).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row, 6)).Merge
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 6)).Merge
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).HorizontalAlignment = xlRight
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Font.Italic = True
        .Cells(row, 1).Value = "Итого по смете:"
        .Cells(row + 1, 1).Value = "в т.ч. ФОТ"
        
        ' Пустая строка
        .Range(.Cells(row + 2, 1), .Cells(row + 2, 13)).Merge
        .Range(.Cells(row + 2, 1), .Cells(row + 2, 13)).Borders.LineStyle = True
        .Range(.Cells(row + 2, 1), .Cells(row + 2, 13)).Borders.Weight = xlMedium
        .Rows(row + 1).RowHeight = 13.5
    End With
End Sub

' Свод прямых затрат в смете
Private Sub render_footer2(MR, MiM, ZPmas)
    With nWS
        row = get_last_row + 1
            
        ' Рамки
        .Range(.Cells(row, 1), .Cells(row + 8, 13)).Borders.LineStyle = True
        .Range(.Cells(row, 1), .Cells(row + 8, 13)).Interior.Color = 14809087
        .Range(.Cells(row, 3), .Cells(row, 13)).Merge
        .Range(.Cells(row, 3), .Cells(row + 8, 3)).IndentLevel = 1
        .Range(.Cells(row, 1), .Cells(row + 8, 13)).Borders(xlEdgeLeft).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 8, 13)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 8, 5)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 8, 8)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 8, 11)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 8, 13)).Borders(xlEdgeBottom).Weight = xlMedium
        .Range(.Cells(row, 2), .Cells(row, 2)).Borders(xlEdgeRight).LineStyle = False
        
        ' Шрифты
        .Range(.Cells(row, 1), .Cells(row + 8, 13)).Font.name = "Arial"
        .Range(.Cells(row, 1), .Cells(row + 8, 13)).Font.Size = 11
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Bold = True
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Size = 14
        .Range(.Cells(row + 8, 1), .Cells(row + 8, 13)).Font.Bold = True
        .Range(.Cells(row + 8, 1), .Cells(row + 8, 13)).Font.Size = 12
        .Range(.Cells(row, 1), .Cells(row, 1)).HorizontalAlignment = xlCenter
        
        ' Текст
        .Cells(row, 1).Value = "II"
        .Cells(row, 3).Value = "Свод прямых затрат в смете"
        .Cells(row + 1, 3).Value = "ФОТ по позициям"
        .Cells(row + 2, 3).Value = "Вывоз мусора" ' Нужно ли??
        .Cells(row + 3, 3).Value = "Материальные ресурсы"
        .Cells(row + 4, 3).Value = "Субподряд"
        .Range(.Cells(row + 5, 3), .Cells(row + 5, 3)).Rows.WrapText = True
        .Cells(row + 5, 3).Value = "Машины, механизмы, з/п механизаторов, в т.ч.:"
        .Range(.Cells(row + 6, 3), .Cells(row + 7, 3)).HorizontalAlignment = xlRight
        .Cells(row + 6, 3).Value = "аренда машин и механизмов"
        .Cells(row + 6, 3).Font.Italic = True
        .Cells(row + 7, 3).Value = "з/п машинистов"
        .Cells(row + 7, 3).Font.Italic = True
        .Cells(row + 8, 3).Value = "Итого прямых затрат в смете"
        
        ' Заполнение данных
        .Cells(row + 3, 7).Value = MR ' Материальные ресурсы
        .Cells(row + 5, 7).Value = MiM ' Машины и механизмы
        .Cells(row + 7, 7).Value = ZPmas ' ЗП Машинистов
        
        
        ' Пустая строка
        .Range(.Cells(row + 9, 1), .Cells(row + 9, 13)).Merge
        .Range(.Cells(row + 9, 1), .Cells(row + 9, 13)).Borders.LineStyle = True
        .Range(.Cells(row + 9, 1), .Cells(row + 9, 13)).Borders.Weight = xlMedium
        .Rows(row + 9).RowHeight = 13.5

            
            
    End With
End Sub

Private Sub render_footer3(NR, SP)
    With nWS
        row = get_last_row + 1
        
        'Свод дополнительных затрат в смете
        .Range(.Cells(row, 1), .Cells(row + 10, 13)).Borders.LineStyle = True
        .Range(.Cells(row, 3), .Cells(row, 13)).Merge
        .Range(.Cells(row, 1), .Cells(row + 10, 13)).Borders(xlEdgeLeft).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 10, 13)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 10, 5)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 10, 8)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 10, 11)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 10, 13)).Borders(xlEdgeBottom).Weight = xlMedium
        .Range(.Cells(row, 3), .Cells(row + 10, 3)).IndentLevel = 1
        
        .Range(.Cells(row, 1), .Cells(row + 10, 13)).Font.name = "Arial"
        .Range(.Cells(row, 1), .Cells(row + 10, 13)).Font.Size = 11
        .Range(.Cells(row + 2, 1), .Cells(row + 4, 13)).Font.Size = 10
        .Range(.Cells(row + 10, 1), .Cells(row + 10, 13)).Font.Size = 12
        
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Bold = True
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Size = 14
        .Range(.Cells(row, 1), .Cells(row, 1)).HorizontalAlignment = xlCenter
        .Cells(row, 1).Value = "III"
        .Cells(row, 3).Value = "Свод дополнительных затрат в смете"
        .Cells(row + 1, 3).Value = "Накладные расходы, в том числе:"
        .Range(.Cells(row + 2, 3), .Cells(row + 4, 3)).Rows.WrapText = True
        .Range(.Cells(row + 2, 3), .Cells(row + 4, 3)).Font.Italic = True
        .Range(.Cells(row + 2, 3), .Cells(row + 4, 3)).Font.Color = 7434613
        .Range(.Cells(row + 2, 3), .Cells(row + 4, 3)).IndentLevel = 2
        .Cells(row + 2, 3).Value = "Административно-хозяйственные расходы (5% от сметы)"
        .Cells(row + 3, 3).Value = "Расходы на обслуживание работников строительства"
        .Cells(row + 4, 3).Value = "Расходы на организацию работ на строительных площадках (2,48% от сметы)"
        .Cells(row + 5, 3).Value = "Сметная прибыль"
        .Cells(row + 6, 3).Value = "Зимнее удорожание 1,41%"
        .Cells(row + 7, 3).Value = "НДС 20%, в т.ч."
        .Cells(row + 8, 3).Font.Color = 7434613
        .Cells(row + 8, 3).Value = "НДС уплаченный поставщикам"
        .Cells(row + 8, 3).Font.Italic = True
        .Cells(row + 9, 3).Font.Color = 255
        .Cells(row + 9, 3).Value = "НДС к уплате в бюджет"
        .Cells(row + 9, 3).Font.Italic = True
        .Cells(row + 9, 3).HorizontalAlignment = xlRight
        .Cells(row + 10, 3).Font.Size = 14
        .Cells(row + 10, 3).Font.Bold = True
        .Cells(row + 10, 3).Value = "Итого дополнительных затрат в смете"
        Range(.Cells(row + 10, 3), .Cells(row + 10, 3)).Rows.WrapText = True
        'Свод дополнительных затрат в смете
        
        ' Заполнение данных
        .Cells(row + 1, 7).Value = NR ' Накладные расходы
        .Cells(row + 5, 7).Value = SP ' Сметная прибыль
        
        
    End With
End Sub
    
    
Private Sub render_footer4()
    With nWS
        row = get_last_row + 1
        
        'всего затрат в смете
        .Range(.Cells(row, 1), .Cells(row, 13)).Merge
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders.LineStyle = True
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders.Weight = xlMedium
        .Rows(row).RowHeight = 13.5
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Borders.LineStyle = True
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Font.name = "Arial"
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Font.Bold = True
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Font.Size = 12
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Font.Italic = True
        .Cells(row + 1, 3).HorizontalAlignment = xlRight
        .Cells(row + 1, 3).Value = "ВСЕГО ЗАТРАТ В СМЕТЕ"
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Borders(xlEdgeLeft).Weight = xlMedium
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 5)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 8)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 11)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Borders(xlEdgeBottom).Weight = xlMedium
        'всего затрат в смете
        
        
    End With
End Sub

Private Sub render_footer5()
    With nWS
        row = get_last_row + 2
        
        'подвал
        .Range(.Cells(row, 1), .Cells(row, 3)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(row + 2, 1), .Cells(row + 2, 3)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(row + 4, 1), .Cells(row + 4, 3)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(row + 6, 1), .Cells(row + 6, 3)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(row + 8, 1), .Cells(row + 8, 3)).Borders(xlEdgeBottom).Weight = xlThin
        
        .Range(.Cells(row, 1), .Cells(row + 9, 5)).Font.name = "Arial"
        .Range(.Cells(row, 1), .Cells(row + 9, 1)).Font.Bold = True
        .Range(.Cells(row, 1), .Cells(row + 9, 1)).Font.Size = 10
        .Range(.Cells(row, 4), .Cells(row + 9, 4)).Font.Size = 14
        
        .Cells(row, 1).Value = "Зам. Руководителя ДС"
        .Cells(row + 2, 1).Value = "Главный инженер"
        .Cells(row + 4, 1).Value = "Нач. отдела стр. аудита"
        .Cells(row + 6, 1).Value = "Главный экономист"
        .Cells(row + 8, 1).Value = "Руководитель управления финансов"
        .Cells(row + 9, 1).Value = "и экономики ДС"
        .Cells(row, 4).Value = "Павлов М.М."
        .Cells(row + 2, 4).Value = "Гущин И.А."
        .Cells(row + 4, 4).Value = "Игнатова Т.К."
        .Cells(row + 6, 4).Value = "Кодрау И.И."
        .Cells(row + 8, 4).Value = "Мамонтова А.В."
        
        .Cells(row + 1, 3).Value = "(подпись)"
        .Cells(row + 1, 3).Font.Italic = True
        .Cells(row + 1, 3).Font.Size = 7
        .Cells(row + 1, 3).HorizontalAlignment = xlCenter
        .Cells(row + 1, 3).VerticalAlignment = xlTop
        .Cells(row + 3, 3).Value = "(подпись)"
        .Cells(row + 3, 3).Font.Italic = True
        .Cells(row + 3, 3).Font.Size = 7
        .Cells(row + 3, 3).HorizontalAlignment = xlCenter
        .Cells(row + 3, 3).VerticalAlignment = xlTop
        .Cells(row + 5, 3).Value = "(подпись)"
        .Cells(row + 5, 3).Font.Italic = True
        .Cells(row + 5, 3).Font.Size = 7
        .Cells(row + 5, 3).HorizontalAlignment = xlCenter
        .Cells(row + 5, 3).VerticalAlignment = xlTop
        .Cells(row + 7, 3).Value = "(подпись)"
        .Cells(row + 7, 3).Font.Italic = True
        .Cells(row + 7, 3).Font.Size = 7
        .Cells(row + 7, 3).HorizontalAlignment = xlCenter
        .Cells(row + 7, 3).VerticalAlignment = xlTop
        .Cells(row + 9, 3).Value = "(подпись)"
        .Cells(row + 9, 3).Font.Italic = True
        .Cells(row + 9, 3).Font.Size = 7
        .Cells(row + 9, 3).HorizontalAlignment = xlCenter
        .Cells(row + 9, 3).VerticalAlignment = xlTop
        
        .Range(.Cells(row, 6), .Cells(row + 8, 13)).Borders.LineStyle = True
        .Range(.Cells(row, 6), .Cells(row + 8, 13)).Borders.Weight = xlMedium
        .Range(.Cells(row + 3, 6), .Cells(row + 3, 13)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(row + 4, 6), .Cells(row + 4, 13)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(row + 6, 6), .Cells(row + 6, 13)).Borders(xlEdgeBottom).Weight = xlThin
        .Range(.Cells(row + 2, 6), .Cells(row + 8, 10)).Borders(xlEdgeRight).Weight = xlThin
        .Range(.Cells(row, 6), .Cells(row + 1, 9)).Merge
        .Range(.Cells(row, 10), .Cells(row + 1, 11)).Merge
        .Range(.Cells(row, 12), .Cells(row + 1, 12)).Merge
        .Range(.Cells(row, 13), .Cells(row + 1, 13)).Merge
        .Range(.Cells(row, 6), .Cells(row + 9, 13)).Font.name = "Arial"
        .Range(.Cells(row, 6), .Cells(row + 3, 13)).Font.Bold = True
        .Range(.Cells(row + 6, 6), .Cells(row + 6, 13)).Font.Bold = True
        .Range(.Cells(row + 8, 6), .Cells(row + 8, 13)).Font.Bold = True
        .Range(.Cells(row, 6), .Cells(row + 9, 13)).Font.Size = 11
        .Range(.Cells(row, 6), .Cells(row + 9, 13)).VerticalAlignment = xlCenter
        .Range(.Cells(row, 6), .Cells(row + 1, 13)).HorizontalAlignment = xlCenter
        .Range(.Cells(row, 6), .Cells(row + 1, 13)).VerticalAlignment = xlCenter
        .Cells(row, 6).Value = "Показатели"
        .Cells(row, 10).Value = "Коммерческая смета"
        .Cells(row, 12).Value = "Утвержденный бюджет, %"
        .Range(.Cells(row, 12), .Cells(row + 1, 12)).Rows.WrapText = True
        .Cells(row, 13).Value = "Отклонение, %"
        .Range(.Cells(row + 2, 6), .Cells(row + 3, 9)).Merge
        .Range(.Cells(row + 2, 10), .Cells(row + 3, 10)).Merge
        .Range(.Cells(row + 2, 11), .Cells(row + 3, 11)).Merge
        .Range(.Cells(row + 2, 12), .Cells(row + 3, 12)).Merge
        .Range(.Cells(row + 2, 13), .Cells(row + 3, 13)).Merge
        .Range(.Cells(row + 4, 6), .Cells(row + 4, 9)).Merge
        .Range(.Cells(row + 5, 6), .Cells(row + 5, 9)).Merge
        .Range(.Cells(row + 6, 6), .Cells(row + 6, 9)).Merge
        .Range(.Cells(row + 7, 6), .Cells(row + 7, 9)).Merge
        .Range(.Cells(row + 8, 6), .Cells(row + 8, 9)).Merge
        .Range(.Cells(row, 6), .Cells(row + 8, 9)).IndentLevel = 2
        .Range(.Cells(row + 2, 6), .Cells(row + 3, 9)).IndentLevel = 1
        .Range(.Cells(row + 6, 6), .Cells(row + 6, 9)).IndentLevel = 1
        .Range(.Cells(row + 8, 6), .Cells(row + 8, 9)).IndentLevel = 1
        .Range(.Cells(row + 2, 6), .Cells(row + 3, 13)).Interior.Color = 16247773
        .Range(.Cells(row + 8, 6), .Cells(row + 8, 13)).Interior.Color = 16247773
        .Range(.Cells(row + 6, 6), .Cells(row + 6, 13)).Interior.Color = 16247773
        .Cells(row + 2, 6).Value = "Финансовый результат" & Chr(10) & "(прибыль до уплаты налогов в бюджет и АУП)"
        .Cells(row + 2, 6).Rows.WrapText = True
        .Cells(row + 4, 6).Value = "АУП"
        .Cells(row + 4, 6).Font.Italic = True
        .Cells(row + 5, 6).Value = "НДС к уплате в бюджет"
        .Cells(row + 5, 6).Font.Italic = True
        .Cells(row + 6, 6).Value = "Валовая прибыль"
        .Cells(row + 7, 6).Value = "Налог на прибыль"
        .Cells(row + 7, 6).Font.Italic = True
        .Cells(row + 8, 6).Value = "ЧИСТАЯ ПРИБЫЛЬ ОТ ПРОИЗВОДСТВА РАБОТ"
        'подвал
    End With
End Sub

    
