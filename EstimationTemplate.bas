' Class EstimationTemplate
' Содержит шаблоны для рендеринга результата-таблицы сметы

Public nWB As Workbook
Public nWS As Worksheet
Private Genpordryad_row As Integer

Public Sub createBook()
    Workbooks.Add
    Set nWB = ActiveWorkbook
    Set nWS = nWB.Worksheets(1)
    nWS.name = "КомСм"
    
    'Set createBook = nWB
End Sub

Public Sub render_header(object, smetaName)
    ' Ширина первых строк
    nWS.Rows(1).RowHeight = 18
    nWS.Rows(2).RowHeight = 27
    nWS.Rows(3).RowHeight = 27
    
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
    Dim Budget As BudgetController
    Set Budget = New BudgetController
    Dim ShowAgreement As Boolean
    ShowAgreement = Budget.hasOffsets(CStr(object))
    HeaderRange = IIf(ShowAgreement, "A1:G3", "A1:H3")
    nWS.Range(HeaderRange).Merge
    nWS.Range(HeaderRange).HorizontalAlignment = xlCenter
    nWS.Range(HeaderRange).VerticalAlignment = xlCenter
    nWS.Range(HeaderRange).Borders.LineStyle = True
    
    ' Название объекта
    nWS.Cells(1, 1).value = object
    nWS.Cells(1, 1).Font.name = "Arial"
    nWS.Cells(1, 1).Font.Size = 14
    nWS.Cells(1, 1).WrapText = True
    
    ' Договор
    If ShowAgreement Then
        nWS.Cells(1, 8).value = "Договор"
        nWS.Range("H2:H3").Merge
        nWS.Range("H2:H3").value = "РСУ"
        With nWS.Range("H2:H3").Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="РСУ,ФОНД"
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
        End With
        
        nWS.Range("H1:H3").Interior.Color = RGB(219, 219, 240)
    End If
    
    ' Версия сметы
    nWS.Range("H1:M3").Font.name = "Arial"
    nWS.Range("H1:M3").Font.Size = 10
    nWS.Range("H1:M3").Font.Bold = True
    nWS.Range("H1:M3").HorizontalAlignment = xlCenter
    nWS.Range("H1:M3").VerticalAlignment = xlCenter
    nWS.Range("H1:M3").Borders.LineStyle = True
    'nWS.Range("A1:M3").Borders.Weight = xlMedium
    
    
    nWS.Range("I1:M3").Interior.Color = RGB(219, 219, 240)
    nWS.Range("I1:J1").Merge
    nWS.Range("I1:J1").value = "Версия сметы:"
    nWS.Range("I2:J2").Merge
    nWS.Range("I2:J2").value = "Первичная"
    With nWS.Range("I2:J2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Первичная,Корректировка"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    nWS.Range("I3:J3").Merge
    
    ' Статья бюджета
    nWS.Range("K1:M1").Merge
    nWS.Range("K1:M1").value = "Статья бюджета"
    nWS.Range("K2:K2").value = "№"
    nWS.Range("K3:K3").value = Budget.getBudgetItemBySmetaName(CStr(object), CStr(smetaName))
    nWS.Range("L2:M2").Merge
    nWS.Range("L2:M2").value = "Наименование"
    nWS.Range("L3:M3").Merge
    
    
    
    
    ' Смета
    nWS.Range("A4:M4").Merge
    nWS.Range("A4:M4").HorizontalAlignment = xlCenter
    nWS.Range("A4:M4").VerticalAlignment = xlCenter
    nWS.Cells(4, 1).value = "Согласование коммерческих расценок на выполнение работ для физических лиц"
    nWS.Cells(4, 1).Font.name = "Arial"
    nWS.Cells(4, 1).Font.Size = 16
    nWS.Cells(4, 1).Font.Bold = True
    nWS.Range("A5:M5").Merge
    nWS.Range("A5:M5").HorizontalAlignment = xlCenter
    nWS.Range("A5:M5").VerticalAlignment = xlCenter
    nWS.Cells(5, 1).value = smetaName
    
    ' Заголовок таблицы
    nWS.Cells(5, 1).Font.name = "Arial"
    nWS.Cells(5, 1).Font.Size = 14
    nWS.Cells(5, 1).Font.Italic = True
    nWS.Cells(5, 1).Font.Bold = True
    nWS.Cells(5, 1).Font.Color = 8421504
    nWS.Cells(5, 1).WrapText = True
    nWS.Rows(5).RowHeight = 18.75
    nWS.Rows(6).RowHeight = 13.5
    nWS.Rows(7).RowHeight = 12.75
    nWS.Rows(8).RowHeight = 39
    nWS.Rows(9).RowHeight = 13.5
    nWS.Range("A7:A8").Merge
    nWS.Range("A7:M9").HorizontalAlignment = xlCenter
    nWS.Range("A7:M9").VerticalAlignment = xlCenter
    nWS.Range("A7:M9").WrapText = True
    nWS.Range("A7:M9").Borders.LineStyle = True
    nWS.Range("A7:M9").Borders.Weight = xlMedium
    nWS.Range("B7:B8").Merge
    nWS.Range("C7:C8").Merge
    nWS.Range("D7:D8").Merge
    nWS.Range("E7:E8").Merge
    nWS.Range("F7:H7").Merge
    nWS.Range("D7:D8").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("D7:D8").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("I7:K7").Merge
    nWS.Range("L7:M8").Merge
    nWS.Range("G8").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("G8").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("F8:H8").Borders(xlEdgeTop).Weight = xlThin
    nWS.Range("J8").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("J8").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("I8:K8").Borders(xlEdgeTop).Weight = xlThin
    'nWS.Cells(7, 1).Font.Name = "Arial"
    'nWS.Cells(7, 1).Font.Size = 10
    'nWS.Cells(7, 1).Font.Bold = True
    nWS.Range("A7:M9").Font.name = "Arial"
    nWS.Range("A7:M9").Font.Size = 10
    nWS.Range("A7:M9").Font.Bold = True
    nWS.Range("D9").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("D9").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("G9").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("G9").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("J9").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Range("J9").Borders(xlEdgeRight).Weight = xlThin
    nWS.Range("M9").Borders(xlEdgeLeft).Weight = xlThin
    nWS.Cells(7, 1).value = "№ п/п"
    nWS.Cells(7, 2).value = "Шифр расценки"
    nWS.Cells(7, 3).value = "Наименование работ"
    nWS.Cells(7, 4).value = "Ед. измерения"
    nWS.Cells(7, 5).value = "Кол-во"
    nWS.Cells(7, 6).value = "Локальная смета"
    nWS.Cells(8, 6).value = "Стоимость за ед."
    nWS.Cells(8, 7).value = "ИТОГО"
    nWS.Cells(8, 8).value = "% в общей сумме затрат в смете"
    nWS.Cells(7, 9).value = "Коммерческая смета"
    nWS.Cells(8, 9).value = "Стоимость за ед."
    nWS.Cells(8, 10).value = "ИТОГО"
    nWS.Cells(8, 11).value = "% в общей сумме затрат в смете"
    nWS.Cells(7, 12).value = "Финансовый результат"
    nWS.Cells(9, 1).value = 1
    nWS.Cells(9, 2).value = 2
    nWS.Cells(9, 3).value = 3
    nWS.Cells(9, 4).value = 4
    nWS.Cells(9, 5).value = 5
    nWS.Cells(9, 6).value = 6
    nWS.Cells(9, 7).value = 7
    nWS.Cells(9, 8).value = 8
    nWS.Cells(9, 9).value = 9
    nWS.Cells(9, 10).value = 10
    nWS.Cells(9, 11).value = 11
    nWS.Cells(9, 12).value = 12
    nWS.Cells(9, 13).value = 13
    ' заголовок отрисован
End Sub

Public Sub render_final_addons()
    nWS.Activate


    With nWS
        .Range("L3:M3").Formula = _
            "=INDEX(Бюджет!$B$6:$C$2000,MATCH(КомСм!$K$3,Бюджет!$B$6:$B$2000,0),2)"
        ' TODO Всегда ли есть бюджет? Определение первой и последней строки бюджета?
        
        If RangeExists("Генподряд") Then
            'Debug.Print ("range exists!!!")
            .Rows(Genpordryad_row).Insert
            .Range(.Cells(Genpordryad_row, 3), .Cells(Genpordryad_row, 3)).value = "Генподряд"
            .Range(.Cells(Genpordryad_row, 10), .Cells(Genpordryad_row, 10)).FormulaR1C1 = "=GrandTotal*Генподряд"
            .Range(.Cells(Genpordryad_row, 11), .Cells(Genpordryad_row, 11)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        End If
    End With
End Sub

Private Function get_last_row() As Integer
    'get_last_row = nWS.Cells(nWS.Cells.Rows.Count, "A").End(xlUp).row
    get_last_row = nWS.UsedRange.Rows(nWS.UsedRange.Rows.Count).row ' учитывает пустые колонки

End Function

Public Sub render_section(name, Optional text As String = "Раздел: ")
    With nWS
        row = get_last_row + 1
        
        ' HARDFIX TMP Временное решение проблемы с отсутствием разделов
        If name = "LocalSmeta" Then
            Exit Sub
        End If
    
        .Range(.Cells(row, 1), .Cells(row, 13)).Merge
        
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.name = "Arial"
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Size = 14
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Bold = True
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Italic = True
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders.LineStyle = True
        .Range(.Cells(row, 1), .Cells(row, 13)).HorizontalAlignment = xlCenter
        .Range(.Cells(row, 1), .Cells(row, 13)).Borders.Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row, 13)).value = text & name
    End With
End Sub

Public Sub render_subsection(name)
    render_section name, text:="Подраздел: "   ' тот же шаблон, что и для основного раздела
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
        
        .Range(.Cells(row, 1), .Cells(row + 1, 1)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 2), .Cells(row + 1, 2)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 5), .Cells(row + 1, 5)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 8), .Cells(row + 1, 8)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 11), .Cells(row + 1, 11)).Borders(xlEdgeRight).Weight = xlMedium
        
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
        .Range(.Cells(row, 1), .Cells(row, 1)).value = num       ' П/П
        .Range(.Cells(row, 2), .Cells(row, 2)).value = code      ' Шифр расценки
        .Range(.Cells(row, 3), .Cells(row, 3)).value = name      ' Наименование работ
        .Range(.Cells(row + 1, 3), .Cells(row + 1, 3)).value = "в т.ч. ФОТ" 'todo форматирование
        .Range(.Cells(row, 4), .Cells(row, 4)).value = unit      ' Ед. измерения
        .Range(.Cells(row, 5), .Cells(row, 5)).value = amount    ' кол-во
        .Range(.Cells(row, 7), .Cells(row, 7)).value = total     ' ИТОГО
        .Range(.Cells(row + 1, 7), .Cells(row + 1, 7)).value = total_fot ' Итого - ФОТ
        
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

Public Sub render_footer(MR, MiM, ZPmas, NR, SP, EH, EM)
    
    render_footer1
    render_footer2 MR, MiM, ZPmas
    render_footer3 NR, SP, EH, EM
    render_footer4
    render_footer5
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
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 1)).Font.Italic = True
        .Cells(row, 1).value = "Итого по смете:"
        .Cells(row + 1, 1).value = "в т.ч. ФОТ"
        
        ' Форматирование
        .Range(.Cells(row, 7), .Cells(row + 1, 7)).NumberFormat = "#,##0"
        .Range(.Cells(row, 8), .Cells(row + 1, 8)).NumberFormat = "0.0%"
        .Range(.Cells(row, 9), .Cells(row + 1, 9)).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .Range(.Cells(row, 10), .Cells(row + 1, 10)).NumberFormat = "#,##0.00"
        .Range(.Cells(row, 11), .Cells(row + 1, 11)).NumberFormat = "0.0%"
        .Range(.Cells(row, 12), .Cells(row + 1, 12)).NumberFormat = "#,##0.00"
        .Range(.Cells(row, 13), .Cells(row + 1, 13)).NumberFormat = "0.0%"
        
        ' Формулы
        .Names.Add name:="SmetTotal", RefersTo:=nWS.Range(.Cells(row, 7), .Cells(row, 7))
        .Names("SmetTotal").Comment = "Итого по смете"
        
        .Names.Add name:="FOT", RefersTo:=nWS.Range(.Cells(row + 1, 7), .Cells(row + 1, 7))
        .Names("FOT").Comment = "Фонд оплаты труда"
    
        .Range(.Cells(row, 7), .Cells(row, 7)).FormulaR1C1 = _
        "=SUMIF(INDIRECT(ADDRESS(MATCH(""Наименование работ"",C[-4],0)+3,3)):INDIRECT(ADDRESS(ROW()-1,3)),""<>в т.ч. ФОТ"",INDIRECT(ADDRESS(MATCH(""ИТОГО"",C,0)+2,COLUMN())):INDIRECT(ADDRESS(ROW()-1,COLUMN())))"
        .Range(.Cells(row + 1, 7), .Cells(row + 1, 7)).FormulaR1C1 = _
        "=SUMIF(INDIRECT(ADDRESS(MATCH(""Наименование работ"",C[-4],0)+3,3)):INDIRECT(ADDRESS(ROW()-2,3)),""=в т.ч. ФОТ"",INDIRECT(ADDRESS(MATCH(""ИТОГО"",C,0)+2,COLUMN())):INDIRECT(ADDRESS(ROW()-2,COLUMN())))"
        
        .Range(.Cells(row, 8), .Cells(row + 1, 8)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
        .Range(.Cells(row, 10), .Cells(row, 10)).FormulaR1C1 = _
        "=SUMIF(INDIRECT(ADDRESS(MATCH(""Наименование работ"",C[-7],0)+3,3)):INDIRECT(ADDRESS(ROW()-1,3)),""<>в т.ч. ФОТ"",INDIRECT(ADDRESS(MATCH(""ИТОГО"",C,0)+2,COLUMN())):INDIRECT(ADDRESS(ROW()-1,COLUMN())))"
        .Range(.Cells(row + 1, 10), .Cells(row + 1, 10)).FormulaR1C1 = _
        "=SUMIF(INDIRECT(ADDRESS(MATCH(""Наименование работ"",C[-7],0)+3,3)):INDIRECT(ADDRESS(ROW()-2,3)),""=в т.ч. ФОТ"",INDIRECT(ADDRESS(MATCH(""ИТОГО"",C,0)+2,COLUMN())):INDIRECT(ADDRESS(ROW()-2,COLUMN())))"

        .Range(.Cells(row, 11), .Cells(row + 1, 11)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
        .Range(.Cells(row, 12), .Cells(row + 1, 12)).FormulaR1C1 = "=RC[-5]-RC[-2]"
        
        .Range(.Cells(row, 13), .Cells(row + 1, 13)).FormulaR1C1 = "=RC[-1]/GrandTotal"

        
        
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
        .Range(.Cells(row, 1), .Cells(row + 6, 13)).Borders.LineStyle = True
        .Range(.Cells(row, 1), .Cells(row + 6, 13)).Interior.Color = 14809087
        .Range(.Cells(row, 3), .Cells(row, 13)).Merge
        .Range(.Cells(row, 3), .Cells(row + 6, 3)).IndentLevel = 1
        .Range(.Cells(row, 1), .Cells(row + 6, 13)).Borders(xlEdgeLeft).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 6, 13)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 6, 5)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 6, 8)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 6, 11)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row, 1), .Cells(row + 6, 13)).Borders(xlEdgeBottom).Weight = xlMedium
        .Range(.Cells(row, 2), .Cells(row, 2)).Borders(xlEdgeRight).LineStyle = False
        
        ' Шрифты
        .Range(.Cells(row, 1), .Cells(row + 6, 13)).Font.name = "Arial"
        .Range(.Cells(row, 1), .Cells(row + 6, 13)).Font.Size = 11
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Bold = True
        .Range(.Cells(row, 1), .Cells(row, 13)).Font.Size = 14
        .Range(.Cells(row + 6, 1), .Cells(row + 6, 13)).Font.Bold = True
        .Range(.Cells(row + 6, 1), .Cells(row + 6, 13)).Font.Size = 12
        .Range(.Cells(row, 1), .Cells(row, 1)).HorizontalAlignment = xlCenter
        
        ' Текст
        .Cells(row, 1).value = "II"
        .Cells(row, 3).value = "Свод прямых затрат в смете"
        .Cells(row + 1, 3).value = "ФОТ по позициям"
        .Cells(row + 2, 3).value = "Материальные ресурсы"
        .Range(.Cells(row + 3, 3), .Cells(row + 3, 3)).Rows.WrapText = True
        .Cells(row + 3, 3).value = "Машины, механизмы, з/п механизаторов, в т.ч.:"
        .Range(.Cells(row + 4, 3), .Cells(row + 5, 3)).HorizontalAlignment = xlRight
        .Cells(row + 4, 3).value = "аренда машин и механизмов"
        .Cells(row + 4, 3).Font.Italic = True
        .Cells(row + 5, 3).value = "з/п машинистов"
        .Cells(row + 5, 3).Font.Italic = True
        .Cells(row + 6, 3).value = "Итого прямых затрат в смете"
        
        
        
        ' Форматирование
        '.Range(.Cells(row + 3, 7), .Cells(row + 8, 7)).Font.Bold = True
        '.Range(.Cells(row + 8, 7), .Cells(row + 8, 7)).Font.Size = 12
        '.Range(.Cells(row + 8, 10), .Cells(row + 8, 10)).Font.Bold = True
        '.Range(.Cells(row + 8, 12), .Cells(row + 8, 12)).Font.Bold = True
        .Range(.Cells(row + 6, 12), .Cells(row + 6, 12)).Font.Size = 14
        .Range(.Cells(row + 6, 12), .Cells(row + 6, 12)).Font.ColorIndex = 41
        
        .Range(.Cells(row + 1, 7), .Cells(row + 6, 13)).VerticalAlignment = xlCenter
        .Range(.Cells(row + 1, 7), .Cells(row + 6, 7)).NumberFormat = "#,##0"
        .Range(.Cells(row + 1, 8), .Cells(row + 6, 8)).NumberFormat = "0.0%"
        .Range(.Cells(row + 1, 9), .Cells(row + 6, 9)).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .Range(.Cells(row + 1, 10), .Cells(row + 6, 10)).NumberFormat = "#,##0"
        .Range(.Cells(row + 1, 11), .Cells(row + 6, 11)).NumberFormat = "0.0%"
        .Range(.Cells(row + 1, 12), .Cells(row + 6, 12)).NumberFormat = "#,##0"
        .Range(.Cells(row + 1, 13), .Cells(row + 6, 13)).NumberFormat = "0.0%"
        
        ' Заполнение данных
        .Cells(row + 2, 7).value = MR ' Материальные ресурсы 'TMP
        .Cells(row + 3, 7).value = MiM ' Машины и механизмы  'TMP
        .Cells(row + 5, 7).value = ZPmas ' ЗП Машинистов     'TMP
        
        ' Формулы
        .Range(.Cells(row + 1, 7), .Cells(row + 1, 7)).FormulaR1C1 = "=FOT"
        .Range(.Cells(row + 4, 7), .Cells(row + 4, 7)).FormulaR1C1 = "=R[-1]C-R[1]C"
        .Range(.Cells(row + 6, 7), .Cells(row + 6, 7)).FormulaR1C1 = "=R[-5]C+R[-4]C+R[-3]C"
        
        .Range(.Cells(row + 1, 8), .Cells(row + 6, 8)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
        .Range(.Cells(row + 1, 10), .Cells(row + 1, 10)).FormulaR1C1 = "=R[-3]C"
        .Range(.Cells(row + 2, 10), .Cells(row + 2, 10)).value = 0 'TODO
        .Range(.Cells(row + 3, 10), .Cells(row + 3, 10)).FormulaR1C1 = "=R[1]C+R[2]C"
        .Range(.Cells(row + 4, 10), .Cells(row + 4, 10)).value = 0 'TODO
        .Range(.Cells(row + 5, 10), .Cells(row + 5, 10)).value = 0 'TODO
        .Range(.Cells(row + 6, 10), .Cells(row + 6, 10)).FormulaR1C1 = "=SUM(R[-5]C:R[-3]C)"
        
        .Range(.Cells(row + 1, 11), .Cells(row + 6, 11)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
        .Range(.Cells(row + 1, 12), .Cells(row + 6, 12)).FormulaR1C1 = "=RC[-5]-RC[-2]"
        
        .Range(.Cells(row + 1, 13), .Cells(row + 6, 13)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
        
        ' Пустая строка
        .Range(.Cells(row + 7, 1), .Cells(row + 7, 13)).Merge
        .Range(.Cells(row + 7, 1), .Cells(row + 7, 13)).Borders.LineStyle = True
        .Range(.Cells(row + 7, 1), .Cells(row + 7, 13)).Borders.Weight = xlMedium
        .Rows(row + 7).RowHeight = 13.5

            
            
    End With
End Sub



Private Sub render_footer3(NR, SP, EH, EM)
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
        
        .Cells(row, 1).value = "III"
        .Cells(row, 3).value = "Свод дополнительных затрат в смете"
        .Cells(row + 1, 3).value = "Накладные расходы, в том числе:"
        .Range(.Cells(row + 2, 3), .Cells(row + 4, 3)).Rows.WrapText = True
        .Range(.Cells(row + 2, 3), .Cells(row + 4, 3)).Font.Italic = True
        .Range(.Cells(row + 2, 3), .Cells(row + 4, 3)).Font.Color = 7434613
        .Range(.Cells(row + 2, 3), .Cells(row + 4, 3)).IndentLevel = 2
        '.Cells(row + 2, 3).Value = "Административно-хозяйственные расходы (5% от сметы)" ' OLD
        .Rows(row + 2).RowHeight = 28
        .Cells(row + 2, 3).Formula = "=""Административно-хозяйственные расходы ("" & Round(АУП * 100,1) & ""% от сметы)"""
        .Cells(row + 3, 3).value = "Расходы на обслуживание работников строительства"
        '.Cells(row + 4, 3).Value = "Расходы на организацию работ на строительных площадках (2,48% от сметы)" ' OLD
        .Rows(row + 4).RowHeight = 42
        .Cells(row + 4, 3).Formula = "=""Расходы на организацию работ на строительных площадках ("" & Round(НР * 100,1) & ""% от сметы)"""
        .Cells(row + 5, 3).value = "Сметная прибыль"
        .Cells(row + 6, 3).value = "Зимнее удорожание 1,41%"
        .Cells(row + 7, 3).value = "НДС 20%, в т.ч."
        .Cells(row + 8, 3).Font.Color = 7434613
        .Cells(row + 8, 3).value = "НДС уплаченный поставщикам"
        .Cells(row + 8, 3).Font.Italic = True
        .Cells(row + 9, 3).Font.Color = 255
        .Cells(row + 9, 3).value = "НДС к уплате в бюджет"
        .Cells(row + 9, 3).Font.Italic = True
        .Cells(row + 9, 3).HorizontalAlignment = xlRight
        .Cells(row + 10, 3).Font.Size = 14
        .Range(.Cells(row + 10, 3), .Cells(row + 10, 13)).Font.Bold = True
        .Cells(row + 10, 3).value = "Итого дополнительных затрат в смете"
        .Range(.Cells(row + 10, 3), .Cells(row + 10, 3)).Rows.WrapText = True
        'Свод дополнительных затрат в смете
        
        ' Форматирование
        .Range(.Cells(row + 1, 7), .Cells(row + 10, 13)).VerticalAlignment = xlCenter
        .Range(.Cells(row + 1, 7), .Cells(row + 10, 7)).NumberFormat = "#,##0"
        .Range(.Cells(row + 1, 8), .Cells(row + 10, 8)).NumberFormat = "0.0%"
        .Range(.Cells(row + 1, 9), .Cells(row + 10, 9)).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .Range(.Cells(row + 1, 10), .Cells(row + 10, 10)).NumberFormat = "#,##0"
        .Range(.Cells(row + 1, 11), .Cells(row + 10, 11)).NumberFormat = "0.0%"
        .Range(.Cells(row + 1, 12), .Cells(row + 10, 12)).NumberFormat = "#,##0"
        .Range(.Cells(row + 1, 13), .Cells(row + 10, 13)).NumberFormat = "0.0%"
        .Range(.Cells(row + 8, 10), .Cells(row + 9, 10)).NumberFormat = "#,##0.00"
        
        .Range(.Cells(row + 10, 12), .Cells(row + 10, 13)).Font.Size = 14
        .Range(.Cells(row + 10, 12), .Cells(row + 10, 13)).Font.ColorIndex = 41 'blue
        
        .Range(.Cells(row + 2, 10), .Cells(row + 4, 11)).Font.ColorIndex = 15 'grey
        .Range(.Cells(row + 2, 10), .Cells(row + 2, 10)).Font.ColorIndex = 3 'red
        .Range(.Cells(row + 2, 10), .Cells(row + 4, 11)).Font.Italic = True
        
        .Range(.Cells(row + 8, 10), .Cells(row + 8, 10)).Font.ColorIndex = 15 'grey
        .Range(.Cells(row + 9, 10), .Cells(row + 9, 10)).Font.ColorIndex = 3 'red
        .Range(.Cells(row + 8, 10), .Cells(row + 9, 10)).Font.Italic = True
        .Range(.Cells(row, 1), .Cells(row, 1)).HorizontalAlignment = xlCenter
        
        ' Заполнение данных
        .Cells(row + 1, 7).value = NR ' Накладные расходы
        .Cells(row + 5, 7).value = SP ' Сметная прибыль
        
        ' Формулы
        .Range(.Cells(row + 6, 7), .Cells(row + 6, 7)).FormulaR1C1 = "=(SmetTotal-" & Dbl2Str(EH) & "-" & Dbl2Str(EM) & ")*0.0141"
        .Range(.Cells(row + 7, 7), .Cells(row + 7, 7)).FormulaR1C1 = "=(SmetTotal+R[-1]C)*0.2"
        .Range(.Cells(row + 10, 7), .Cells(row + 10, 7)).FormulaR1C1 = "=SUM(R[-5]C:R[-3]C)+R[-9]C"
        
        .Range(.Cells(row + 1, 8), .Cells(row + 1, 8)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        .Range(.Cells(row + 5, 8), .Cells(row + 7, 8)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        .Range(.Cells(row + 10, 8), .Cells(row + 10, 8)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
        .Range(.Cells(row + 1, 10), .Cells(row + 1, 10)).FormulaR1C1 = "=SUM(R[1]C:R[3]C)"
        '.Range(.Cells(row + 2, 10), .Cells(row + 2, 10)).FormulaR1C1 = "=R[10]C[-3]/100*5" 'OLD
        .Range(.Cells(row + 2, 10), .Cells(row + 2, 10)).FormulaR1C1 = "=GrandTotal*АУП"
        
        .Range(.Cells(row + 3, 10), .Cells(row + 3, 10)).FormulaR1C1 = "=(R[-10]C+R[-6]C)/100*30.9"
        '.Range(.Cells(row + 4, 10), .Cells(row + 4, 10)).FormulaR1C1 = "=GrandTotal/100*2.23/1.2" 'OLD
        .Range(.Cells(row + 4, 10), .Cells(row + 4, 10)).FormulaR1C1 = "=GrandTotal*НР"
        
        Genpordryad_row = row + 4
        
        .Range(.Cells(row + 7, 10), .Cells(row + 7, 10)).FormulaR1C1 = "=R[1]C+R[2]C"
        .Range(.Cells(row + 8, 10), .Cells(row + 8, 10)).FormulaR1C1 = "=(R[-4]C+R[-12]C+R[-14]C)*0.2"
        .Range(.Cells(row + 9, 10), .Cells(row + 9, 10)).FormulaR1C1 = "=R[-2]C[-3]-R[-1]C"
        .Range(.Cells(row + 10, 10), .Cells(row + 10, 10)).FormulaR1C1 = "=R[-9]C+R[-3]C"
        
        .Range(.Cells(row + 1, 11), .Cells(row + 7, 11)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        .Range(.Cells(row + 10, 11), .Cells(row + 10, 11)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
        .Range(.Cells(row + 1, 12), .Cells(row + 1, 12)).FormulaR1C1 = "=RC[-5]-RC[-2]"
        .Range(.Cells(row + 5, 12), .Cells(row + 7, 12)).FormulaR1C1 = "=RC[-5]-RC[-2]"
        .Range(.Cells(row + 10, 12), .Cells(row + 10, 12)).FormulaR1C1 = "=RC[-5]-RC[-2]"
        
        .Range(.Cells(row + 1, 13), .Cells(row + 1, 13)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        .Range(.Cells(row + 5, 13), .Cells(row + 7, 13)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        .Range(.Cells(row + 10, 13), .Cells(row + 10, 13)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
    End With
End Sub
    
Private Function RangeExists(name As String) As Boolean
    Dim test As Variant
    On Error Resume Next
    Set test = ActiveWorkbook.Names.Item(name)
    RangeExists = Err.Number = 0
End Function
    
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
        .Cells(row + 1, 3).value = "ВСЕГО ЗАТРАТ В СМЕТЕ"
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Borders(xlEdgeLeft).Weight = xlMedium
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 5)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 8)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 11)).Borders(xlEdgeRight).Weight = xlMedium
        .Range(.Cells(row + 1, 1), .Cells(row + 1, 13)).Borders(xlEdgeBottom).Weight = xlMedium
        'всего затрат в смете
        
        ' Форматирование
        .Range(.Cells(row + 1, 7), .Cells(row + 1, 7)).NumberFormat = "#,##0"
        .Range(.Cells(row + 1, 8), .Cells(row + 1, 8)).NumberFormat = "0.0%"
        .Range(.Cells(row + 1, 9), .Cells(row + 1, 9)).NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
        .Range(.Cells(row + 1, 10), .Cells(row + 1, 10)).NumberFormat = "#,##0.00"
        .Range(.Cells(row + 1, 11), .Cells(row + 1, 11)).NumberFormat = "0.0%"
        .Range(.Cells(row + 1, 12), .Cells(row + 1, 12)).NumberFormat = "#,##0.00"
        .Range(.Cells(row + 1, 13), .Cells(row + 1, 13)).NumberFormat = "0.0%"
        
        ' Формулы
        nWS.Names.Add name:="GrandTotal", RefersTo:=nWS.Range(nWS.Cells(row + 1, 7), nWS.Cells(row + 1, 7))
        nWS.Names("GrandTotal").Comment = "ВСЕГО ЗАТРАТ В СМЕТЕ"
        
        .Range(.Cells(row + 1, 7), .Cells(row + 1, 7)).FormulaR1C1 = "=R[-14]C+R[-2]C"
        .Range(.Cells(row + 1, 8), .Cells(row + 1, 8)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        .Range(.Cells(row + 1, 10), .Cells(row + 1, 10)).FormulaR1C1 = "=R[-14]C+R[-2]C"
        .Range(.Cells(row + 1, 11), .Cells(row + 1, 11)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        .Range(.Cells(row + 1, 12), .Cells(row + 1, 12)).FormulaR1C1 = "=RC[-5]-RC[-2]"
        .Range(.Cells(row + 1, 13), .Cells(row + 1, 13)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        
    End With
End Sub

Function Dbl2Str(dbl) As String
    Dbl2Str = Replace(CStr(dbl), ",", ".")
End Function

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
        
        .Cells(row, 1).value = "Зам. Руководителя ДС"
        .Cells(row + 2, 1).value = "Главный инженер"
        .Cells(row + 4, 1).value = "Нач. отдела стр. аудита"
        .Cells(row + 6, 1).value = "Главный экономист"
        .Cells(row + 8, 1).value = "Руководитель управления финансов"
        .Cells(row + 9, 1).value = "и экономики ДС"
        .Cells(row, 4).value = "Павлов М.М."
        .Cells(row + 2, 4).value = "Гущин И.А."
        .Cells(row + 4, 4).value = "Игнатова Т.К."
        .Cells(row + 6, 4).value = "Кодрау И.И."
        .Cells(row + 8, 4).value = "Мамонтова А.В."
        
        .Cells(row + 1, 3).value = "(подпись)"
        .Cells(row + 1, 3).Font.Italic = True
        .Cells(row + 1, 3).Font.Size = 7
        .Cells(row + 1, 3).HorizontalAlignment = xlCenter
        .Cells(row + 1, 3).VerticalAlignment = xlTop
        .Cells(row + 3, 3).value = "(подпись)"
        .Cells(row + 3, 3).Font.Italic = True
        .Cells(row + 3, 3).Font.Size = 7
        .Cells(row + 3, 3).HorizontalAlignment = xlCenter
        .Cells(row + 3, 3).VerticalAlignment = xlTop
        .Cells(row + 5, 3).value = "(подпись)"
        .Cells(row + 5, 3).Font.Italic = True
        .Cells(row + 5, 3).Font.Size = 7
        .Cells(row + 5, 3).HorizontalAlignment = xlCenter
        .Cells(row + 5, 3).VerticalAlignment = xlTop
        .Cells(row + 7, 3).value = "(подпись)"
        .Cells(row + 7, 3).Font.Italic = True
        .Cells(row + 7, 3).Font.Size = 7
        .Cells(row + 7, 3).HorizontalAlignment = xlCenter
        .Cells(row + 7, 3).VerticalAlignment = xlTop
        .Cells(row + 9, 3).value = "(подпись)"
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
        .Cells(row, 6).value = "Показатели"
        .Cells(row, 10).value = "Коммерческая смета"
        .Cells(row, 12).value = "Утвержденный бюджет, %"
        .Range(.Cells(row, 12), .Cells(row + 1, 12)).Rows.WrapText = True
        .Cells(row, 13).value = "Отклонение, %"
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
        .Cells(row + 2, 6).value = "Финансовый результат" & Chr(10) & "(прибыль до уплаты налогов в бюджет и АУП)"
        .Cells(row + 2, 6).Rows.WrapText = True
        .Cells(row + 4, 6).value = "АУП"
        .Range(.Cells(row + 4, 6), .Cells(row + 4, 13)).Font.Italic = True
        .Cells(row + 5, 6).value = "НДС к уплате в бюджет"
        .Range(.Cells(row + 5, 6), .Cells(row + 5, 13)).Font.Italic = True
        .Cells(row + 6, 6).value = "Валовая прибыль"
        .Cells(row + 7, 6).value = "Налог на прибыль"
        .Range(.Cells(row + 7, 6), .Cells(row + 7, 13)).Font.Italic = True
        .Cells(row + 8, 6).value = "ЧИСТАЯ ПРИБЫЛЬ ОТ ПРОИЗВОДСТВА РАБОТ"
        'подвал
        
        ' Форматирование
        .Range(.Cells(row + 2, 10), .Cells(row + 8, 10)).NumberFormat = "#,##0"
        .Range(.Cells(row + 2, 11), .Cells(row + 8, 13)).NumberFormat = "0.0%"
        .Range(.Cells(row + 2, 10), .Cells(row + 8, 13)).HorizontalAlignment = xlCenter
        
        ' Формулы
        .Range(.Cells(row + 2, 10), .Cells(row + 2, 10)).FormulaR1C1 = "=R[-4]C[-3]-R[-4]C+R[-14]C+R[-7]C"
        .Range(.Cells(row + 4, 10), .Cells(row + 4, 10)).FormulaR1C1 = "=R[-16]C"
        .Range(.Cells(row + 5, 10), .Cells(row + 5, 10)).FormulaR1C1 = "=R[-10]C"
        .Range(.Cells(row + 6, 10), .Cells(row + 6, 10)).FormulaR1C1 = "=R[-4]C-R[-2]C-R[-1]C"
        .Range(.Cells(row + 7, 10), .Cells(row + 7, 10)).FormulaR1C1 = "=(R[-5]C-R[-3]C-R[-2]C)/100*20"
        .Range(.Cells(row + 8, 10), .Cells(row + 8, 10)).FormulaR1C1 = "=R[-2]C-R[-1]C"
        
        .Range(.Cells(row + 2, 11), .Cells(row + 8, 11)).FormulaR1C1 = "=RC[-1]/GrandTotal"
        .Range(.Cells(row + 6, 11), .Cells(row + 6, 11)).FormulaR1C1 = "=R[-4]C-R[-2]C-R[-1]C"
        
        '.Range(.Cells(row + 2, 12), .Cells(row + 2, 12)).Value = 0.24  'TMP
        .Range(.Cells(row + 2, 12), .Cells(row + 2, 12)).Formula = "=НДС_к_уплате_в_бюджет+Чистая_прибыль+Налог_на_прибыль"
        '.Range(.Cells(row + 4, 12), .Cells(row + 4, 12)).FormulaR1C1 = "=719230093/15519448760" 'TMP
        .Range(.Cells(row + 4, 12), .Cells(row + 4, 12)).Formula = "=АУП"
        '.Range(.Cells(row + 5, 12), .Cells(row + 5, 12)).FormulaR1C1 = "=696881732/15519448760" 'TMP
        .Range(.Cells(row + 5, 12), .Cells(row + 5, 12)).Formula = "=НДС_к_уплате_в_бюджет"
        .Range(.Cells(row + 6, 12), .Cells(row + 6, 12)).FormulaR1C1 = "=Чистая_прибыль+Налог_на_прибыль"
        '.Range(.Cells(row + 7, 12), .Cells(row + 7, 12)).FormulaR1C1 = "=463538550/15519448760" 'TMP
        .Range(.Cells(row + 7, 12), .Cells(row + 7, 12)).Formula = "=Налог_на_прибыль"
        '.Range(.Cells(row + 8, 12), .Cells(row + 8, 12)).Value = 0.119  'TMP
        .Range(.Cells(row + 8, 12), .Cells(row + 8, 12)).Formula = "=Чистая_прибыль"
        
        .Range(.Cells(row + 2, 13), .Cells(row + 8, 13)).FormulaR1C1 = "=RC[-2]-RC[-1]"
        .Range(.Cells(row + 6, 13), .Cells(row + 6, 13)).FormulaR1C1 = "=R[-4]C-R[-2]C-R[-1]C"
        
    End With
End Sub

    
