Public ivars

Sub test()
'Set newbook = Workbooks.Add
'newbook.Windows(1).Caption = "КомСм" Название файла
Dim oldWB As Workbook: Set oldWB = ActiveWorkbook
Dim oldWS As Worksheet: Set oldWS = ActiveWorkbook.ActiveSheet

Workbooks.Add
Dim nWB As Workbook: Set nWB = ActiveWorkbook
Dim nWS As Worksheet: Set nWS = nWB.Worksheets(1)



Worksheets(1).name = "КомСм"
Worksheets(1).Rows(1).RowHeight = 27.25
Worksheets(1).Rows(2).RowHeight = 42.75

Worksheets(1).Columns(1).ColumnWidth = 3.45
Worksheets(1).Columns(2).ColumnWidth = 16.2
Worksheets(1).Columns(3).ColumnWidth = 40.3
Worksheets(1).Columns(4).ColumnWidth = 12.15
Worksheets(1).Columns(5).ColumnWidth = 9.75
Worksheets(1).Columns(6).ColumnWidth = 11.88
Worksheets(1).Columns(7).ColumnWidth = 17.6
Worksheets(1).Columns(8).ColumnWidth = 14.6
Worksheets(1).Columns(9).ColumnWidth = 16
Worksheets(1).Columns(10).ColumnWidth = 16.75
Worksheets(1).Columns(11).ColumnWidth = 14.7
Worksheets(1).Columns(12).ColumnWidth = 16.7
Worksheets(1).Columns(13).ColumnWidth = 18.3

Range("A1:J2").Merge
Range("A1:J2").HorizontalAlignment = xlCenter
Range("A1:J2").VerticalAlignment = xlCenter
Range("A1:J2").Borders.LineStyle = True
' Worksheets(1).Cells(1, 1).Value = "Объект: Реконструкция (в режиме реставрации с приспособлением к современному использованию) объекта капитального строительства " & Chr(34) & "Комплекс зданий по адресу: г. Москва, Красная площадь, д.5" & Chr(34) & " для размещения Музейно-выставочного комплекса Музеев Московского Кремля"
CopyCell oldWS.Range("G4"), nWS.Cells(1, 1)

Worksheets(1).Cells(1, 1).Font.name = "Arial"
Worksheets(1).Cells(1, 1).Font.Size = 14
Worksheets(1).Cells(1, 1).WrapText = True
Range("K1:M1").Merge
Range("K1:M2").HorizontalAlignment = xlCenter
Range("K1:M2").VerticalAlignment = xlCenter
Range("K1:M2").Borders.LineStyle = True
Range("K1:M2").Interior.Color = 3243501
Worksheets(1).Cells(1, 11).Value = "Статья Бюджета"
Worksheets(1).Cells(1, 11).Font.name = "Arial"
Worksheets(1).Cells(1, 11).Font.Size = 12
Worksheets(1).Cells(1, 11).Font.Bold = True
Worksheets(1).Cells(1, 11).Font.Italic = True
Worksheets(1).Cells(2, 11).Font.name = "Arial"
Worksheets(1).Cells(2, 11).Font.Size = 10
Worksheets(1).Cells(2, 11).Font.Bold = True
'Worksheets(1).Cells(2, 11).Formula = ""
Worksheets(1).Cells(2, 12).Font.name = "Arial"
Worksheets(1).Cells(2, 12).Font.Size = 10
Worksheets(1).Cells(2, 12).Font.Bold = True
'Worksheets(1).Cells(2, 12).Formula = ""
Worksheets(1).Cells(2, 13).Font.name = "Arial"
Worksheets(1).Cells(2, 13).Font.Size = 10
Worksheets(1).Cells(2, 13).Font.Bold = True
Worksheets(1).Cells(2, 13).WrapText = True
'Worksheets(1).Cells(2, 13).Formula = ""
Range("A3:M3").Merge
Range("A3:M3").HorizontalAlignment = xlCenter
Range("A3:M3").VerticalAlignment = xlCenter
Worksheets(1).Cells(3, 1).Value = "Согласование коммерческих расценок на выполнение работ для физических лиц"
Worksheets(1).Cells(3, 1).Font.name = "Arial"
Worksheets(1).Cells(3, 1).Font.Size = 16
Worksheets(1).Cells(3, 1).Font.Bold = True
Range("A4:M4").Merge
Range("A4:M4").HorizontalAlignment = xlCenter
Range("A4:M4").VerticalAlignment = xlCenter
' Worksheets(1).Cells(4, 1).Value = "02-02-24 ДОП Корпус Б. Устройство кровли с фонарями"
CopyCell oldWS.Range("G12"), nWS.Cells(4, 1)

Worksheets(1).Cells(4, 1).Font.name = "Arial"
Worksheets(1).Cells(4, 1).Font.Size = 14
Worksheets(1).Cells(4, 1).Font.Italic = True
Worksheets(1).Cells(4, 1).Font.Bold = True
Worksheets(1).Cells(4, 1).Font.Color = 8421504
Worksheets(1).Cells(4, 1).WrapText = True
Worksheets(1).Rows(4).RowHeight = 18.75
Worksheets(1).Rows(5).RowHeight = 13.5
Worksheets(1).Rows(6).RowHeight = 12.75
Worksheets(1).Rows(7).RowHeight = 39
Worksheets(1).Rows(8).RowHeight = 13.5
Range("A6:A7").Merge
Range("A6:M8").HorizontalAlignment = xlCenter
Range("A6:M8").VerticalAlignment = xlCenter
Range("A6:M8").WrapText = True
Range("A6:M8").Borders.LineStyle = True
Range("A6:M8").Borders.Weight = xlMedium
Range("B6:B7").Merge
Range("C6:C7").Merge
Range("D6:D7").Merge
Range("E6:E7").Merge
Range("F6:H6").Merge
Range("D6:D7").Borders(xlEdgeLeft).Weight = xlThin
Range("D6:D7").Borders(xlEdgeRight).Weight = xlThin
Range("I6:K6").Merge
Range("L6:M7").Merge
Range("G7").Borders(xlEdgeLeft).Weight = xlThin
Range("G7").Borders(xlEdgeRight).Weight = xlThin
Range("F7:H7").Borders(xlEdgeTop).Weight = xlThin
Range("J7").Borders(xlEdgeLeft).Weight = xlThin
Range("J7").Borders(xlEdgeRight).Weight = xlThin
Range("I7:K7").Borders(xlEdgeTop).Weight = xlThin
'Worksheets(1).Cells(6, 1).Font.Name = "Arial"
'Worksheets(1).Cells(6, 1).Font.Size = 10
'Worksheets(1).Cells(6, 1).Font.Bold = True
Worksheets(1).Range("A6:M8").Font.name = "Arial"
Worksheets(1).Range("A6:M8").Font.Size = 10
Worksheets(1).Range("A6:M8").Font.Bold = True
Range("D8").Borders(xlEdgeLeft).Weight = xlThin
Range("D8").Borders(xlEdgeRight).Weight = xlThin
Range("G8").Borders(xlEdgeLeft).Weight = xlThin
Range("G8").Borders(xlEdgeRight).Weight = xlThin
Range("J8").Borders(xlEdgeLeft).Weight = xlThin
Range("J8").Borders(xlEdgeRight).Weight = xlThin
Range("M8").Borders(xlEdgeLeft).Weight = xlThin
Worksheets(1).Cells(6, 1).Value = "№ п/п"
Worksheets(1).Cells(6, 2).Value = "Шифр расценки"
Worksheets(1).Cells(6, 3).Value = "Наименование работ"
Worksheets(1).Cells(6, 4).Value = "Ед. измерения"
Worksheets(1).Cells(6, 5).Value = "Кол-во"
Worksheets(1).Cells(6, 6).Value = "Локальная смета"
Worksheets(1).Cells(7, 6).Value = "Стоимость за ед."
Worksheets(1).Cells(7, 7).Value = "ИТОГО"
Worksheets(1).Cells(7, 8).Value = "% в общей сумме затрат в смете"
Worksheets(1).Cells(6, 9).Value = "Коммерческая смета"
Worksheets(1).Cells(7, 9).Value = "Стоимость за ед."
Worksheets(1).Cells(7, 10).Value = "ИТОГО"
Worksheets(1).Cells(7, 11).Value = "% в общей сумме затрат в смете"
Worksheets(1).Cells(6, 12).Value = "Финансовый результат"
Worksheets(1).Cells(8, 1).Value = 1
Worksheets(1).Cells(8, 2).Value = 2
Worksheets(1).Cells(8, 3).Value = 3
Worksheets(1).Cells(8, 4).Value = 4
Worksheets(1).Cells(8, 5).Value = 5
Worksheets(1).Cells(8, 6).Value = 6
Worksheets(1).Cells(8, 7).Value = 7
Worksheets(1).Cells(8, 8).Value = 8
Worksheets(1).Cells(8, 9).Value = 9
Worksheets(1).Cells(8, 10).Value = 10
Worksheets(1).Cells(8, 11).Value = 11
Worksheets(1).Cells(8, 12).Value = 12
Worksheets(1).Cells(8, 13).Value = 13
' заголовок отрисован

'Раздел
Range("A9:M9").Merge
Worksheets(1).Range("A9:M9").Font.name = "Arial"
Worksheets(1).Range("A9:M9").Font.Size = 14
Worksheets(1).Range("A9:M9").Font.Bold = True
Range("A9:M9").Borders.LineStyle = True
Range("A9:M9").Borders.Weight = xlMedium
'Раздел

'ячейка для расценок
Worksheets(1).Range("A10:M11").Font.name = "Arial"
Worksheets(1).Range("A10:M11").Font.Size = 10
Worksheets(1).Range("A10:M11").Font.Bold = True
Range("A10:M11").Borders.LineStyle = True
Range("A10:M11").Borders(xlEdgeLeft).Weight = xlMedium
Range("A10:M11").Borders(xlEdgeRight).Weight = xlMedium
Range("A10:M11").Borders(xlEdgeBottom).Weight = xlThin
Range("A10:M10").Borders(xlEdgeBottom).LineStyle = xlDot
Range("A10:M10").Borders(xlEdgeBottom).Weight = xlThin
'ячейка для расценок

'ИТОГО по смете
Range("A12:M12").Merge
Range("A12:M12").Borders.LineStyle = True
Range("A12:M12").Borders.Weight = xlMedium
Worksheets(1).Rows(12).RowHeight = 13.5

Range("A13:M14").Borders.LineStyle = True
Worksheets(1).Range("A13:M14").Font.name = "Arial"
Worksheets(1).Range("A13:M13").Font.Size = 11
Worksheets(1).Range("A14:M14").Font.Size = 10
Worksheets(1).Range("A13:M13").Font.Bold = True
Range("A13:M14").Borders(xlEdgeLeft).Weight = xlMedium
Range("A13:M14").Borders(xlEdgeRight).Weight = xlMedium
Range("A13:M14").Borders(xlEdgeBottom).Weight = xlMedium
Range("A13:F13").Merge
Range("A14:F14").Merge
Range("A14:M14").HorizontalAlignment = xlRight
Worksheets(1).Range("A14:M14").Font.Italic = True
Worksheets(1).Cells(13, 1).Value = "Итого по смете:"
Worksheets(1).Cells(14, 1).Value = "в т.ч. ФОТ"

Range("A15:M15").Merge
Range("A15:M15").Borders.LineStyle = True
Range("A15:M15").Borders.Weight = xlMedium
Worksheets(1).Rows(15).RowHeight = 13.5
'ИТОГО по смете

'Свод прямых затрат в смете
Range("A16:M24").Borders.LineStyle = True
Range("A16:M24").Interior.Color = 14809087
Range("C16:M16").Merge
Range("C16:C24").IndentLevel = 1
Range("A16:M24").Borders(xlEdgeLeft).Weight = xlMedium
Range("A16:M24").Borders(xlEdgeRight).Weight = xlMedium
Range("A16:E24").Borders(xlEdgeRight).Weight = xlMedium
Range("A16:H24").Borders(xlEdgeRight).Weight = xlMedium
Range("A16:K24").Borders(xlEdgeRight).Weight = xlMedium
Range("A16:M24").Borders(xlEdgeBottom).Weight = xlMedium

Worksheets(1).Range("A16:M24").Font.name = "Arial"
Worksheets(1).Range("A16:M24").Font.Size = 11
Worksheets(1).Range("A16:M16").Font.Bold = True
Worksheets(1).Range("A16:M16").Font.Size = 14
Worksheets(1).Range("A24:M24").Font.Bold = True
Worksheets(1).Range("A24:M24").Font.Size = 12

Range("A16:A16").HorizontalAlignment = xlCenter
Worksheets(1).Cells(16, 1).Value = "II"
Worksheets(1).Cells(16, 3).Value = "Свод прямых затрат в смете"
Worksheets(1).Cells(17, 3).Value = "ФОТ по позициям"
Worksheets(1).Cells(18, 3).Value = "Вывоз мусора"
Worksheets(1).Cells(19, 3).Value = "Материальные ресурсы"
Worksheets(1).Cells(20, 3).Value = "Субподряд"
Range("C21:M21").Rows.WrapText = True
Worksheets(1).Cells(21, 3).Value = "Машины, механизмы, з/п механизаторов, в т.ч.:"
Range("C22:C23").HorizontalAlignment = xlRight
Worksheets(1).Cells(22, 3).Value = "аренда машин и механизмов"
Worksheets(1).Cells(22, 3).Font.Italic = True
Worksheets(1).Cells(23, 3).Value = "з/п машинистов"
Worksheets(1).Cells(23, 3).Font.Italic = True
Worksheets(1).Cells(24, 3).Value = "Итого прямых затрат в смете"

Range("A25:M25").Merge
Range("A25:M25").Borders.LineStyle = True
Range("A25:M25").Borders.Weight = xlMedium
Worksheets(1).Rows(25).RowHeight = 13.5
'Свод прямых затрат в смете

'Свод дополнительных затрат в смете
Range("A26:M36").Borders.LineStyle = True
Range("C26:M26").Merge
Range("A26:M36").Borders(xlEdgeLeft).Weight = xlMedium
Range("A26:M36").Borders(xlEdgeRight).Weight = xlMedium
Range("A26:E36").Borders(xlEdgeRight).Weight = xlMedium
Range("A26:H36").Borders(xlEdgeRight).Weight = xlMedium
Range("A26:K36").Borders(xlEdgeRight).Weight = xlMedium
Range("A26:M36").Borders(xlEdgeBottom).Weight = xlMedium
Range("C26:C36").IndentLevel = 1

Worksheets(1).Range("A26:M36").Font.name = "Arial"
Worksheets(1).Range("A26:M36").Font.Size = 11
Worksheets(1).Range("A28:M30").Font.Size = 10
Worksheets(1).Range("A36:M36").Font.Size = 12

Worksheets(1).Range("A26:M26").Font.Bold = True
Worksheets(1).Range("A26:M26").Font.Size = 14
Range("A26:A26").HorizontalAlignment = xlCenter
Worksheets(1).Cells(26, 1).Value = "III"
Worksheets(1).Cells(26, 3).Value = "Свод дополнительных затрат в смете"
Worksheets(1).Cells(27, 3).Value = "Накладные расходы, в том числе:"
Range("C28:C30").Rows.WrapText = True
Range("C28:C30").Font.Italic = True
Range("C28:C30").Font.Color = 7434613
Range("C28:C30").IndentLevel = 2
Worksheets(1).Cells(28, 3).Value = "Административно-хозяйственные расходы (5% от сметы)"
Worksheets(1).Cells(29, 3).Value = "Расходы на обслуживание работников строительства"
Worksheets(1).Cells(30, 3).Value = "Расходы на организацию работ на строительных площадках (2,48% от сметы)"
Worksheets(1).Cells(31, 3).Value = "Сметная прибыль"
Worksheets(1).Cells(32, 3).Value = "Зимнее удорожание 1,41%"
Worksheets(1).Cells(33, 3).Value = "НДС 20%, в т.ч."
Worksheets(1).Cells(34, 3).Font.Color = 7434613
Worksheets(1).Cells(34, 3).Value = "НДС уплаченный поставщикам"
Worksheets(1).Cells(34, 3).Font.Italic = True
Worksheets(1).Cells(35, 3).Font.Color = 255
Worksheets(1).Cells(35, 3).Value = "НДС к уплате в бюджет"
Worksheets(1).Cells(35, 3).Font.Italic = True
Worksheets(1).Cells(35, 3).HorizontalAlignment = xlRight
Worksheets(1).Cells(36, 3).Font.Size = 14
Worksheets(1).Cells(36, 3).Font.Bold = True
Worksheets(1).Cells(36, 3).Value = "Итого дополнительных затрат в смете"
Range("C36:C36").Rows.WrapText = True
'Свод дополнительных затрат в смете

'всего затрат в смете
Range("A37:M37").Merge
Range("A37:M37").Borders.LineStyle = True
Range("A37:M37").Borders.Weight = xlMedium
Worksheets(1).Rows(37).RowHeight = 13.5
Range("A38:M38").Borders.LineStyle = True
Worksheets(1).Range("A38:M38").Font.name = "Arial"
Worksheets(1).Range("A38:M38").Font.Bold = True
Worksheets(1).Range("A38:M38").Font.Size = 12
Worksheets(1).Range("A38:M38").Font.Italic = True
Worksheets(1).Cells(38, 3).HorizontalAlignment = xlRight
Worksheets(1).Cells(38, 3).Value = "ВСЕГО ЗАТРАТ В СМЕТЕ"
Range("A38:M38").Borders(xlEdgeLeft).Weight = xlMedium
Range("A38:M38").Borders(xlEdgeRight).Weight = xlMedium
Range("A38:E38").Borders(xlEdgeRight).Weight = xlMedium
Range("A38:H38").Borders(xlEdgeRight).Weight = xlMedium
Range("A38:K38").Borders(xlEdgeRight).Weight = xlMedium
Range("A38:M38").Borders(xlEdgeBottom).Weight = xlMedium
'всего затрат в смете

'подвал
Range("A40:C40").Borders(xlEdgeBottom).Weight = xlThin
Range("A42:C42").Borders(xlEdgeBottom).Weight = xlThin
Range("A44:C44").Borders(xlEdgeBottom).Weight = xlThin
Range("A46:C46").Borders(xlEdgeBottom).Weight = xlThin
Range("A48:C48").Borders(xlEdgeBottom).Weight = xlThin

Worksheets(1).Range("A40:E49").Font.name = "Arial"
Worksheets(1).Range("A40:A49").Font.Bold = True
Worksheets(1).Range("A40:A49").Font.Size = 10
Worksheets(1).Range("D40:D49").Font.Size = 14

Worksheets(1).Cells(40, 1).Value = "Зам. Руководителя ДС"
Worksheets(1).Cells(42, 1).Value = "Главный инженер"
Worksheets(1).Cells(44, 1).Value = "Нач. отдела стр. аудита"
Worksheets(1).Cells(46, 1).Value = "Главный экономист"
Worksheets(1).Cells(48, 1).Value = "Руководитель управления финансов"
Worksheets(1).Cells(49, 1).Value = "и экономики ДС"
Worksheets(1).Cells(40, 4).Value = "Павлов М.М."
Worksheets(1).Cells(42, 4).Value = "Гущин И.А."
Worksheets(1).Cells(44, 4).Value = "Игнатова Т.К."
Worksheets(1).Cells(46, 4).Value = "Кодрау И.И."
Worksheets(1).Cells(48, 4).Value = "Мамонтова А.В."

Worksheets(1).Cells(41, 3).Value = "(подпись)"
Worksheets(1).Cells(41, 3).Font.Italic = True
Worksheets(1).Cells(41, 3).Font.Size = 7
Worksheets(1).Cells(41, 3).HorizontalAlignment = xlCenter
Worksheets(1).Cells(41, 3).VerticalAlignment = xlTop
Worksheets(1).Cells(43, 3).Value = "(подпись)"
Worksheets(1).Cells(43, 3).Font.Italic = True
Worksheets(1).Cells(43, 3).Font.Size = 7
Worksheets(1).Cells(43, 3).HorizontalAlignment = xlCenter
Worksheets(1).Cells(43, 3).VerticalAlignment = xlTop
Worksheets(1).Cells(45, 3).Value = "(подпись)"
Worksheets(1).Cells(45, 3).Font.Italic = True
Worksheets(1).Cells(45, 3).Font.Size = 7
Worksheets(1).Cells(45, 3).HorizontalAlignment = xlCenter
Worksheets(1).Cells(45, 3).VerticalAlignment = xlTop
Worksheets(1).Cells(47, 3).Value = "(подпись)"
Worksheets(1).Cells(47, 3).Font.Italic = True
Worksheets(1).Cells(47, 3).Font.Size = 7
Worksheets(1).Cells(47, 3).HorizontalAlignment = xlCenter
Worksheets(1).Cells(47, 3).VerticalAlignment = xlTop
Worksheets(1).Cells(49, 3).Value = "(подпись)"
Worksheets(1).Cells(49, 3).Font.Italic = True
Worksheets(1).Cells(49, 3).Font.Size = 7
Worksheets(1).Cells(49, 3).HorizontalAlignment = xlCenter
Worksheets(1).Cells(49, 3).VerticalAlignment = xlTop

Range("F40:M48").Borders.LineStyle = True
Range("F40:M48").Borders.Weight = xlMedium
Range("F43:M43").Borders(xlEdgeBottom).Weight = xlThin
Range("F44:M44").Borders(xlEdgeBottom).Weight = xlThin
Range("F46:M46").Borders(xlEdgeBottom).Weight = xlThin
Range("F42:J48").Borders(xlEdgeRight).Weight = xlThin
Range("F40:I41").Merge
Range("J40:K41").Merge
Range("L40:L41").Merge
Range("M40:M41").Merge
Worksheets(1).Range("F40:M49").Font.name = "Arial"
Worksheets(1).Range("F40:M43").Font.Bold = True
Worksheets(1).Range("F46:M46").Font.Bold = True
Worksheets(1).Range("F48:M48").Font.Bold = True
Worksheets(1).Range("F40:M49").Font.Size = 11
Worksheets(1).Range("F40:M49").VerticalAlignment = xlCenter
Range("F40:M41").HorizontalAlignment = xlCenter
Range("F40:M41").VerticalAlignment = xlCenter
Worksheets(1).Cells(40, 6).Value = "Показатели"
Worksheets(1).Cells(40, 10).Value = "Коммерческая смета"
Worksheets(1).Cells(40, 12).Value = "Утвержденный бюджет, %"
Range("L40:L41").Rows.WrapText = True
Worksheets(1).Cells(40, 13).Value = "Отклонение, %"
Range("F42:I43").Merge
Range("J42:J43").Merge
Range("K42:K43").Merge
Range("L42:L43").Merge
Range("M42:M43").Merge
Range("F44:I44").Merge
Range("F45:I45").Merge
Range("F46:I46").Merge
Range("F47:I47").Merge
Range("F48:I48").Merge
Range("F40:I48").IndentLevel = 2
Range("F42:I43").IndentLevel = 1
Range("F46:I46").IndentLevel = 1
Range("F48:I48").IndentLevel = 1
Range("F42:M43").Interior.Color = 16247773
Range("F48:M48").Interior.Color = 16247773
Range("F46:M46").Interior.Color = 16247773
Worksheets(1).Cells(42, 6).Value = "Финансовый результат" & Chr(10) & "(прибыль до уплаты налогов в бюджет и АУП)"
Worksheets(1).Cells(42, 6).Rows.WrapText = True
Worksheets(1).Cells(44, 6).Value = "АУП"
Worksheets(1).Cells(44, 6).Font.Italic = True
Worksheets(1).Cells(45, 6).Value = "НДС к уплате в бюджет"
Worksheets(1).Cells(45, 6).Font.Italic = True
Worksheets(1).Cells(46, 6).Value = "Валовая прибыль"
Worksheets(1).Cells(47, 6).Value = "Налог на прибыль"
Worksheets(1).Cells(47, 6).Font.Italic = True
Worksheets(1).Cells(48, 6).Value = "ЧИСТАЯ ПРИБЫЛЬ ОТ ПРОИЗВОДСТВА РАБОТ"
'подвал



End Sub

Sub CopyCell(FromRange, ToRange)
    v = FromRange.Value
    ToRange.Value = "copied" & v
End Sub

Sub ParseSource()
Const A_COL As Integer = 1
Const B_COL As Integer = 2
Const C_COL As Integer = 3
Const D_COL As Integer = 4
Const E_COL As Integer = 5
Const F_COL As Integer = 6
Const G_COL As Integer = 7
Const H_COL As Integer = 8
Const I_COL As Integer = 9
Const J_COL As Integer = 10

Const O_COL As Integer = 15
Const P_COL As Integer = 16
Const Q_COL As Integer = 17
Const R_COL As Integer = 18
Const S_COL As Integer = 19

Const X_COL As Integer = 24
Const Y_COL As Integer = 25

Set ivars = CreateObject("Scripting.Dictionary")
Set glvars = CreateObject("Scripting.Dictionary")

Dim ws As Worksheet: Set ws = ActiveWorkbook.ActiveSheet
Offset = 1
lastRow = ws.Cells(ws.Cells.Rows.Count, "A").End(xlUp).row
Dim bl As Collection: Set bl = New Collection
'Dim ivars As Collection: Set ivars = New Collection
'Dim ci As Integer
'ci = 1
'Dim in_section: Set in_section = False
currentItem = -1
' lastRow = 350 ' tmp
For i = Offset To lastRow
    If is_abcd(ws, i, A:=1, B:=1, C:=-1) Then
        Debug.Print "Название объекта: " & ws.Cells(i, G_COL).Value & " |#" & i
    ElseIf is_abcd(ws, i, A:=52) Then
        ' 1st level (c = 1) - название сметы
        ' 2nd level (c = 3) - новая локальная смена
        ' 3rd level (c = 4) - новый раздел
        ' 4th level (c = 5) - новый подраздел
        bl.Add (i)

        If is_abcd(ws, i, C:=1) Then
            Debug.Print "Раздел Название сметы: " & ws.Cells(i, G_COL).Value & " |#" & i
        ElseIf is_abcd(ws, i, C:=3) Then
            Debug.Print "Локальная смета: " & ws.Cells(i, G_COL).Value & " |#" & i
        ElseIf is_abcd(ws, i, C:=4) Then
            Debug.Print "Новые раздел: " & ws.Cells(i, G_COL).Value & " |#" & i
            'in_section = True
        ElseIf is_abcd(ws, i, C:=5) Then
            Debug.Print "Новый подраздел: " & ws.Cells(i, G_COL).Value & " |#" & i
            'in_section = True
        End If
    ElseIf is_abcd(ws, i, A:=51) Then
        startLine = bl(bl.Count)
        endLine = i
        Level = bl.Count
        
        If is_same_block(ws, startLine, endLine) Then
            'colorize
            'ws.Range(ws.Cells(startLine, bl.Count), ws.Cells(endLine, bl.Count)).Interior.ColorIndex = ci
            If is_abcd(ws, i, C:=1) Then
                Debug.Print "Конец сметы: " & ws.Cells(i, G_COL).Value & " |#" & i
            ElseIf is_abcd(ws, i, C:=3) Then
                Debug.Print "Конец локальной сметы: " & ws.Cells(i, G_COL).Value & " |#" & i
            ElseIf is_abcd(ws, i, C:=4) Then
                Debug.Print "Конец раздела: " & ws.Cells(i, G_COL).Value & " |#" & i
            ElseIf is_abcd(ws, i, C:=5) Then
                Debug.Print "Конец подраздела: " & ws.Cells(i, G_COL).Value & " |#" & i
                'in_section = False
            End If
        Else
            Err.Raise Number:=vbObjectError + 513, _
              Description:="Incorrect 52-51 nesting on lines: " & startLine & " and " & endLine
        End If

        'ci = ci + 1
        bl.Remove (bl.Count)
    ElseIf is_abcd(ws, i, A:=Array(17, 18), B:=1) Then
        ' Игнорируем синие строки
        If is_black(ws, i) Then
            ' Основные позиции (без запятой)
            If Not has_comma(ws, i, E_COL) Then
                itemNum = ws.Cells(i, E_COL).Value
                
                If currentItem > 0 And currentItem <> itemNum Then
                    'подвести итоги по предыдущей позиции
                    Total = ivars(currentItem)(P_COL) + _
                    ivars(currentItem)(Q_COL) + _
                    ivars(currentItem)(S_COL) + _
                    ivars(currentItem)(X_COL) + _
                    ivars(currentItem)(Y_COL) + _
                    ivars(currentItem)(O_COL)
                    Debug.Print ("Total for " & currentItem & ": " & Total)
                    Debug.Print ("7-11 for " & currentItem & ": " & ivars(currentItem)(S_COL))
                End If
                currentItem = itemNum
                
                Debug.Print "Работы: " & ws.Cells(i, G_COL).Value & " |#" & i
                Debug.Print "П/П: " & ws.Cells(i, E_COL).Value
                Debug.Print "Шифр расценки: " & ws.Cells(i, F_COL).Value
                Debug.Print "Наименование работ: " & ws.Cells(i, G_COL).Value
                Debug.Print "Ед. измерения: " & ws.Cells(i, H_COL).Value
                Debug.Print "Количество: " & ws.Cells(i, I_COL).Value
                
                'Dim itemvars As Collection: Set itemvars = New Collection
                'ivars.Add itemvars, itemNum
                
                
                add_to_vars ws, i, P_COL, itemNum
                add_to_vars ws, i, Q_COL, itemNum
                add_to_vars ws, i, S_COL, itemNum
                add_to_vars ws, i, X_COL, itemNum
                add_to_vars ws, i, Y_COL, itemNum
                
                glvars("МИМ") = glvars("МИМ") + ws.Cells(i, Q_COL).Value
                glvars("ЗПМАШ") = glvars("ЗПМАШ") + ws.Cells(i, R_COL).Value
                
            Else
            ' субпозиции (с запятой)
            If currentItem > 0 And currentItem <> itemNum Then
                    'подвести итоги по предыдущей позиции 'TMP переделать
                    Total = ivars(currentItem)(P_COL) + _
                    ivars(currentItem)(Q_COL) + _
                    ivars(currentItem)(S_COL) + _
                    ivars(currentItem)(X_COL) + _
                    ivars(currentItem)(Y_COL) + _
                    ivars(currentItem)(O_COL)
                    Debug.Print ("Total for " & currentItem & ": " & Total)
                End If
                currentItem = itemNum
                
                itemNum = ws.Cells(i, E_COL).Value
                itemNum = Split(itemNum, ",")(0)
                
                add_to_vars ws, i, O_COL, itemNum
                add_to_vars ws, i, X_COL, itemNum
                add_to_vars ws, i, Y_COL, itemNum
            End If ' итоги по пред. позиции
            If is_abcd(ws, i, A:=17) Then
                glvars("МР") = glvars("МР") + ws.Cells(i, P_COL).Value ' or use O in both cases?
            ElseIf is_abcd(ws, i, A:=18) Then
                glvars("МР") = glvars("МР") + ws.Cells(i, O_COL).Value
            End If ' is_abcd(A=17)
            
            glvars("НАКЛРАСХ") = glvars("НАКЛРАСХ") + ws.Cells(i, X_COL).Value
            glvars("СМЕТПРИБ") = glvars("СМЕТПРИБ") + ws.Cells(i, Y_COL).Value
        End If ' is_black
    End If ' is_abcd(A=17|18)
Next i

TraverseDictionary glvars
End Sub

Private Sub TraverseDictionary(d, Optional indention As String = " ", Optional ByVal i = 1, Optional ByVal depth = 0)

    For Each key In d.Keys
        Debug.Print (vbNewLine & indention & key & ":");
        If VarType(d(key)) = 9 Then
            depth = depth + 1
            TraverseDictionary d(key), indention & "    ", i, depth
        Else
            Debug.Print (" " & d(key))
        End If
        i = i + 1
    Next
End Sub

'Dim ivars As Collection: Set ivars = New Collection

Sub add_to_vars(ws, row, col, itemNum)
    'If ivars Is Nothing Then
        'Set ivars = CreateObject("Scripting.Dictionary")
    'End If
    If Not ivars.exists(itemNum) Then
        Dim itemvars: Set itemvars = CreateObject("Scripting.Dictionary")
        ivars.Add itemNum, itemvars
    End If
    If Not ivars(itemNum).exists(col) Then
        ivars(itemNum).Add col, 0
    End If
        
    ivars(itemNum)(col) = ivars(itemNum)(col) + ws.Cells(row, col).Value
End Sub



Function is_same_block(ws, row_1, row_2) As Boolean
    With ws
        If .Cells(row_1, 2) = .Cells(row_2, 2) And _
           .Cells(row_1, 3) = .Cells(row_2, 3) And _
           .Cells(row_1, 4) = .Cells(row_2, 4) Then
            is_same_block = True
        Else
            is_same_block = False
        End If
    End With
End Function

Function is_abcd(ws, row, Optional ByVal A As Variant = Null, Optional ByVal B As Variant = Null, Optional ByVal C As Variant = Null, Optional ByVal d As Variant = Null) As Boolean
    ret_flag = True
    vars = Array(A, B, C, d)
    
    For i = 0 To 3
        If Not IsNull(vars(i)) Then
            If IsArray(vars(i)) Then
                ret_flag = ret_flag And cell_in_array(ws, row, i + 1, vars(i))
            Else
                ret_flag = ret_flag And cell_equals(ws, row, i + 1, vars(i))
            End If
        End If
    Next i
    is_abcd = ret_flag
End Function

Private Function cell_equals(ws, row, col, val) As Boolean
    cell_equals = ws.Cells(row, col).Value = val
End Function

Private Function cell_in_array(ws, row, col, arr) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = ws.Cells(row, col).Value Then
            cell_in_array = True
            Exit Function
        End If
    Next i
    cell_in_array = False
End Function

Private Function is_black(ws, row) As Boolean
    is_black = ws.Cells(row, 7).Font.Color = 0
End Function

Private Function has_comma(ws, row, Optional ByVal col = 5) As Boolean
    has_comma = InStr(ws.Cells(row, col).Value, ",") > 0
End Function
