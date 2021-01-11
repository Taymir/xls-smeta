
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
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer
    Dim ws As Worksheet: Set ws = ActiveWorkbook.ActiveSheet
    Dim blocks As Collection: Set blocks = New Collection
    Dim constr As EstimationConstructor
    Set constr = New EstimationConstructor
    constr.init ws
    firstRow = 1
    lastRow = ws.Cells(ws.Cells.Rows.Count, "A").End(xlUp).row
    'lastRow = 450 ' TMP
    
    For i = firstRow To lastRow
        ' A=1 B=1 C=1 - Название объекта
        If is_abcd(ws, i, A:=1, B:=1, C:=-1) Then
            constr.add_to_global "Name", i, G_COL
            
        ' A=52 - начало блока
        ElseIf is_abcd(ws, i, A:=52) Then
            ' 1st level (c = 1) - название сметы
            ' (c = 2) - не встречалось
            ' 2nd level (c = 3) - новая локальная смена
            ' 3rd level (c = 4) - новый раздел
            ' 4th level (c = 5) - новый подраздел
            blocks.Add (i)
            
            ' A=52 C=1 - название сметы
            If is_abcd(ws, i, C:=1) Then
                constr.add_to_global "SmetaName", i, G_COL
            ' A=52 С=3 - локальная смета
            ElseIf is_abcd(ws, i, C:=3) Then
                constr.add_to_global "LocalSmeta", i, G_COL
            ' A=52 C=4 - начало раздела
            ElseIf is_abcd(ws, i, C:=4) Then
                constr.add_section_col i, G_COL
            ' A=52 C=5 - начало подраздела
            ElseIf is_abcd(ws, i, C:=5) Then
                constr.add_subsection_col i, G_COL
            End If
            
        ' A=51 - Конец блока
        ElseIf is_abcd(ws, i, A:=51) Then
            startLine = blocks(blocks.Count)
            endLine = i
            
            If Not is_same_block(ws, startLine, endLine) Then
                Err.Raise Number:=vbObjectError + 513, _
                  Description:="Incorrect 52-51 nesting on lines: " & startLine & " and " & endLine
            End If
    
            blocks.Remove (blocks.Count)
            
        ' A=17|18 B=1 - позиции сметы
        ElseIf is_abcd(ws, i, A:=Array(17, 18), B:=1) Then
            ' Только строки с черным текстом (игнорируем синие)
            If is_black(ws, i) Then
                itemNum = ws.Cells(i, E_COL).Value
                'игнорируем позиции без номера
                If itemNum <> "" Then
                    ' Основные позиции (без запятой в номере)
                    If Not has_comma(ws, i, E_COL) Then
                        constr.add_item itemNum
                        
                        constr.add_item_vars i, E_COL, itemNum ' П/П
                        constr.add_item_vars i, F_COL, itemNum ' Шифр расценки
                        constr.add_item_vars i, G_COL, itemNum ' Наименование работ
                        
                        constr.add_item_vars i, H_COL, itemNum ' Единица измерения
                        constr.add_item_vars i, I_COL, itemNum ' количество
                        
                        constr.add_item_vars i, P_COL, itemNum
                        constr.add_item_vars i, Q_COL, itemNum
                        constr.add_item_vars i, S_COL, itemNum
                        constr.add_item_vars i, X_COL, itemNum
                        constr.add_item_vars i, Y_COL, itemNum
                        
                        constr.add_to_global "MiM", i, Q_COL
                        constr.add_to_global "ZPmas", i, R_COL
                        
                    ' субпозиции (с запятой)
                    Else
                        itemNum = Split(itemNum, ",")(0) ' оставим часть до запятой

                        constr.add_item_vars i, O_COL, itemNum
                        constr.add_item_vars i, X_COL, itemNum
                        constr.add_item_vars i, Y_COL, itemNum
                    End If ' has_comma
                    
                    If is_abcd(ws, i, A:=17) Then
                        constr.add_to_global "MR", i, P_COL
                    Else
                        constr.add_to_global "MR", i, O_COL
                    End If
                    constr.add_to_global "NR", i, X_COL
                    constr.add_to_global "SP", i, Y_COL
                End If ' itemNum <> ""
            End If ' is_black
        End If ' is_abcd(A=17|18)
    Next i
    
    'constr.test2
    constr.render
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    Debug.Print "Execution time: " & SecondsElapsed
End Sub

Public Sub test3()
    Dim e As EstimationConstructor
    Set e = New EstimationConstructor
    e.test
    
End Sub

Public Sub test4()
    Dim s As New Collection
    Debug.Print Exists(s, "test")
    s.Add 123, "test"
    Debug.Print Exists(s, "test")
    
End Sub


Function Exists(coll As Collection, ByVal key As String) As Boolean

    On Error GoTo EH

    IsObject (coll.Item(key))
    
    Exists = True
EH:
End Function


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
