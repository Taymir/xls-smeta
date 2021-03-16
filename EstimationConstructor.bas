' Class EstimationConstructor
' Сохраняет данные и формирует на их основе смету

'константы с номерами столбцов
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
Const K_COL As Integer = 11
Const L_COL As Integer = 12
Const M_COL As Integer = 13
Const N_COL As Integer = 14
Const O_COL As Integer = 15
Const P_COL As Integer = 16
Const Q_COL As Integer = 17
Const R_COL As Integer = 18
Const S_COL As Integer = 19
Const T_COL As Integer = 20
Const U_COL As Integer = 21
Const V_COL As Integer = 22
Const W_COL As Integer = 23
Const X_COL As Integer = 24
Const Y_COL As Integer = 25
Const Z_COL As Integer = 26

Const EH_COL As Integer = 138
Const EM_COL As Integer = 143
Const GM_COL As Integer = 195


Private sWS As Worksheet
Private gvars
Private sects As Section
Private lastItem
Public tpl As EstimationTemplate

Private Sub Class_Initialize()
    Set gvars = CreateObject("Scripting.Dictionary")
    Set sects = New Section
End Sub

Public Sub init(sourceWS As Worksheet)
    Set sWS = sourceWS
End Sub

Public Sub add_localsmeta_col(row, col)
    add_section sWS.Cells(row, col).value
End Sub


Private Function new_section(name As String) As Section
    Set new_section = New Section
    new_section.name = name
    Set new_section.Items = New Collection
End Function


Public Sub add_section_col(row, col)
    add_section sWS.Cells(row, col).value, 2
End Sub


Public Sub add_subsection_col(row, col)
    add_section sWS.Cells(row, col).value, 3
End Sub


Private Sub add_section(name As String, Optional level As Integer = 1)
    Set sect = new_section(name)
    
    Dim where_to_add As Section
    Set where_to_add = sects
    For l = 1 To level - 1
        cnt = where_to_add.Items.Count
        Set where_to_add = where_to_add.Items(cnt)
    Next l
    where_to_add.Items.Add sect
End Sub


Private Function get_last_section() As Section
    Dim last_section As Section
    Set last_section = sects
    last_item = sects.Items.Count
    While has_subsection(last_section)
        Set last_section = last_section.Items(last_item)
        last_item = last_section.Items.Count
    Wend
    Set get_last_section = last_section
End Function

Private Function has_subsection(sect As Section) As Boolean
    cnt = sect.Items.Count
    
    If cnt > 0 Then
        If TypeOf sect.Items(cnt) Is Section Then
            has_subsection = True
            Exit Function
        End If
    End If
    has_subsection = False
End Function

Public Sub add_item(itemNum)
    Dim itm As Item
    Set itm = New Item
    itm.name = itemNum
    Set itm.Items = CreateObject("Scripting.Dictionary")
    
    Set sect = get_last_section
    sect.Items.Add itm, CStr(itemNum)
End Sub


Public Sub add_item_vars(row, col, ByVal itemNum)
    Dim sect As Section
    Dim itm As Item
    Set sect = get_last_section
    itemNum = CStr(itemNum)

    If Not HasKey(sect.Items, itemNum) Then
        add_item (itemNum)
    End If
    
    Set itm = sect.Items(itemNum)
    If Not itm.Items.Exists(col) Then
        itm.Items.Add col, sWS.Cells(row, col).value
    Else
        itm.Items(col) = itm.Items(col) + sWS.Cells(row, col).value
    End If
End Sub

Function HasKey(coll As Collection, ByVal strKey As String) As Boolean
On Error GoTo IsMissingError
        Dim val As Variant
'        val = coll(strKey)
        HasKey = IsObject(coll(strKey))
        HasKey = True
        On Error GoTo 0
        Exit Function
IsMissingError:
        HasKey = False
        On Error GoTo 0
End Function


Public Sub test()
    add_section "LocalSmeta", 1
    add_section "Главная", 2
    add_item (0)
    add_section "вторая", 2
    add_item (1)
    add_item (2)
    add_item (3)
    add_section "Третья", 2
    add_section "первый подраздел третьей секции", 3
    add_item (4)
    add_section "второй подраздел третьей секции", 3
    add_item (5)
    add_item (6)
    
    
    ' add_section LocalSmeta
    ' add_section 1
    ' end_section 1 ' go up
    ' add_section 2
    ' end_section 2 ' go up
    ' end_section LocalSmeta
    
    
    ' or use localSmeta as section?
    
    print_sects sects
End Sub

Public Sub render()
    Set tpl = New EstimationTemplate
    tpl.createBook
    tpl.render_header gvars("Name"), gvars("SmetaName")
    
    ' TODO Особые условия для 1го уровня вложенности (LocalSmeta):
    ' Если один элемент на первом уровне, то не применяем к нему
    render_section sects
        
    tpl.render_footer _
        MR:=gvars("MR"), _
        MiM:=gvars("MiM"), _
        ZPmas:=gvars("ZPmas"), _
        NR:=gvars("NR"), _
        SP:=gvars("SP"), _
        EH:=gvars("EH"), _
        EM:=gvars("EM")
End Sub

Private Sub render_section(sect, Optional level As Integer = 1)
    cnt = sect.Items.Count
    For i = 1 To cnt
        Set el = sect.Items(i)
        If TypeOf el Is Section Then
            
            If level = 1 Then
                If cnt > 1 Then
                    'FIX: smeta_mode
                    tpl.render_section el.name
                End If
            ElseIf level = 2 Then
                tpl.render_section el.name
            ElseIf level = 3 Then
                tpl.render_subsection el.name
            End If
            
            render_section el, level:=level + 1
        Else
            render_item tpl, el
        End If
    Next i
End Sub

Private Sub render_item(tpl, Item)
    num = Item.Items(E_COL)
    code = Item.Items(F_COL)
    name = Item.Items(G_COL)
    unit = Item.Items(H_COL)
    amount = Item.Items(I_COL)
    total_fot = Item.Items(S_COL)
    total = get_total_for_pos(Item.Items)
        
    units_mult = GetNumeric(unit)
    unit = GetRestPart(unit, units_mult)
    amount = amount * CInt(units_mult)

    tpl.render_item num, code, name, unit, amount, total, total_fot
End Sub


Private Sub print_sects(sect, Optional offset As String = "")
    For s = 1 To sect.Items.Count
        If TypeOf sect.Items(s) Is Section Then
            Debug.Print offset & "+" & sect.Items(s).name
            print_sects sect.Items(s), offset & "  "
        Else
            Debug.Print offset & "|" & sect.Items(s).name
        End If
    Next s
End Sub

Function GetRestPart(str, numeric_part) As String
    GetRestPart = str
    If Len(numeric_part) > 1 Then
        Start = Len(numeric_part) + 1
        GetRestPart = Mid(GetRestPart, Start)
    End If
    GetRestPart = LCase(GetRestPart)
    GetRestPart = Replace(GetRestPart, " ", "")
    GetRestPart = Replace(GetRestPart, "мп", "м")
End Function

Function GetNumeric(CellRef)
    Dim StringLength As Integer
    StringLength = Len(CellRef)
    Result = 1
    For i = 1 To StringLength
        If IsNumeric(Mid(CellRef, i, 1)) Then
            Result = Result & Mid(CellRef, i, 1)
        Else
            Exit For
        End If
    Next i
    GetNumeric = Result
End Function

Public Sub add_to_global(name, row, col)
    gvars(name) = gvars(name) + sWS.Cells(row, col).value
End Sub

Public Function get_global(name)
    get_global = gvars(name)
End Function


Private Function get_total_for_pos(Item) As Double
    get_total_for_pos = Item(P_COL) + _
    Item(Q_COL) + _
    Item(S_COL) + _
    Item(X_COL) + _
    Item(Y_COL) + _
    Item(O_COL)
    
    ' FIX for транспортный сборник не расписывается по составляющим. Если total == 0, total = GM_COL
    If get_total_for_pos = 0 Then
        get_total_for_pos = Item(GM_COL)
        ' FIX при total = 0, добавлять к гл.пер. MiM значение GM_COL
        gvars("MiM") = gvars("MiM") + Item(GM_COL)
    End If
End Function

