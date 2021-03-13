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
Private sects As Collection
Private lastItem
Public tpl As EstimationTemplate

Private Sub Class_Initialize()
    Set gvars = CreateObject("Scripting.Dictionary")
    Set sects = New Collection
End Sub

Public Sub init(sourceWS As Worksheet)
    Set sWS = sourceWS
End Sub


Private Function is_subsection(Optional ByVal current As Integer = -1, Optional ByVal curitem As Integer = -1) As Boolean
    is_subsection = False
    
    If current = -1 Then
        current = sects.Count
    End If
    If curitem = -1 Then
        curitem = sects(current).Items.Count
    End If
    If curitem > 0 Then
        If TypeOf sects(current).Items(curitem) Is Section Then
            is_subsection = True
        End If
    End If
End Function
Public Sub add_section_col(row, col)
    add_section (sWS.Cells(row, col).value)
End Sub

Public Sub add_section(name As String)
    Dim sect As Section
    Set sect = New Section
    
    sect.name = name
    Set sect.Items = New Collection
    
    sects.Add sect
End Sub

Public Sub add_subsection_col(row, col)
    add_subsection (sWS.Cells(row, col).value)
End Sub

Public Sub add_subsection(name)
    Dim sect As Section
    Set sect = New Section
    
    sect.name = name
    Set sect.Items = New Collection
    
    current = sects.Count
    sects(current).Items.Add sect
End Sub

Private Function get_last_section() As Section
    last_section = sects.Count
    
    'HARDFIX TMP: Если item добавляется в отсутствии sect-ов, добавим элемент LocalSmeta в качестве 0-го раздела
    If last_section = 0 Then
        Dim sect As Section
        Set sect = New Section
        
        sect.name = "LocalSmeta"
        Set sect.Items = New Collection
        
        sects.Add sect
        last_section = 1
    End If
    
    last_item = sects(last_section).Items.Count
    
    If is_subsection(last_section, last_item) Then
        Set get_last_section = sects(last_section).Items(last_item)
    Else
        Set get_last_section = sects(last_section)
    End If
End Function

Public Sub add_item(itemNum)
    'last_section = sects.Count
    'last_item = sects(last_section).items.Count
    
    Dim itm As Item
    Set itm = New Item
    itm.name = itemNum
    Set itm.Items = CreateObject("Scripting.Dictionary")
    
    Set sect = get_last_section
    sect.Items.Add itm, CStr(itemNum)
    
    'If is_subsection(last_section, last_item) Then
    '    sects(last_section).items(last_item).items.Add itm, CStr(itemNum)
    'Else
    '    sects(last_section).items.Add itm, CStr(itemNum)
    'End If
End Sub


Public Sub add_item_vars(row, col, ByVal itemNum)
    Set sect = get_last_section
    itemNum = CStr(itemNum)

    If Not HasKey(sect.Items, itemNum) Then
        add_item (itemNum)
    End If
    If Not sect.Items(itemNum).Items.Exists(col) Then
        sect.Items(itemNum).Items.Add col, sWS.Cells(row, col).value
    Else
        sect.Items(itemNum).Items(col) = sect.Items(itemNum).Items(col) + sWS.Cells(row, col).value
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

Public Sub test2()
    print_sects
End Sub

Public Sub Test()
    add_section ("Главная")
    add_item (0)
    add_section ("вторая")
    add_item (1)
    add_item (2)
    add_item (3)
    add_section ("Третья")
    add_subsection ("первый подраздел третьей секции")
    add_item (4)
    add_subsection ("второй подраздел третьей секции")
    add_item (5)
    add_item (6)
    
    
    print_sects
End Sub

Public Sub render()
    Set tpl = New EstimationTemplate
    tpl.createBook
    tpl.render_header gvars("Name"), gvars("SmetaName")
    
    For s = 1 To sects.Count
        tpl.render_section sects(s).name
        
        For i = 1 To sects(s).Items.Count
            If is_subsection(s, i) Then
                tpl.render_subsection sects(s).Items(i).name
                
                For ii = 1 To sects(s).Items(i).Items.Count
                    render_item tpl, sects(s).Items(i).Items(ii)
                Next ii
            Else ' is not subsection == is item
                render_item tpl, sects(s).Items(i)
            End If
        Next i
    Next s
        
    tpl.render_footer _
        MR:=gvars("MR"), _
        MiM:=gvars("MiM"), _
        ZPmas:=gvars("ZPmas"), _
        NR:=gvars("NR"), _
        SP:=gvars("SP"), _
        EH:=gvars("EH"), _
        EM:=gvars("EM")
    
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

Private Sub print_sects()
    For s = 1 To sects.Count
        Debug.Print "+" & sects(s).name
        For i = 1 To sects(s).Items.Count
            If is_subsection(s, i) Then
                'Debug.Print "subsection"
                Debug.Print "  +" & sects(s).Items(i).name
                For ii = 1 To sects(s).Items(i).Items.Count
                    Debug.Print "  |" & sects(s).Items(i).Items(ii).name
                    'TraverseDictionary (sects(s).items(i).items(ii).items)
                    Debug.Print get_total_for_pos(sects(s).Items(i).Items(ii).Items)
                Next ii
                Debug.Print "  |___"
            Else
                'Debug.Print "not subsection"
                Debug.Print "|" & sects(s).Items(i).name
                Debug.Print get_total_for_pos(sects(s).Items(i).Items)
            End If
        Next i
        Debug.Print "|___"
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


Public Sub set_name(str As name)
    name = str
End Sub

Public Sub set_object(str As name)
    object = str
End Sub


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

Private Sub TraverseDictionary(d, Optional indention As String = " ", Optional ByVal i = 1, Optional ByVal depth = 0)

    For Each Key In d.Keys
        Debug.Print (vbNewLine & indention & Key & ":");
        If VarType(d(Key)) = 9 Then
            depth = depth + 1
            TraverseDictionary d(Key), indention & "    ", i, depth
        Else
            Debug.Print (" " & d(Key))
        End If
        i = i + 1
    Next
End Sub

