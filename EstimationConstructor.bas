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


Private sWS As Worksheet
Private dWS As Worksheet
Private ivars
Private gvars
Private sects
Private name As String
Private object As String



Private Sub class_initialize()
    Set ivars = CreateObject("Scripting.Dictionary")
    Set gvars = CreateObject("Scripting.Dictionary")
    Set sects = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(sourceWS As Worksheet)
    sWS = sourceWS
End Sub

Public Sub add_col_to_vars(row, col, itemNum)
    If Not ivars.exists(itemNum) Then
        Dim itemvars: Set itemvars = CreateObject("Scripting.Dictionary")
        ivars.Add itemNum, itemvars
    End If
    If Not ivars(itemNum).exists(col) Then
        ivars(itemNum).Add col, 0
    End If
    
    ivars(itemNum)(col) = ivars(itemNum)(col) + sWS.Cells(row, col).Value
End Sub

Public Sub add_key_to_vars(val, key, itemNum)
    If Not ivars.exists(itemNum) Then
        Dim itemvars: Set itemvars = CreateObject("Scripting.Dictionary")
        ivars.Add itemNum, itemvars
    End If
    If Not ivars(itemNum).exists(key) Then
        ivars(itemNum).Add key, 0
    End If
    
    ivars(itemNum)(key) = val
End Sub

Public Sub add_to_global(name, row, col)
    gvars(name) = gvars(name) + sWS.Cells(row, col).Value
End Sub

Public Sub set_name(str As name)
    name = str
End Sub

Public Sub set_object(str As name)
    object = str
End Sub

Public Sub add_section()
    
End Sub

Public Sub add_subsection()

End Sub


Private Function get_total_for_pos(Item As Integer) As Double
    get_total_for_pos = ivars(Item)(P_COL) + _
    ivars(Item)(Q_COL) + _
    ivars(Item)(S_COL) + _
    ivars(Item)(X_COL) + _
    ivars(Item)(Y_COL) + _
    ivars(Item)(O_COL)
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

