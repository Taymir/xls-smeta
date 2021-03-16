' Класс BugetController
' Работа с бюджетами

Public Function IsKnownBudget(name As String) As Boolean
    If getBudgetRow(name) = -1 Then
        IsKnownBudget = False
    Else
        IsKnownBudget = True
    End If
    
End Function

Private Function getBudgetRow(name As String, Optional SearchRange As String = "B:B") As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Список бюджетов")
    Set FoundCell = ws.Range(SearchRange).Find(What:=name)
    If Not FoundCell Is Nothing Then
        getBudgetRow = FoundCell.row
    Else
        getBudgetRow = -1
    End If
End Function
Public Function AliasByObjectName(name) As String
    row = getBudgetRow(name)
    If row = -1 Then
        AliasByObjectName = ""
    Else
        AliasByObjectName = ws.Range(ws.Cells(row, 1), ws.Cells(row, 1)).value
    End If
End Function

Public Function hasOffsets(name As String) As Boolean
    Dim sh As Worksheet
    If IsKnownBudget(name) Then
        Set sh = getBudgetSheet(name)
    Else
        Set sh = ThisWorkbook.Sheets("default")
    End If
    If sh.Range("C1:Q1").Find("Offset") Is Nothing Then
        hasOffsets = False
    Else
        hasOffsets = True
    End If
End Function

Public Function getBudgetItemBySmetaName(name As String, smetaName As String) As String
    Dim sh As Worksheet
    If IsKnownBudget(name) Then
        Set sh = getBudgetSheet(name)
    Else
        Set sh = ThisWorkbook.Sheets("default")
    End If
    ' Выделить номер статьи из смета нейм
    smetaName = Split(smetaName, " ", 2)(0)
    ' Искать статью
    Dim res As Range
    Set res = sh.Range("A12:A2000").Find(smetaName)
    ' Если не найдена, вернуть первую
    If Not res Is Nothing Then
        Set res = res.offset(0, 1)
    Else
        Set res = sh.Range("B12:B12")
    End If
    
    getBudgetItemBySmetaName = res.value
End Function

Public Function ObjectNameByAlias(alias As String) As String
    row = getBudgetRow(name, "A:A")
    If row = -1 Then
        ObjectNameByAlias = ""
    Else
        ObjectNameByAlias = ws.Range(ws.Cells(row, 1), ws.Cells(row, 1)).value
    End If
End Function

Public Function getBudgetSheet(name As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Список бюджетов")
    
    row = getBudgetRow(name)
    If row = -1 Then
        Err.Raise Number:=vbObjectError + 515, _
        Description:="Unknown budget name: " & name
    Else
        On Error GoTo Err
        sheetname = ws.Range(ws.Cells(row, 1), ws.Cells(row, 1)).value
        Set getBudgetSheet = ThisWorkbook.Sheets(sheetname)
    End If
        
    Exit Function
Err:
    Err.Raise Number:=vbObjectError + 516, _
        Description:="Лист с бюджетом не найден: " & name
        
End Function

Public Function addBudgetSheet(wb As Workbook, name As String) As Worksheet
    Dim sh As Worksheet
    If IsKnownBudget(name) Then
        Set sh = getBudgetSheet(name)
        sh.Copy after:=wb.Sheets(wb.Sheets.Count)
    Else
        ThisWorkbook.Sheets("default").Copy after:=wb.Sheets(wb.Sheets.Count)
        Set BudgetForm.Sheet = ActiveSheet
        BudgetForm.Show
    End If
    
    RefreshFormulas ActiveSheet.Range("A1:Q10")
    
    Set addBudgetSheet = ActiveSheet
End Function

Public Sub RefreshFormulas(refRange As Range)
    Dim rng As Range

    For Each rng In refRange
        rng.Formula = rng.Formula
    Next
End Sub
