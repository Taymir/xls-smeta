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
    Dim FoundCell As Range
    Set ws = ThisWorkbook.Sheets("Список бюджетов")
    Set FoundCell = ws.Range(SearchRange).Find(What:=name)
    If Not FoundCell Is Nothing Then
        getBudgetRow = FoundCell.row
    Else
        getBudgetRow = -1
    End If
End Function
Public Function AliasByObjectName(name As String) As String
    row = getBudgetRow(name)
    If row = -1 Then
        AliasByObjectName = ""
    Else
        AliasByObjectName = ws.Range(ws.Cells(row, 1), ws.Cells(row, 1)).value
    End If
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
