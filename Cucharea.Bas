Sub multiplica()
    Dim cell As Object
    Dim count As Integer
    multip = ActiveWorkbook.Sheets(5).Cells(1, 1).Value
    count = 0
    Selection.SpecialCells(xlCellTypeVisible).Select
    For Each cell In Selection
        cell.Value = Val(cell.Value) * multip
    Next cell
    


End Sub
