Attribute VB_Name = "Módulo1"
Sub primeiro_teste()
    
    Dim limite As Long
    Dim determinante As Double
    Dim valoraleatorio As Double
    
    limite = 10
    
    For i = 1 To limite
        For j = 1 To limite
            valoraleatorio = Fix((Rnd - 0.5) * 20)
            ActiveCell.Value = valoraleatorio
            ActiveCell.Offset(1, 0).Select
        Next j
        ActiveCell.Offset(-limite, 1).Select
    Next i
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = "Determinante da Matriz:"
    Range(ActiveCell, ActiveCell).EntireColumn.AutoFit
    
    ActiveCell.Offset(0, 1).Select
    determinante = WorksheetFunction.MDeterm(Range("A1:J10"))
    ActiveCell.Value = determinante
    Range(ActiveCell, ActiveCell).EntireColumn.AutoFit
    ActiveCell.Offset(0, -12).Select

End Sub
