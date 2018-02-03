Sub EliminarRepetidos()
    Dim Row As Long
    Dim ws As Worksheet

    Set ws = ActiveSheet
    Row = 1
    
    
    
    'Apagar visualización de saltos de página
    ActiveSheet.DisplayPageBreaks = False
    
    'Apagar parpadeo de pantalla
    Application.ScreenUpdating = False
    
    'Minimiza la ventana para mayor velocidad
    ActiveWindow.WindowState = xlMinimized
    
    While ws.Cells(Row, 2).Value <> ""
        If ws.Cells(Row, 2).Value = ws.Cells(Row + 1, 2).Value Then
            Range("B" & (Row + 1)).EntireRow.Delete
        Else
            Row = Row + 1
        End If
    Wend
    'Se maximiza cuando termine el while
    ActiveWindow.WindowState = xlMaximized
     Application.ScreenUpdating = True
    ActiveSheet.DisplayPageBreaks = True
    

End Sub
