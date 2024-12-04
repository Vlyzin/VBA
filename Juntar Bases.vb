Sub Botão2_Clique()
    Dim ws As Worksheet
    Dim wsConsolidado As Worksheet
    Dim novaPasta As Workbook
    Dim ultimaLinha As Long
    Dim linhaDestino As Long
    
    Set wsConsolidado = ThisWorkbook.Worksheets.Add
    wsConsolidado.Name = "Consolidado"
    
    With ThisWorkbook.Worksheets("Macro")
        .Range("A8:T8").Copy Destination:=wsConsolidado.Cells(1, 1)
    End With
    
    linhaDestino = 2
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Macro" And ws.Name <> "Consolidado" Then
            ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ws.Range("A2:T" & ultimaLinha).Copy Destination:=wsConsolidado.Cells(linhaDestino, 1)
            linhaDestino = linhaDestino + ultimaLinha - 1
        End If
    Next ws

    Set novaPasta = Workbooks.Add
    wsConsolidado.Copy Before:=novaPasta.Sheets(1)
    novaPasta.SaveAs "C:\Users\vinicius.domingues\Documents\Projeto\Base Bayer\Teste.xlsx" 'Ajustar conforme local de usuario
    novaPasta.Close SaveChanges:=False
    
    MsgBox "Consolidação e exportação concluídas!"
End Sub
