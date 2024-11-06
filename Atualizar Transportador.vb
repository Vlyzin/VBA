Sub Atualizar_Trp()
    Dim wsBase As Worksheet
    Dim wsRetorno As Worksheet
    Dim wsMacro As Worksheet
    Dim ultimaLinhaBase As Long
    Dim ultimaLinhaRetorno As Long
    Dim i As Long
    Dim j As Long
    Dim numeroRemessa As String
    

    Set wsBase = ThisWorkbook.Sheets("Base")
    Set wsRetorno = ThisWorkbook.Sheets("Retorno")
    Set wsMacro = ThisWorkbook.Sheets("Macro")
    
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaRetorno = wsRetorno.Cells(wsRetorno.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To ultimaLinhaRetorno
        numeroRemessa = wsRetorno.Cells(i, 1).Value

        For j = 2 To ultimaLinhaBase
            If wsBase.Cells(j, 1).Value = numeroRemessa Then
    
                wsBase.Cells(j, 13).Value = wsRetorno.Cells(i, 13).Value ' Coluna M
                wsBase.Cells(j, 16).Value = wsRetorno.Cells(i, 16).Value ' Coluna P
                wsBase.Cells(j, 17).Value = wsRetorno.Cells(i, 17).Value ' Coluna Q
                wsBase.Cells(j, 18).Value = wsRetorno.Cells(i, 18).Value ' Coluna R
                Exit For
            End If
        Next j
    Next i
    

    Dim novoArquivo As Workbook
    Set novoArquivo = Workbooks.Add
    wsBase.Copy Before:=novoArquivo.Sheets(1)
    novoArquivo.SaveAs "C:\Users\vinicius.domingues\Documents\Projeto\Base Bayer\CP Report Fixo Consolidado.xlsx" ' Ajustar conforme local de usuario
    novoArquivo.Close False
    
    MsgBox "Atualização concluída e arquivo salvo!"
End Sub