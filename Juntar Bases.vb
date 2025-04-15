Sub Button1_Click()
    Dim ws As Worksheet
    Dim wsConsolidado As Worksheet
    Dim wsBravo As Worksheet
    Dim ultimaLinha As Long
    Dim linhaDestino As Long
    Dim totalLinhas As Long
    Dim cel As Range
    Dim wbDestino As Workbook
    Dim wsBase As Worksheet
    Dim caminho As String
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
    Next ws
    
    On Error Resume Next
    Set wsBravo = ThisWorkbook.Worksheets("BRAVO - LEM")
    On Error GoTo 0
    
    If wsBravo Is Nothing Then
        MsgBox "Planilha 'BRAVO - LEM' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    Set wsConsolidado = ThisWorkbook.Worksheets.Add
    wsConsolidado.Name = "Consolidado"
    
    wsBravo.Range("A1:S1").Copy wsConsolidado.Range("A1")
    
    linhaDestino = 2
    totalLinhas = 0
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Consolidado" And ws.Name <> "Macro" Then
            ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            If ultimaLinha > 1 Then
                ws.Range("A2:S" & ultimaLinha).Copy wsConsolidado.Cells(linhaDestino, 1)
                totalLinhas = totalLinhas + (ultimaLinha - 1)
                linhaDestino = linhaDestino + (ultimaLinha - 1)
            End If
        End If
    Next ws
    
    If totalLinhas > 0 Then
        Dim ultimaLinhaConsolidado As Long
        ultimaLinhaConsolidado = wsConsolidado.Cells(wsConsolidado.Rows.Count, "A").End(xlUp).Row
        
        wsConsolidado.Range("S2:S" & ultimaLinhaConsolidado).Formula = _
            "=IF(ISBLANK(M2),""Aguardando Previsão/Data"",NETWORKDAYS(L2,M2)-1)"
        
        wsConsolidado.Range("O2:O" & ultimaLinhaConsolidado).Formula = _
            "=IF(P2<>""EXPEDIÇÃO"",""Expedido""," & _
            "IF(NETWORKDAYS(C2,TODAY())>=5,""Risco Alto""," & _
            "IF(NETWORKDAYS(C2,TODAY())>=3,IF(NETWORKDAYS(C2,TODAY())<5,""Risco Médio"","""")," & _
            "IF(NETWORKDAYS(C2,TODAY())<=2,""Risco Baixo"",""""))))"
        
        wsConsolidado.Range("U2:U" & ultimaLinhaConsolidado).Value = "$"
        
        For Each cel In wsConsolidado.Range("P2:P" & ultimaLinhaConsolidado)
            If cel.Value = "ENTREGUE" Or cel.Value = "CANCELADA" Or cel.Value = "RECUSADA" Then
                wsConsolidado.Cells(cel.Row, "T").Value = "*"
            End If
        Next cel
        
        caminho = "C:\Users\vinicius.domingues\Documents\Projeto\Base Bayer\Base falsa.xlsm"
        
        If Dir(caminho) = "" Then
            MsgBox "Arquivo destino não encontrado: " & caminho, vbExclamation
            Exit Sub
        End If
        
        Application.ScreenUpdating = False
        Set wbDestino = Workbooks.Open(caminho, ReadOnly:=False)
        On Error Resume Next
        Set wsBase = wbDestino.Worksheets("Base")
        On Error GoTo 0
        
        If wsBase Is Nothing Then
            MsgBox "Planilha 'Base' não encontrada no arquivo destino!", vbExclamation
            wbDestino.Close False
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        wsBase.Cells.ClearContents
        
        wsConsolidado.UsedRange.Copy wsBase.Range("A1")
        
        wbDestino.Close True
        Application.ScreenUpdating = True
    End If
    
    MsgBox "Processo concluído com sucesso!" & vbCrLf & _
           "Total de linhas consolidadas: " & totalLinhas, vbInformation
End Sub
