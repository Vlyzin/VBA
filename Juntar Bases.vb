Sub Button11_Click()
'Juntar bases
    Dim wbOrigem As Workbook
    Dim wbDestino As Workbook
    Dim wsOrigem As Worksheet
    Dim wsConsolidado As Worksheet
    Dim wsBase As Worksheet
    Dim arquivo As Variant
    Dim ultimaLinha As Long
    Dim rngCopiar As Range
    Dim cel As Range
    
    Set wbDestino = ThisWorkbook
    
    arquivo = Application.GetOpenFilename("Arquivos Excel (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls", , "Selecione o arquivo para consolidar")
    
    If arquivo = False Then
        MsgBox "Nenhum arquivo selecionado. A macro será cancelada.", vbExclamation
        Exit Sub
    End If
    
    Set wbOrigem = Workbooks.Open(arquivo)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each wsOrigem In wbOrigem.Worksheets
        If wsOrigem.AutoFilterMode Then
            wsOrigem.AutoFilterMode = False
        End If
    Next wsOrigem
    
    On Error Resume Next
    Set wsConsolidado = wbOrigem.Worksheets("Consolidado")
    On Error GoTo 0
    
    If Not wsConsolidado Is Nothing Then
        Application.DisplayAlerts = False
        wsConsolidado.Delete
        Application.DisplayAlerts = True
    End If
    
    Set wsConsolidado = wbOrigem.Worksheets.Add(After:=wbOrigem.Worksheets(wbOrigem.Worksheets.Count))
    wsConsolidado.Name = "Consolidado"
    
    On Error Resume Next
    Set wsOrigem = wbOrigem.Worksheets("BRAVO - LEM")
    On Error GoTo 0
    
    If wsOrigem Is Nothing Then
        MsgBox "A aba 'BRAVO - LEM' não foi encontrada no arquivo selecionado.", vbExclamation
        GoTo FecharArquivo
    End If
    
    wsOrigem.Range("A1:S1").Copy wsConsolidado.Range("A1")
    
    For Each wsOrigem In wbOrigem.Worksheets
        If wsOrigem.Name <> "Consolidado" Then

            ultimaLinha = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
            
            If ultimaLinha > 1 Then
                Set rngCopiar = wsOrigem.Range("A2:S" & ultimaLinha)
                
                ultimaLinha = wsConsolidado.Cells(wsConsolidado.Rows.Count, "A").End(xlUp).Row
                If ultimaLinha = 1 And wsConsolidado.Range("A1").Value <> "" Then
                    ultimaLinha = 1
                ElseIf wsConsolidado.Range("A1").Value = "" Then
                    ultimaLinha = 0
                End If

                rngCopiar.Copy wsConsolidado.Range("A" & ultimaLinha + 1)
            End If
        End If
    Next wsOrigem
    
    ultimaLinha = wsConsolidado.Cells(wsConsolidado.Rows.Count, "A").End(xlUp).Row
    If ultimaLinha >= 2 Then
        wsConsolidado.Range("S2:S" & ultimaLinha).Formula = "=IF(ISBLANK(M2),""Aguardando Previsão/Data"",NETWORKDAYS(L2,M2)-1)"
    End If
    
    If ultimaLinha >= 2 Then
        wsConsolidado.Range("O2:O" & ultimaLinha).Formula = "=IF(P2<>""EXPEDIÇÃO"",""Expedido""," & _
                      "IF(NETWORKDAYS(C2,TODAY())>=5,""Risco Alto""," & _
                      "IF(NETWORKDAYS(C2,TODAY())>=3,IF(NETWORKDAYS(C2,TODAY())<5,""Risco Médio"","""")," & _
                      "IF(NETWORKDAYS(C2,TODAY())<=2,""Risco Baixo"",""""))))"
    End If
    
    If ultimaLinha >= 1 Then
        wsConsolidado.Range("U1").Value = "$"
        If ultimaLinha >= 2 Then
            wsConsolidado.Range("U2:U" & ultimaLinha).Value = "$"
        End If
    End If
    
    If ultimaLinha >= 2 Then
        For Each cel In wsConsolidado.Range("P2:P" & ultimaLinha)
            If UCase(cel.Value) = "ENTREGUE" Or UCase(cel.Value) = "CANCELADA" Or UCase(cel.Value) = "RECUSADA" Then
                wsConsolidado.Cells(cel.Row, "T").Value = "*"
            End If
        Next cel
    End If
    
    On Error Resume Next
    Set wsBase = wbDestino.Worksheets("Base")
    On Error GoTo 0
    
    If wsBase Is Nothing Then
        MsgBox "A aba 'Base' não foi encontrada no workbook de destino.", vbExclamation
        GoTo FecharArquivo
    End If
    
    wsBase.Cells.Clear
    
    If wsConsolidado.Cells(1, 1).Value <> "" Then
        ultimaLinha = wsConsolidado.Cells(wsConsolidado.Rows.Count, "A").End(xlUp).Row
        wsConsolidado.Range("A1:U" & ultimaLinha).Copy
        wsBase.Range("A1").PasteSpecial xlPasteAll
        Application.CutCopyMode = False
    End If
    
    MsgBox "Processo concluído com sucesso!", vbInformation
    
FecharArquivo:
    wbOrigem.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
