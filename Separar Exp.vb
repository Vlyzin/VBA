Sub Button7_Click()
    'Separar Expedição
    Dim wsBase As Worksheet, wsTransportador As Worksheet, wsValidacao As Worksheet
    Dim novoArquivo As Workbook
    Dim caminho As String, nomeArquivo As String
    Dim dictCombinacoes As Object
    Dim dialogo As FileDialog

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set dialogo = Application.FileDialog(msoFileDialogFolderPicker)
    dialogo.Title = "Selecione a pasta para salvar os arquivos"

    If dialogo.Show = -1 Then
        caminho = dialogo.SelectedItems(1) & "\"
    Else
        MsgBox "Operação cancelada. Nenhum arquivo foi salvo.", vbExclamation
        Exit Sub
    End If

    Set wsBase = ThisWorkbook.Worksheets("Base")
    If wsBase.AutoFilterMode Then wsBase.AutoFilterMode = False

    Dim ultimaLinha As Long
    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, "P").End(xlUp).Row

    ' Aplicar filtro padrão: EXPEDIÇÃO e TRÂNSITO
    wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=16, Criteria1:=Array("EXPEDIÇÃO", "TRÂNSITO"), Operator:=xlFilterValues

    ' Criar dicionário com as combinações de transportador/centro com base no filtro
    Set dictCombinacoes = CreateObject("Scripting.Dictionary")
    Dim cel As Range, nomeTransportador As String, nomeCentro As String

    On Error Resume Next ' Em caso de não haver visíveis (evita erro)
    For Each cel In wsBase.Range("H2:H" & ultimaLinha).SpecialCells(xlCellTypeVisible)
        nomeTransportador = Trim(cel.Value)
        nomeCentro = Trim(wsBase.Cells(cel.Row, "E").Value)

        If nomeTransportador <> "" Then
            If UCase(nomeTransportador) = "BRAVO" Then
                If Not dictCombinacoes.Exists(nomeTransportador & "|") Then dictCombinacoes.Add nomeTransportador & "|", Array(nomeTransportador, "")
            ElseIf nomeCentro <> "" Then
                If Not dictCombinacoes.Exists(nomeTransportador & "|" & nomeCentro) Then dictCombinacoes.Add nomeTransportador & "|" & nomeCentro, Array(nomeTransportador, nomeCentro)
            End If
        End If
    Next cel
    On Error GoTo 0

    ' Incluir também as linhas ENTREGUE + FORA DO PRAZO
    Dim i As Long
    For i = 2 To ultimaLinha
        If Trim(wsBase.Cells(i, "P").Value) = "ENTREGUE" And Trim(wsBase.Cells(i, "N").Value) = "FORA DO PRAZO" Then
            nomeTransportador = Trim(wsBase.Cells(i, "H").Value)
            nomeCentro = Trim(wsBase.Cells(i, "E").Value)

            If nomeTransportador <> "" Then
                If UCase(nomeTransportador) = "BRAVO" Then
                    If Not dictCombinacoes.Exists(nomeTransportador & "|") Then dictCombinacoes.Add nomeTransportador & "|", Array(nomeTransportador, "")
                ElseIf nomeCentro <> "" Then
                    If Not dictCombinacoes.Exists(nomeTransportador & "|" & nomeCentro) Then dictCombinacoes.Add nomeTransportador & "|" & nomeCentro, Array(nomeTransportador, nomeCentro)
                End If
            End If
        End If
    Next i

    ' Para cada combinação, exportar os dados
    For Each combinacao In dictCombinacoes.Items()
        nomeTransportador = combinacao(0)
        nomeCentro = combinacao(1)

        ' Primeiro limpar filtros
        If wsBase.AutoFilterMode Then wsBase.AutoFilterMode = False

        ' Filtrar pela combinação
        wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=8, Criteria1:=nomeTransportador
        If UCase(nomeTransportador) <> "BRAVO" Then
            wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=5, Criteria1:=nomeCentro
        End If

        ' Filtrar status: EXPEDIÇÃO, TRÂNSITO, ou ENTREGUE + FORA DO PRAZO
        wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=16, Criteria1:=Array("EXPEDIÇÃO", "TRÂNSITO"), Operator:=xlFilterValues

        Dim rgExportar As Range
        On Error Resume Next
        Set rgExportar = wsBase.Range("A2:S" & ultimaLinha).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        ' Acrescentar ENTREGUE + FORA DO PRAZO manualmente
        For i = 2 To ultimaLinha
            If Trim(wsBase.Cells(i, "P").Value) = "ENTREGUE" And Trim(wsBase.Cells(i, "N").Value) = "FORA DO PRAZO" Then
                If Trim(wsBase.Cells(i, "H").Value) = nomeTransportador And Trim(wsBase.Cells(i, "E").Value) = nomeCentro Then
                    If rgExportar Is Nothing Then
                        Set rgExportar = wsBase.Range("A" & i & ":S" & i)
                    Else
                        Set rgExportar = Union(rgExportar, wsBase.Range("A" & i & ":S" & i))
                    End If
                End If
            End If
        Next i

        If Not rgExportar Is Nothing Then
            Set novoArquivo = Workbooks.Add
            Set wsValidacao = novoArquivo.Worksheets.Add(Before:=novoArquivo.Worksheets(1))
            wsValidacao.Name = "ListaValidacao"

            Dim listaValidacao As Variant
            listaValidacao = Array("Atraso na descarga vigente / anterior Bayer", "Atraso na descarga anterior outras empresas", "Atraso liberação veículo posto fiscal", "Condições climáticas / trajeto", "Agendado", "Aguardando agendamento", "Veículo chegou na data limite", "Lei do motorista", "Problemas mecânicos", "Atraso operacional", "Feriado municipal/estadual", "Alteração no local de entrega", "Cliente fechado", "NF recusada/ Cancelada", "NF sem transporte (AG)", "Frete dedicado", "FOB", "Atraso na escolta")
            For i = LBound(listaValidacao) To UBound(listaValidacao)
                wsValidacao.Cells(i + 1, 1).Value = listaValidacao(i)
            Next i

            Set wsTransportador = novoArquivo.Worksheets.Add(After:=novoArquivo.Worksheets(1))
            wsTransportador.Name = Left(nomeTransportador & IIf(nomeCentro <> "", " " & nomeCentro, ""), 31)

            wsBase.Range("A1:S1").Copy Destination:=wsTransportador.Range("A1")
            rgExportar.Copy Destination:=wsTransportador.Range("A2")

            With wsTransportador.Range("Q2:Q5000").Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ListaValidacao!$A$1:$A$18"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = "Selecione o motivo"
                .ErrorTitle = "Valor inválido"
                .InputMessage = "Por favor, selecione um motivo da lista dropdown"
                .ErrorMessage = "Você deve selecionar um item da lista pré-definida de motivos."
            End With

            wsValidacao.Visible = xlSheetVeryHidden
            wsTransportador.Columns.AutoFit

            nomeArquivo = nomeTransportador & IIf(nomeCentro <> "", " " & nomeCentro, "")
            nomeArquivo = Replace(Replace(Replace(nomeArquivo, "\", "-"), "/", "-"), ":", "-")

            novoArquivo.SaveAs caminho & nomeArquivo & ".xlsx", FileFormat:=xlOpenXMLWorkbook
            novoArquivo.Close False
        End If
    Next

    If wsBase.AutoFilterMode Then wsBase.AutoFilterMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Processo concluído! Foram exportados " & dictCombinacoes.Count & " arquivos em:" & vbNewLine & caminho, vbInformation
End Sub
