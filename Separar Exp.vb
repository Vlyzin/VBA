Sub Button7_Click()
    Dim wsBase As Worksheet, wsTransportador As Worksheet, wsValidacao As Worksheet
    Dim novoArquivo As Workbook
    Dim caminho As String, nomeArquivo As String, dataAtual As String
    Dim dictCombinacoes As Object
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    caminho = "C:\Users\vinicius.domingues\Documents\Projeto\Base Bayer"
    dataAtual = Format(Date, "dd.mm")
    If Right(caminho, 1) <> "\" Then caminho = caminho & "\"
    
    Set wsBase = ThisWorkbook.Worksheets("Base")
    If wsBase.AutoFilterMode Then wsBase.AutoFilterMode = False
    
    Dim ultimaLinha As Long
    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, "P").End(xlUp).Row
    wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=16, Criteria1:=Array("EXPEDIÇÃO", "TRÂNSITO"), Operator:=xlFilterValues
    
    Set dictCombinacoes = CreateObject("Scripting.Dictionary")
    
    Dim cel As Range, nomeTransportador As String, nomeCentro As String
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
    
    For Each combinacao In dictCombinacoes.Items()
        nomeTransportador = combinacao(0)
        nomeCentro = combinacao(1)
        
        wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=8, Criteria1:=nomeTransportador
        If UCase(nomeTransportador) <> "BRAVO" Then wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=5, Criteria1:=nomeCentro
        
        If wsBase.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count > 1 Then
            Set novoArquivo = Workbooks.Add
            
            Set wsValidacao = novoArquivo.Worksheets.Add(Before:=novoArquivo.Worksheets(1))
            wsValidacao.Name = "ListaValidacao"
            Dim listaValidacao As Variant, i As Long
            listaValidacao = Array("Atraso na descarga vigente / anterior Bayer", "Atraso na descarga anterior outras empresas", "Atraso liberação veículo posto fiscal", "Condições climáticas / trajeto", "Agendado", "Aguardando agendamento", "Veículo chegou na data limite", "Lei do motorista", "Problemas mecânicos", "Atraso operacional", "Feriado municipal/estadual", "Alteração no local de entrega", "Cliente fechado", "NF recusada/ Cancelada", "NF sem transporte (AG)", "Frete dedicado", "FOB", "Atraso na escolta")
            For i = LBound(listaValidacao) To UBound(listaValidacao)
                wsValidacao.Cells(i + 1, 1).Value = listaValidacao(i)
            Next i

            Set wsTransportador = novoArquivo.Worksheets.Add(After:=novoArquivo.Worksheets(1))
            wsTransportador.Name = Left(nomeTransportador & IIf(nomeCentro <> "", " " & nomeCentro, ""), 31)
            
            wsBase.Range("A1:S1").Copy Destination:=wsTransportador.Range("A1")
            wsBase.Range("A2:S" & ultimaLinha).SpecialCells(xlCellTypeVisible).Copy Destination:=wsTransportador.Range("A2")
            
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
            
            nomeArquivo = nomeTransportador & IIf(nomeCentro <> "", " " & nomeCentro, "") & " " & dataAtual
            nomeArquivo = Replace(Replace(Replace(nomeArquivo, "\", "-"), "/", "-"), ":", "-")
            
            novoArquivo.SaveAs caminho & nomeArquivo & ".xlsx", FileFormat:=xlOpenXMLWorkbook
            novoArquivo.Close False
        End If
        
        If UCase(nomeTransportador) <> "BRAVO" Then wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=5
    Next
    
    If wsBase.AutoFilterMode Then wsBase.AutoFilterMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Processo concluído! Foram exportados " & dictCombinacoes.Count & " arquivos.", vbInformation
End Sub