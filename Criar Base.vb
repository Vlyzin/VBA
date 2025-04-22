Sub Button9_Click()
    'Criar Base
    Dim wsSAP As Worksheet, wsBase As Worksheet, wsItinerario As Worksheet
    Dim lastRowSAP As Long, lastRowBase As Long
    Dim rngToCopy As Range
    Dim remessaCol As Long, notaFiscalCol As Long
    Dim cell As Range
    Dim i As Long
    Dim dataEmissao As Date
    Dim prazoTransportadora As Integer
    Dim centro As String
    Dim transportador As String
    Dim cidadeDestino As String
    Dim estadoDestino As String
    Dim dataSLA As Date
    Dim rng As Range
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Set wsSAP = ThisWorkbook.Worksheets("SAP")
    Set wsBase = ThisWorkbook.Worksheets("Base")
    Set wsItinerario = ThisWorkbook.Worksheets("Itinerario")
    On Error GoTo 0
    
    If wsSAP Is Nothing Then
        MsgBox "Planilha 'SAP' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    If wsBase Is Nothing Then
        MsgBox "Planilha 'Base' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    If wsItinerario Is Nothing Then
        MsgBox "Planilha 'Itinerario' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    With wsSAP
        .Columns("H").Delete Shift:=xlToLeft
        .Columns("I").Delete Shift:=xlToLeft
        
        .Columns("C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        
        lastRowSAP = .Cells(.Rows.Count, "B").End(xlUp).Row
        .Range("B2:B" & lastRowSAP).TextToColumns _
            Destination:=.Range("B2"), _
            DataType:=xlDelimited, _
            Other:=True, _
            OtherChar:="-"
        
        .Columns("C").Delete Shift:=xlToLeft
        
        lastRowSAP = .Cells(.Rows.Count, "E").End(xlUp).Row
        With .Range("E1:E" & lastRowSAP)
            .Replace "DC Uberaba CP Cerrados", "DC Uberaba CP", xlPart
            .Replace "DC Rio Verde Loja CP", "DC Rio Verde CP", xlPart
        End With
        
        For i = 1 To lastRowSAP
            If .Cells(i, "H").Value = "" Then
                Select Case .Cells(i, "E").Value
                    Case "DC Carazinho CP"
                        .Cells(i, "H").Value = "LUFT"
                    Case "DC Cuiaba CP", "DC Uberaba CP", "DC LEM CP", "DC Querencia CP"
                        .Cells(i, "H").Value = "BRAVO"
                    Case "DC Paulinia CP", "DC Rio Verde CP", "DC Ibipora CP", "WH Belford Roxo CP"
                        .Cells(i, "H").Value = "TONIATO"
                End Select
            End If
        Next i
        
        With .Range("H1:H" & lastRowSAP)
            .Replace "BRAVO SERVICOS LOGISTICOS LTDA", "BRAVO", xlPart
            .Replace "GT SOLUCOES LOGISTICAS SA", "TONIATO", xlPart
            .Replace "TRANSPORTES LUFT LTDA", "LUFT", xlPart
        End With
        
        Set rngToCopy = .Range("A2").CurrentRegion.Offset(1, 0).Resize(.Range("A2").CurrentRegion.Rows.Count - 1)
    End With
    
    With wsBase
        lastRowBase = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        If lastRowBase = 0 Then
            wsSAP.Range("A1").CurrentRegion.Rows(1).Copy Destination:=.Range("A1")
            lastRowBase = 1
        ElseIf lastRowBase = 1 And .Range("A1").Value = "" Then
            lastRowBase = 0
        End If
        
        rngToCopy.Copy Destination:=.Cells(lastRowBase + 1, "A")
        
        lastRowBase = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        remessaCol = 0
        notaFiscalCol = 0
        
        On Error Resume Next
        remessaCol = Application.Match("Remessa", .Rows(1), 0)
        notaFiscalCol = Application.Match("Nota fiscal", .Rows(1), 0)
        On Error GoTo 0
        
        If remessaCol = 0 Or notaFiscalCol = 0 Then
            MsgBox "Colunas 'Remessa' e/ou 'Nota fiscal' não encontradas!", vbExclamation
        Else
            If lastRowBase > 1 Then
                .Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(remessaCol, notaFiscalCol), Header:=xlYes
            End If
        End If
        
        lastRowBase = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        If lastRowBase > 1 Then
            If .Range("O2").Formula <> "" Then
                .Range("O2").AutoFill Destination:=.Range("O2:O" & lastRowBase)
            End If
            
            If .Range("S2").Formula <> "" Then
                .Range("S2").AutoFill Destination:=.Range("S2:S" & lastRowBase)
            End If
            
            For Each cell In .Range("P2:P" & lastRowBase)
                If cell.Value = "" Then
                    cell.Value = "EXPEDIÇÃO"
                End If
            Next cell
        End If
        
' 2. CÁLCULO DE PRAZOS
If lastRowBase > 1 Then
    For i = 2 To lastRowBase
        If .Cells(i, 21).Value = "$" Then
            GoTo ProximaLinha
        End If
        
        dataEmissao = .Cells(i, 3).Value
        centro = .Cells(i, 4).Value
        transportador = .Cells(i, 8).Value
        cidadeDestino = .Cells(i, 10).Value
        estadoDestino = .Cells(i, 11).Value 
        
        prazoTransportadora = 0
        
        Set rng = wsItinerario.Range("A2:A" & wsItinerario.Cells(wsItinerario.Rows.Count, "A").End(xlUp).Row)
        For Each cell In rng
            If cell.Value = centro And cell.Offset(0, 1).Value = transportador And _
               cell.Offset(0, 3).Value = cidadeDestino And cell.Offset(0, 4).Value = estadoDestino Then
                prazoTransportadora = cell.Offset(0, 5).Value
                Exit For
            End If
        Next cell
        
        If transportador = "LUFT" And centro = "MWA5" Then
            dataSLA = AplicarRegraEspecial(dataEmissao)
        Else
            dataSLA = AdicionarDiasUteis(dataEmissao, 2)
        End If
        
        dataSLA = AdicionarDiasUteis(dataSLA, prazoTransportadora)

        .Cells(i, 12).Value = dataSLA
        
ProximaLinha:
    Next i
End If
    End With
    
    Application.ScreenUpdating = True
    
    With wsSAP
        lastRowSAP = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRowSAP > 1 Then
            .Range("A1:Z" & lastRowSAP).ClearContents
        End If
    End With
    
    With wsBase
        lastRowBase = .Cells(.Rows.Count, "A").End(xlUp).Row
        If lastRowBase > 1 Then
            .Range("U2:U" & lastRowBase).Value = "$"
        End If
    End With
    
    Call Macro_Botão6_Clique
    
    wsBase.UsedRange.Columns.AutoFit
        With wsBase.UsedRange
        .Columns.AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    MsgBox "Feito!"
End Sub

Function AplicarRegraEspecial(dataEmissao As Date) As Date
    Dim diaSemana As Integer
    Dim quartaFeiraMesmaSemana As Date
    Dim sextaFeiraSeguinte As Date
    
    diaSemana = Weekday(dataEmissao, vbMonday) ' 1 = Segunda, 2 = Terça, ..., 5 = Sexta
    
    If diaSemana >= 5 Or diaSemana <= 2 Then  ' Se a NF foi emitida entre sexta (5) e terça (2)
        quartaFeiraMesmaSemana = dataEmissao + (3 - diaSemana + 7) Mod 7
        AplicarRegraEspecial = quartaFeiraMesmaSemana
    ElseIf diaSemana = 3 Or diaSemana = 4 Then ' Quarta (3) ou Quinta (4)
        sextaFeiraSeguinte = dataEmissao + (5 - diaSemana)
        AplicarRegraEspecial = sextaFeiraSeguinte
    Else
        AplicarRegraEspecial = dataEmissao
    End If
End Function

Function AdicionarDiasUteis(dataInicial As Date, dias As Integer) As Date
    Dim i As Integer
    Dim dataTemp As Date
    Dim feriados As Variant
    Dim feriado As Date
    
    feriados = Array(DateSerial(Year(dataInicial), 12, 25), DateSerial(Year(dataInicial), 1, 1), DateSerial(Year(dataInicial), 3, 3), DateSerial(Year(dataInicial), 3, 4), DateSerial(Year(dataInicial), 4, 18), DateSerial(Year(dataInicial), 4, 21), DateSerial(Year(dataInicial), 5, 1), DateSerial(Year(dataInicial), 6, 19), DateSerial(Year(dataInicial), 9, 7), DateSerial(Year(dataInicial), 10, 12), DateSerial(Year(dataInicial), 11, 2), DateSerial(Year(dataInicial), 11, 15))
    
    dataTemp = dataInicial
    
    Do While dias > 0
        dataTemp = dataTemp + 1
        
        If Weekday(dataTemp, vbMonday) <= 5 And Not EhFeriado(dataTemp, feriados) Then
            dias = dias - 1
        End If
    Loop
    
    AdicionarDiasUteis = dataTemp
End Function

Function EhFeriado(data As Date, feriados As Variant) As Boolean
    Dim i As Integer
    EhFeriado = False
    For i = LBound(feriados) To UBound(feriados)
        If data = feriados(i) Then
            EhFeriado = True
            Exit Function
        End If
    Next i
End Function
