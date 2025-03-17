Sub Botão1_Clique()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim dataEmissao As Date
    Dim prazoTransportadora As Integer
    Dim centro As String
    Dim transportador As String
    Dim cidadeDestino As String
    Dim estadoDestino As String
    Dim dataSLA As Date
    Dim rng As Range
    Dim cell As Range
    
    Set ws1 = ThisWorkbook.Sheets("Planilha1")
    Set ws2 = ThisWorkbook.Sheets("Itinerario")
    
    ultimaLinha = ws1.Cells(ws1.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To ultimaLinha
        dataEmissao = ws1.Cells(i, 2).Value
        centro = ws1.Cells(i, 3).Value
        transportador = ws1.Cells(i, 4).Value
        cidadeDestino = ws1.Cells(i, 6).Value
        estadoDestino = ws1.Cells(i, 7).Value
        
        prazoTransportadora = 0
        
        Set rng = ws2.Range("A2:A" & ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row)
        For Each cell In rng
            If cell.Value = centro And cell.Offset(0, 1).Value = transportador And cell.Offset(0, 3).Value = cidadeDestino And cell.Offset(0, 4).Value = estadoDestino Then
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
        
        ws1.Cells(i, 5).Value = dataSLA
    Next i
    
    MsgBox ("Feito!")
End Sub

Function AplicarRegraEspecial(dataEmissao As Date) As Date
    Dim diaSemana As Integer
    Dim quartaFeiraMesmaSemana As Date
    Dim sextaFeiraSeguinte As Date
    
    diaSemana = Weekday(dataEmissao, vbMonday) ' 1 = Segunda, 2 = Terça, ..., 5 = Sexta
    
  
    If diaSemana >= 5 Or diaSemana <= 2 Then  ' Se a NF foi emitida entre sexta (5) e terça (2)

        quartaFeiraMesmaSemana = dataEmissao + (3 - diaSemana + 7) Mod 7
        AplicarRegraEspecial = quartaFeiraMesmaSemana + 1

    ElseIf diaSemana = 3 Or diaSemana = 4 Then ' Quarta (3) ou Quinta (4)

        sextaFeiraSeguinte = dataEmissao + (5 - diaSemana)
        AplicarRegraEspecial = AdicionarDiasUteis(sextaFeiraSeguinte, 1)
    Else
        AplicarRegraEspecial = dataEmissao
    End If
End Function

Function AdicionarDiasUteis(dataInicial As Date, dias As Integer) As Date
    Dim i As Integer
    Dim dataTemp As Date
    Dim feriados As Variant
    Dim feriado As Date
    
    feriados = Array(DateSerial(Year(dataInicial), 12, 25), DateSerial(Year(dataInicial), 1, 1), DateSerial(Year(dataInicial), 3, 3), DateSerial(Year(dataInicial), 3, 4))
    
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
