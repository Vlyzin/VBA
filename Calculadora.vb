Sub Botão1_Clique()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim dataEmissao As Date
    Dim prazoTransportadora As Integer
    Dim centro As String
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
        cidadeDestino = ws1.Cells(i, 5).Value
        estadoDestino = ws1.Cells(i, 6).Value
        prazoTransportadora = 0
        Set rng = ws2.Range("A:A").Find(centro, LookIn:=xlValues, LookAt:=xlWhole)
        If Not rng Is Nothing Then
            For Each cell In ws2.Range(rng, ws2.Cells(ws2.Rows.Count, "A").End(xlUp))
                If cell.Value = centro And cell.Offset(0, 3).Value = cidadeDestino And cell.Offset(0, 4).Value = estadoDestino Then
                    prazoTransportadora = cell.Offset(0, 5).Value
                    Exit For
                End If
            Next cell
        End If
        dataSLA = AdicionarDiasUteis(dataEmissao, 2)
        dataSLA = AdicionarDiasUteis(dataSLA, prazoTransportadora)
        ws1.Cells(i, 4).Value = dataSLA
    Next i
    MsgBox ("Feito!")
End Sub
 
Function AdicionarDiasUteis(dataInicial As Date, dias As Integer) As Date
    Dim i As Integer
    Dim dataTemp As Date
    Dim feriados As Variant
    Dim feriado As Date
    ' Definir os feriados
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
