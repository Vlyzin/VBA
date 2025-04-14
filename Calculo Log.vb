Sub Macro_Botão6_Clique()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim diasAtraso As Long
    Dim status As String
    Dim comentario As String
    Dim dataPrevisao As Date
    Dim dataEmissao As Date
    Dim chamado As String
    Dim dataSLA As Date
    
    Set ws = ThisWorkbook.Sheets("Base")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ws.Columns("N").FormatConditions.Delete
    
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, "S").Value) Then
            diasAtraso = CLng(ws.Cells(i, "S").Value)
        Else
            diasAtraso = 0
        End If
        
        status = ws.Cells(i, "P").Value
        comentario = ws.Cells(i, "Q").Value
        dataPrevisao = ws.Cells(i, "M").Value
        dataEmissao = ws.Cells(i, "C").Value
        dataSLA = ws.Cells(i, "L").Value
        chamado = ws.Cells(i, "R").Value
        
        If InStr(ws.Cells(i, "T").Value, "*") = 0 Then
            If diasAtraso >= 30 And ws.Cells(i, "N").Value = "Avaliar" Then
                ws.Cells(i, "N").Value = "AVALIAR"
            ElseIf status = "ENTREGUE" And dataPrevisao > Date Then
                ws.Cells(i, "N").Value = "AVALIAR"
            ElseIf dataPrevisao < dataEmissao Then
                ws.Cells(i, "N").Value = "AVALIAR"
            ElseIf (status = "TRÂNSITO" Or status = "EXPEDIÇÃO") And dataPrevisao < Date Then
                ws.Cells(i, "N").Value = "AVALIAR"
            ElseIf diasAtraso = 0 Then
                ws.Cells(i, "N").Value = "NO PRAZO"
            ElseIf dataPrevisao < dataSLA Then
                ws.Cells(i, "N").Value = "NO PRAZO"
            ElseIf comentario = "Frete dedicado" Or comentario = "NF sem transporte (AG)" Or comentario = "FOB" Or (comentario = "Agendado" And diasAtraso <= 2) Then
                ws.Cells(i, "N").Value = "NO PRAZO"
            ElseIf comentario = "Agendado" And diasAtraso > 2 Then
                ws.Cells(i, "N").Value = "AVALIAR"
            ElseIf comentario = "Atraso na descarga vigente / anterior Bayer" And diasAtraso <= 1 And chamado <> "" Then
                ws.Cells(i, "N").Value = "NO PRAZO"
            ElseIf comentario = "Atraso na descarga vigente / anterior Bayer" And diasAtraso > 1 And chamado <> "" Then
                ws.Cells(i, "N").Value = "AVALIAR"
            ElseIf comentario = "Atraso na liberação veículo posto fiscal" Then
                ws.Cells(i, "N").Value = "AVALIAR"
            ElseIf comentario = "Condições climáticas / trajeto" Then
                ws.Cells(i, "N").Value = "AVALIAR"
            ElseIf comentario = "Aguardando agendamento" Then
                ws.Cells(i, "N").Value = "AVALIAR"
            ElseIf comentario = "Veículo chegou na data limite" And diasAtraso <= 1 And chamado <> "" Then
                ws.Cells(i, "N").Value = "NO PRAZO"
            ElseIf comentario = "Feraido municipal/estadual" And diasAtraso <= 1 Then
                ws.Cells(i, "N").Value = "NO PRAZO"
            ElseIf comentario = "Alteração no local de entrega" And diasAtraso <= 1 And chamado <> "" Then
                ws.Cells(i, "N").Value = "NO PRAZO"
            ElseIf comentario = "Cliente fechado" And diasAtraso <= 1 And chamado <> "" Then
                ws.Cells(i, "N").Value = "NO PRAZO"
            ElseIf comentario = "Cliente fechado" And diasAtraso > 1 And chamado <> "" Then
                ws.Cells(i, "N").Value = "AVALIAR"
            ElseIf comentario = "Atraso na escolta" And diasAtraso <= 2 And chamado <> "" Then
                ws.Cells(i, "N").Value = "NO PRAZO"
            Else
                ws.Cells(i, "N").Value = "FORA DO PRAZO"
            End If
        End If
    Next i
    
    With ws.Range("N2:N" & lastRow)
        
        .FormatConditions.Add Type:=xlTextString, String:="NO PRAZO", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.count).Interior.Color = RGB(198, 239, 206)
        
        .FormatConditions.Add Type:=xlTextString, String:="AVALIAR", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.count).Interior.Color = RGB(255, 235, 156)
        
        .FormatConditions.Add Type:=xlTextString, String:="FORA DO PRAZO", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.count).Interior.Color = RGB(255, 199, 206)
    End With
    
    MsgBox "Cálculo concluído e formatação aplicada na coluna N!"
End Sub
