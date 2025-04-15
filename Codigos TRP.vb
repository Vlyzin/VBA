Sub Luft()
    Dim wsRetorno As Worksheet, wsBase As Worksheet
    Dim ultimaLinhaRetorno As Long, ultimaLinhaBase As Long
    Dim i As Long, j As Long
    Dim nfRetorno As String, nfBase As String
    Dim transportadora As String, statusBase As String
    Dim textoAV As String, novaData As Date, dataAtualM As Variant
    Dim encontrouNF As Boolean
    Dim textosEspecificos As Variant
    
    textosEspecificos = Array("EMISSAO DE MANIFESTO DE ENTREGA", "CHEGADA NA FILIAL", _
                             "EMISSAO DE MANIFESTO DE VIAGEM", "DEVOLUÇÃO TOTAL", _
                             "NF DEVOLVIDA A ORIGEM", "ENTREGA FINALIZADA", "EMISSAO DE CTE")
    
    Set wsRetorno = ThisWorkbook.Sheets("Retorno")
    Set wsBase = ThisWorkbook.Sheets("Base")
    
    ultimaLinhaRetorno = wsRetorno.Cells(wsRetorno.Rows.count, "A").End(xlUp).Row
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.count, "B").End(xlUp).Row
    
    For i = 2 To ultimaLinhaRetorno
        nfRetorno = CStr(wsRetorno.Cells(i, "A").Value)
        If nfRetorno <> "" Then
            encontrouNF = False
            
            For j = 2 To ultimaLinhaBase
                nfBase = CStr(wsBase.Cells(j, "B").Value)
                
                If nfBase = nfRetorno Then
                    transportadora = CStr(wsBase.Cells(j, "H").Value)
                    statusBase = CStr(wsBase.Cells(j, "P").Value)
                    
                    If transportadora = "LUFT" And (statusBase = "EXPEDIÇÃO" Or statusBase = "TRÂNSITO") Then
                        encontrouNF = True
                        textoAV = CStr(wsRetorno.Cells(i, "AV").Value)
                        
                        If Not IsInArray(textoAV, textosEspecificos) Then
                            textoAV = "OUTRO"
                        End If
                        
                        Select Case textoAV
                            Case "EMISSAO DE MANIFESTO DE ENTREGA", "CHEGADA NA FILIAL", "EMISSAO DE MANIFESTO DE VIAGEM"

                                wsBase.Cells(j, "P").Value = "TRÂNSITO"
                                
                                If IsDate(wsRetorno.Cells(i, "AI").Value) Then
                                    novaData = CDate(wsRetorno.Cells(i, "AI").Value)
                                    dataAtualM = wsBase.Cells(j, "M").Value
                                    
                                    If IsEmpty(dataAtualM) Or dataAtualM = "" Then
                                        wsBase.Cells(j, "M").Value = novaData
                                    ElseIf IsDate(dataAtualM) And novaData > CDate(dataAtualM) Then
                                        wsBase.Cells(j, "M").Value = novaData
                                    End If
                                End If
                                
                            Case "DEVOLUÇÃO TOTAL", "NF DEVOLVIDA A ORIGEM"

                                wsBase.Cells(j, "P").Value = "RECUSADA"
                                wsBase.Cells(j, "M").ClearContents
                                
                            Case "ENTREGA FINALIZEDA"

                                wsBase.Cells(j, "P").Value = "ENTREGUE"
                                If IsDate(wsRetorno.Cells(i, "AF").Value) Then
                                    wsBase.Cells(j, "M").Value = CDate(wsRetorno.Cells(i, "AF").Value)
                                End If
                                
                            Case "EMISSAO DE CTE", "", "OUTRO"

                                If IsDate(wsRetorno.Cells(i, "AI").Value) Then
                                    novaData = CDate(wsRetorno.Cells(i, "AI").Value)
                                    dataAtualM = wsBase.Cells(j, "M").Value
                                    
                                    If IsEmpty(dataAtualM) Or dataAtualM = "" Then
                                        wsBase.Cells(j, "M").Value = novaData
                                    ElseIf IsDate(dataAtualM) And novaData > CDate(dataAtualM) Then
                                        wsBase.Cells(j, "M").Value = novaData
                                    End If
                                End If
                        End Select
                    End If
                End If
                
                If encontrouNF Then Exit For
            Next j
        End If
    Next i
    
    MsgBox "Processamento LUFT concluído!", vbInformation
End Sub

Function IsInArray(valor As String, arr As Variant) As Boolean
    Dim elemento As Variant
    For Each elemento In arr
        If CStr(elemento) = valor Then
            IsInArray = True
            Exit Function
        End If
    Next elemento
    IsInArray = False
End Function

Sub Bravo()
    Dim wsRetorno As Worksheet, wsBase As Worksheet
    Dim ultimaLinhaRetorno As Long, ultimaLinhaBase As Long
    Dim i As Long, j As Long
    Dim nfRetorno As String, nfBase As String
    Dim transportadora As String, statusBase As String, codigoFilial As String
    Dim filialRetorno As String
    Dim dataF As Variant, dataB As Variant, dataC As Variant
    Dim encontrouNF As Boolean
    Dim filialValida As Boolean
    
    ' Definir as planilhas
    Set wsRetorno = ThisWorkbook.Sheets("Retorno")
    Set wsBase = ThisWorkbook.Sheets("Base")
    
    ' Encontrar última linha em ambas as planilhas
    ultimaLinhaRetorno = wsRetorno.Cells(wsRetorno.Rows.count, "A").End(xlUp).Row
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.count, "B").End(xlUp).Row
    
    ' Loop através da planilha Retorno
    For i = 2 To ultimaLinhaRetorno
        nfRetorno = CStr(wsRetorno.Cells(i, "A").Value)
        If nfRetorno <> "" Then
            encontrouNF = False
            
            ' Buscar NF na planilha Base
            For j = 2 To ultimaLinhaBase
                nfBase = CStr(wsBase.Cells(j, "B").Value)
                
                ' Verificar se a NF corresponde
                If nfBase = nfRetorno Then
                    transportadora = CStr(wsBase.Cells(j, "H").Value)
                    statusBase = CStr(wsBase.Cells(j, "P").Value)
                    codigoFilial = CStr(wsBase.Cells(j, "D").Value)
                    filialRetorno = CStr(wsRetorno.Cells(i, "I").Value)
                    
                    ' Verificar transportadora, status e filial
                    filialValida = False
                    If transportadora = "BRAVO" And (statusBase = "EXPEDIÇÃO" Or statusBase = "TRÂNSITO") Then
                        ' Verificar combinações de filial
                        Select Case filialRetorno
                            Case "BRAVO LEM"
                                If codigoFilial = "MWA4" Then filialValida = True
                            Case "BRAVO CUIABA"
                                If codigoFilial = "MWA0" Then filialValida = True
                            Case "BRAVO QUERENCIA"
                                If codigoFilial = "MWB0" Then filialValida = True
                            Case "BRAVO UBERABA"
                                If codigoFilial = "MWA1" Then filialValida = True
                        End Select
                        
                        If filialValida Then
                            encontrouNF = True
                            
                            ' Primeira validação - Coluna F (ENTREGUE)
                            dataF = wsRetorno.Cells(i, "F").Value
                            If Not IsEmpty(dataF) And dataF <> "n.i." And IsDate(dataF) Then
                                wsBase.Cells(j, "M").Value = CDate(dataF)
                                wsBase.Cells(j, "P").Value = "ENTREGUE"
                                
                            ' Segunda validação - Coluna B (TRÂNSITO)
                            Else
                                dataB = wsRetorno.Cells(i, "B").Value
                                If Not IsEmpty(dataB) And IsDate(dataB) Then
                                    wsBase.Cells(j, "M").Value = CDate(dataB)
                                    wsBase.Cells(j, "P").Value = "TRÂNSITO"
                                    
                                ' Terceira validação - Coluna C (apenas atualiza data)
                                Else
                                    dataC = wsRetorno.Cells(i, "C").Value
                                    If Not IsEmpty(dataC) And IsDate(dataC) Then
                                        wsBase.Cells(j, "M").Value = CDate(dataC)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                If encontrouNF Then Exit For
            Next j
        End If
    Next i
    
    MsgBox "Processamento BRAVO concluído!", vbInformation
End Sub

Sub Toniato()
    Dim wsRetorno As Worksheet, wsBase As Worksheet
    Dim ultimaLinhaRetorno As Long, ultimaLinhaBase As Long
    Dim i As Long, j As Long
    Dim nfRetorno As String, nfBase As String
    Dim transportadora As String, statusBase As String, codigoFilial As String
    Dim filialRetorno As String, textoV As String
    Dim novaData As Date, dataAtualM As Variant
    Dim encontrouNF As Boolean
    Dim filialValida As Boolean
    
    Set wsRetorno = ThisWorkbook.Sheets("Retorno")
    Set wsBase = ThisWorkbook.Sheets("Base")
    
    ultimaLinhaRetorno = wsRetorno.Cells(wsRetorno.Rows.count, "G").End(xlUp).Row
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.count, "B").End(xlUp).Row
    
    For i = 2 To ultimaLinhaRetorno
        nfRetorno = CStr(wsRetorno.Cells(i, "G").Value)
        If nfRetorno <> "" Then
            encontrouNF = False
            
            For j = 2 To ultimaLinhaBase
                nfBase = CStr(wsBase.Cells(j, "B").Value)
                
                If nfBase = nfRetorno Then
                    transportadora = CStr(wsBase.Cells(j, "H").Value)
                    statusBase = CStr(wsBase.Cells(j, "P").Value)
                    codigoFilial = CStr(wsBase.Cells(j, "D").Value)
                    filialRetorno = CStr(wsRetorno.Cells(i, "M").Value)
                    textoV = CStr(wsRetorno.Cells(i, "V").Value)
                    
                    filialValida = False
                    If transportadora = "TONIATO" And (statusBase = "EXPEDIÇÃO" Or statusBase = "TRÂNSITO") Then
                        Select Case filialRetorno
                            Case "RIO VERDE"
                                If codigoFilial = "MWC0" Then filialValida = True
                            Case "PAULINIA"
                                If codigoFilial = "MWA5" Then filialValida = True
                            Case "BELFORD ROXO"
                                If codigoFilial = "MWA6" Then filialValida = True
                            Case "IBIPORA"
                                If codigoFilial = "MWA3" Then filialValida = True
                        End Select
                        
                        If filialValida Then
                            encontrouNF = True
                            
                            Select Case textoV
                                Case "Transporte em Trânsito"
                                    wsBase.Cells(j, "P").Value = "TRÂNSITO"
                                    If IsDate(wsRetorno.Cells(i, "R").Value) Then
                                        novaData = CDate(wsRetorno.Cells(i, "R").Value)
                                        dataAtualM = wsBase.Cells(j, "M").Value
                                        
                                        If IsEmpty(dataAtualM) Or dataAtualM = "" Then
                                            wsBase.Cells(j, "M").Value = novaData
                                        ElseIf IsDate(dataAtualM) And novaData > CDate(dataAtualM) Then
                                            wsBase.Cells(j, "M").Value = novaData
                                        End If
                                    End If
                                    
                                Case "Entrega Realizada Normalmente"
                                    wsBase.Cells(j, "P").Value = "ENTREGUE"
                                    If IsDate(wsRetorno.Cells(i, "S").Value) Then
                                        wsBase.Cells(j, "M").Value = CDate(wsRetorno.Cells(i, "S").Value)
                                    End If
                                    
                                Case "Mercadoria Devolvida ao Cliente de Origem"
                                    wsBase.Cells(j, "P").Value = "RECUSADA"
                                    wsBase.Cells(j, "N").ClearContents
                                    
                                Case Else
                                    If IsDate(wsRetorno.Cells(i, "R").Value) Then
                                        novaData = CDate(wsRetorno.Cells(i, "R").Value)
                                        dataAtualM = wsBase.Cells(j, "M").Value
                                        
                                        If IsEmpty(dataAtualM) Or dataAtualM = "" Then
                                            wsBase.Cells(j, "M").Value = novaData
                                        ElseIf IsDate(dataAtualM) And novaData > CDate(dataAtualM) Then
                                            wsBase.Cells(j, "M").Value = novaData
                                        End If
                                    End If
                            End Select
                        End If
                    End If
                End If
                
                If encontrouNF Then Exit For
            Next j
        End If
    Next i
    
    MsgBox "Processamento TONIATO concluído!", vbInformation
End Sub
