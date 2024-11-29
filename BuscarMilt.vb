Sub Macro_Botão7_Clique()
    Dim wsBase As Worksheet
    Dim wsBasebay As Worksheet
    Dim ultimaLinhaBase As Long
    Dim ultimaLinhaBasebay As Long
    Dim i As Long
    Dim remessa As String
    Dim resultado As String
    Dim j As Long
    Dim remessaEncontrada As Range
    
    
    Set wsBase = ThisWorkbook.Sheets("Base")
    Set wsBasebay = ThisWorkbook.Sheets("Basebay")
    
    
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaBasebay = wsBasebay.Cells(wsBasebay.Rows.Count, "A").End(xlUp).Row
    
    
    For i = 2 To ultimaLinhaBase
        remessa = wsBase.Cells(i, 1).Value
        resultado = "Remessa não localizada"
        
        
        Set remessaEncontrada = wsBasebay.Columns(1).Find(remessa, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not remessaEncontrada Is Nothing Then
            
            For j = remessaEncontrada.Row To ultimaLinhaBasebay
                If wsBasebay.Cells(j, 14).Value = "87335070" Or wsBasebay.Cells(j, 14).Value = "88565177" Then
                    resultado = "CONTÉM"
                    Exit For
                End If
            Next j
            
            
            If resultado = "Remessa não localizada" Then
                resultado = "NÃO CONTÉM"
            End If
        End If
        
        
        wsBase.Cells(i, 20).Value = resultado
    Next i
    
    MsgBox "Processo concluído!", vbInformation
End Sub