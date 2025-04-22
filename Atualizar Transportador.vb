Sub Macro_Boa()
    'Atualizar via Retorno
    Application.DisplayAlerts = False
    Dim wsBase As Worksheet
    Dim wsRetorno As Worksheet
    Dim wsMacro As Worksheet
    Dim ultimaLinhaBase As Long
    Dim ultimaLinhaRetorno As Long
    Dim i As Long
    Dim j As Long
    Dim numeroRemessa As String
    Dim dataAtualizada As Boolean
    Dim ultimaRemessa As String
    Dim dataRetorno As Variant, dataBase As Variant

    Set wsBase = ThisWorkbook.Sheets("Base")
    Set wsRetorno = ThisWorkbook.Sheets("Retorno")
    Set wsMacro = ThisWorkbook.Sheets("Macro")
    
    ultimaLinhaBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaRetorno = wsRetorno.Cells(wsRetorno.Rows.Count, "A").End(xlUp).Row
    
    ultimaRemessa = wsRetorno.Cells(ultimaLinhaRetorno, 1).Value
    dataRetorno = wsRetorno.Cells(ultimaLinhaRetorno, 13).Value
    
    For i = 2 To ultimaLinhaRetorno
        numeroRemessa = wsRetorno.Cells(i, 1).Value

        For j = 2 To ultimaLinhaBase
            If wsBase.Cells(j, 1).Value = numeroRemessa Then
                If i = ultimaLinhaRetorno Then
                    dataBase = wsBase.Cells(j, 13).Value
                End If
                
                wsBase.Cells(j, 13).Value = wsRetorno.Cells(i, 13).Value
                wsBase.Cells(j, 16).Value = wsRetorno.Cells(i, 16).Value
                wsBase.Cells(j, 17).Value = wsRetorno.Cells(i, 17).Value
                wsBase.Cells(j, 18).Value = wsRetorno.Cells(i, 18).Value
                Exit For
            End If
        Next j
    Next i
    
    If ultimaLinhaRetorno > 1 Then
        For j = 2 To ultimaLinhaBase
            If wsBase.Cells(j, 1).Value = ultimaRemessa Then
                If wsBase.Cells(j, 13).Value <> dataRetorno Then
                    MsgBox "ATENÇÃO: A remessa " & ultimaRemessa & " não foi atualizada corretamente." & vbCrLf & _
                           "Data no Retorno: " & dataRetorno & vbCrLf & _
                           "Data na Base: " & wsBase.Cells(j, 13).Value & vbCrLf & vbCrLf & _
                           "Por favor, verifique manualmente.", vbExclamation, "Verificação de Dados"
                    dataAtualizada = False
                    Exit For
                Else
                    dataAtualizada = True
                End If
            End If
        Next j
    End If

    Dim novoArquivo As Workbook
    Dim caminhoSalvar As String
    Dim dialogo As FileDialog
    
    Set dialogo = Application.FileDialog(msoFileDialogFolderPicker)
    dialogo.Title = "Selecione a pasta para salvar o arquivo"
    
    If dialogo.Show = -1 Then
        caminhoSalvar = dialogo.SelectedItems(1) & "\"
    Else
        MsgBox "Operação cancelada. O arquivo não foi salvo.", vbExclamation
        Exit Sub
    End If
    
    Set novoArquivo = Workbooks.Add
    wsBase.Copy Before:=novoArquivo.Sheets(1)
    novoArquivo.SaveAs caminhoSalvar & "CP Report Fixo Consolidado.xlsx"
    novoArquivo.Close False
    
    If dataAtualizada Or ultimaLinhaRetorno <= 1 Then
        MsgBox "Atualização concluída e arquivo salvo em:" & vbCrLf & caminhoSalvar, vbInformation
    End If
    
    Application.DisplayAlerts = True
    
    Call Macro_Botão6_Clique
End Sub
