Sub Button10_Click()
    'Chamar TRP
    Dim wsRetorno As Worksheet, wsBase As Worksheet
    Dim valorTransportadora As String
    Dim valorBravo As String
    Dim valorToniato As String
    Dim caminhoArquivo As String
    Dim dialogo As FileDialog
    
    Set wsRetorno = ThisWorkbook.Sheets("Retorno")
    Set wsBase = ThisWorkbook.Sheets("Base")
    
    On Error Resume Next
    valorTransportadora = UCase(Trim(wsRetorno.Cells(2, "R").Value))
    valorBravo = UCase(Trim(wsRetorno.Cells(2, "I").Value))
    valorToniato = UCase(Trim(wsRetorno.Cells(2, "B").Value))
    On Error GoTo 0
    
    Select Case True
        Case valorTransportadora = "LUFT"
            Call Luft
        Case InStr(1, valorBravo, "BRAVO") > 0
            Call Bravo
        Case valorToniato = "GT SOLUCOES LOGISTICAS SA" Or valorToniato = "TRANSPORTES TONIATO LTDA"
            Call Toniato
        Case Else
    End Select
    
    Call Macro_Botão6_Clique
    
    Set dialogo = Application.FileDialog(msoFileDialogFolderPicker)
    dialogo.Title = "Selecione a pasta para salvar o arquivo"
    
    If dialogo.Show = -1 Then
        caminhoArquivo = dialogo.SelectedItems(1) & "\CP Report Fixo Consolidado.xlsx"
    Else
        MsgBox "Operação cancelada. O arquivo não foi salvo.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Application.DisplayAlerts = False
    
    wsBase.Copy
    ActiveWorkbook.SaveAs FileName:=caminhoArquivo, FileFormat:=xlOpenXMLWorkbook
    ActiveWorkbook.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    MsgBox "Processo concluído! Arquivo salvo em:" & vbNewLine & caminhoArquivo, vbInformation
End Sub
