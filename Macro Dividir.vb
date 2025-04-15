Sub Macro_suprema()
    Application.DisplayAlerts = False
    Dim wsBase As Worksheet
    Dim wsBotao As Worksheet
    Dim transportador As Range
    Dim centro As Range
    Dim transportadoresUnicos As Collection
    Dim centrosUnicos As Collection
    Dim novaPlanilha As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long, j As Long
    Dim nomePlanilha As String
    Dim comandoLogPlanilha As Worksheet
    Dim comandoLogCriada As Boolean
    Dim wsBravoDCpaulinia As Worksheet
    Dim wsBravoUberaba As Worksheet
    Dim ws As Worksheet
    Dim FilePath As String
    Dim FileName As String
    Dim FullPath As String
    Dim TodayDate As String
    Dim NewWorkbook As Workbook
    Dim wsOutrosTransp As Worksheet
    Dim outrosTranspCriada As Boolean
    
    Set wsBase = ThisWorkbook.Sheets("Base")
    Set wsBotao = ThisWorkbook.Sheets("Macro")
    ultimaLinha = wsBase.Cells(wsBase.Rows.Count, "H").End(xlUp).Row
    Set transportadoresUnicos = New Collection
    On Error Resume Next
    For Each transportador In wsBase.Range("H2:H" & ultimaLinha)
        transportadoresUnicos.Add Trim(transportador.Value), CStr(Trim(transportador.Value))
    Next transportador
    On Error GoTo 0
    comandoLogCriada = False
    outrosTranspCriada = False
    
    For i = 1 To transportadoresUnicos.Count
        wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=8, Criteria1:=transportadoresUnicos(i)
        Set centrosUnicos = New Collection
        On Error Resume Next
        For Each centro In wsBase.Range("E2:E" & ultimaLinha).SpecialCells(xlCellTypeVisible)
            centrosUnicos.Add Trim(centro.Value), CStr(Trim(centro.Value))
        Next centro
        On Error GoTo 0
        
        If transportadoresUnicos(i) <> "BRAVO" And transportadoresUnicos(i) <> "TONIATO" And transportadoresUnicos(i) <> "LUFT" And transportadoresUnicos(i) <> "COMANDO LOG" Then
            If Not outrosTranspCriada Then
                Set wsOutrosTransp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                wsOutrosTransp.Name = "OUTROS TRANSP."
                outrosTranspCriada = True
                wsBase.Rows(1).Copy Destination:=wsOutrosTransp.Rows(1)
            End If
            wsBase.Range("A2:S" & ultimaLinha).SpecialCells(xlCellTypeVisible).Copy _
                Destination:=wsOutrosTransp.Rows(wsOutrosTransp.Cells(wsOutrosTransp.Rows.Count, "A").End(xlUp).Row + 1)
        ElseIf transportadoresUnicos(i) = "COMANDO LOG" Then
            If Not comandoLogCriada Then
                Set comandoLogPlanilha = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                comandoLogPlanilha.Name = "COMANDO LOG"
                comandoLogCriada = True
                wsBase.Rows(1).Copy Destination:=comandoLogPlanilha.Rows(1)
            End If
            wsBase.Range("A2:S" & ultimaLinha).SpecialCells(xlCellTypeVisible).Copy _
                Destination:=comandoLogPlanilha.Rows(comandoLogPlanilha.Cells(comandoLogPlanilha.Rows.Count, "A").End(xlUp).Row + 1)
        Else
            For j = 1 To centrosUnicos.Count
                nomePlanilha = transportadoresUnicos(i) & " - " & Replace(centrosUnicos(j), "CP", "")
                If Len(nomePlanilha) > 31 Then
                    nomePlanilha = Left(nomePlanilha, 31)
                End If
                Select Case nomePlanilha
                    Case "BRAVO - DC Cuiaba "
                        nomePlanilha = "BRAVO - CUIABÁ"
                    Case "BRAVO - DC Uberaba "
                        nomePlanilha = "BRAVO - UBERABA"
                    Case "BRAVO - DC LEM "
                        nomePlanilha = "BRAVO - LEM"
                    Case "TONIATO - DC Paulinia "
                        nomePlanilha = "TONIATO - PAULÍNIA"
                    Case "TONIATO - DC Ibipora "
                        nomePlanilha = "TONIATO - IBIPORÃ"
                    Case "TONIATO - DC Rio Verde "
                        nomePlanilha = "TONIATO - RIO VERDE"
                    Case "TONIATO - WH Belford Roxo "
                        nomePlanilha = "TONIATO - B. ROXO"
                    Case "LUFT - DC Paulinia "
                        nomePlanilha = "LUFT - PAULÍNIA"
                    Case "LUFT - DC Carazinho "
                        nomePlanilha = "LUFT - CARAZINHO"
                    Case "BRAVO - DC Querencia "
                        nomePlanilha = "BRAVO - QUERÊNCIA"
                End Select
                
                wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=8, Criteria1:=transportadoresUnicos(i)
                wsBase.Range("A1:S" & ultimaLinha).AutoFilter Field:=5, Criteria1:=centrosUnicos(j)
                
                Set novaPlanilha = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                novaPlanilha.Name = nomePlanilha
                wsBase.Rows(1).Copy Destination:=novaPlanilha.Rows(1)
                wsBase.Range("A2:S" & ultimaLinha).SpecialCells(xlCellTypeVisible).Copy Destination:=novaPlanilha.Rows(2)
                
                wsBase.AutoFilterMode = False
            Next j
        End If
    Next i
    
    On Error Resume Next
    Set wsBravoDCpaulinia = ThisWorkbook.Sheets("BRAVO - DC Paulinia ")
    Set wsBravoUberaba = ThisWorkbook.Sheets("BRAVO - UBERABA")
    On Error GoTo 0
    If Not wsBravoDCpaulinia Is Nothing And Not wsBravoUberaba Is Nothing Then
        ultimaLinha = wsBravoUberaba.Cells(wsBravoUberaba.Rows.Count, "A").End(xlUp).Row + 1
        wsBravoDCpaulinia.Range("A2:S" & wsBravoDCpaulinia.Cells(wsBravoDCpaulinia.Rows.Count, "A").End(xlUp).Row).Copy Destination:=wsBravoUberaba.Range("A" & ultimaLinha)
        Application.DisplayAlerts = False
        wsBravoDCpaulinia.Delete
    End If
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Columns("A:S").AutoFit
        ws.Activate
        ws.Range("A2").Select
        ActiveWindow.FreezePanes = True
    Next ws
    
    ' Define o caminho e o nome do arquivo
    FilePath = "C:\Users\vinicius.domingues\Documents\Projeto\Base Bayer\"
    TodayDate = Format(Date, "dd.mm")
    FileName = "Tracking CP - " & TodayDate & ".xlsx"
    FullPath = FilePath & FileName
    
    Set NewWorkbook = Workbooks.Add
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Macro" And ws.Name <> "Base" And ws.Name <> "Retorno" Then
            ws.Copy After:=NewWorkbook.Sheets(NewWorkbook.Sheets.Count)
        End If
    Next ws
    
    Application.DisplayAlerts = False
    NewWorkbook.Sheets(1).Delete
    
    NewWorkbook.SaveAs FullPath, FileFormat:=xlOpenXMLWorkbook
    
    NewWorkbook.Close False
    
    MsgBox "Arquivo salvo como: " & FullPath
    
    Application.DisplayAlerts = False

    For Each ws In ThisWorkbook.Worksheets
        nomePlanilha = ws.Name
        If nomePlanilha <> "Base" And nomePlanilha <> "Macro" And nomePlanilha <> "Retorno" And _
       nomePlanilha <> "SAP" And nomePlanilha <> "Itinerario" Then 
        ws.Delete
        End If
    Next ws
    
    MsgBox "Processo Finalizado!"
    
    Application.DisplayAlerts = True
End Sub
