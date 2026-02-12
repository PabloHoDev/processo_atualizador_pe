Option Explicit

Sub AtualizarBasePE()
    Dim wbDest As Workbook
    Dim wbBase As Workbook
    Dim wsDest As Worksheet
    Dim wsBase As Worksheet
    Dim lastRowBase As Long
    Dim lastRowDest As Long
    Dim i As Long, j As Long
    Dim obr As String, uf As String
    Dim areaNova As String, statusNovo As String
    Dim areaAtual As String, statusAtual As String
    Dim colObrigacaoBase As Long, colUFBase As Long
    Dim colAreaBase As Long, colStatusBase As Long
    Dim colObrigacaoDest As Long, colAreaDest As Long, colStatusDest As Long
    Dim logWs As Worksheet
    Dim logRow As Long
    Dim foundRow As Long
    Dim colsAtualizadas As String
    Dim aba As Worksheet
    Dim abaProtegida As Boolean

    ' ====== CONFIGURAÇÃO ======
    Const SENHA_ABA As String = ""

    ' ======================
    ' Configuração de arquivos
    ' ======================
    Set wbDest = Workbooks("Backlog_2024 e 2025.xlsx")      ' Arquivo destino aberto
    Set wbBase = Workbooks("RPE.csv")       ' Base unificada aberta
    Set wsBase = wbBase.Sheets("RPE")        ' Aba da base unificada

    ' Criar ou limpar LOG
    On Error Resume Next
    Set logWs = wbBase.Sheets("LOG_ATUALIZACAO")
    If logWs Is Nothing Then
        Set logWs = wbBase.Sheets.Add
        logWs.Name = "LOG_ATUALIZACAO"
    Else
        logWs.Cells.Clear
    End If
    On Error GoTo 0

    logWs.Range("A1:L1").Value = Array("DataHora", "Aba", "Obrigação", "UF", "ColAtualizadas", _
                                      "OldArea", "NewArea", "OldStatus", "NewStatus", "Tipo", "Mensagem", "LinhaBase")
    logRow = 2

    ' ======================

    ' ======================
    With wsBase
        colObrigacaoBase = Application.Match("Obrigação", .Rows(1), 0)
        colUFBase = Application.Match("UF Prestador", .Rows(1), 0)
        colAreaBase = Application.Match("Área da Pendência", .Rows(1), 0)
        colStatusBase = Application.Match("Status", .Rows(1), 0)
        lastRowBase = .Cells(.Rows.Count, colObrigacaoBase).End(xlUp).Row
    End With

    ' ======================
    ' Iterar sobre cada linha da base unificada
    ' ======================
    For i = 2 To lastRowBase

        ' ======== Barra de Progresso ========
        Dim percentual As Double
        percentual = (i - 1) / (lastRowBase - 1) * 100
        Application.StatusBar = "Atualizando Base Geral PE: " & Format(percentual, "0.0") & "% concluído"
        ' ===================================

        ' Ler valores da linha base
        With wsBase
            obr = Trim(CStr(.Cells(i, colObrigacaoBase).Value))
            uf = Trim(CStr(.Cells(i, colUFBase).Value))
            areaNova = Trim(CStr(.Cells(i, colAreaBase).Value))
            statusNovo = Trim(CStr(.Cells(i, colStatusBase).Value))
        End With

        ' Ignorar linhas sem obrigação
        If obr = "" Then GoTo ProximoI

        ' Procurar a obrigação em todas as abas válidas - necessidade de atualizar as nomeclaturas das colunas
        Set wsDest = Nothing
        foundRow = 0

        For Each aba In wbDest.Sheets
            If aba.Name <> "LOG_ATUALIZACAO" _
               And aba.Name <> "VISÃO GERAL - ESTÁGIO" _
               And aba.Name <> "DADOS DE VALIDAÇÃO" _
               And aba.Name <> "INDISPONIVEL TEMPORARIAMENTE" _
               And aba.Name <> "MG" Then
               
                On Error Resume Next
                colObrigacaoDest = Application.Match("OBRIGAÇÃO", aba.Rows(4), 0)
                colAreaDest = Application.Match("ÁREA DA PENDÊNCIA", aba.Rows(4), 0)
                colStatusDest = Application.Match("STATUS", aba.Rows(4), 0)
                On Error GoTo 0
                
                If colObrigacaoDest > 0 Then
                    lastRowDest = aba.Cells(aba.Rows.Count, colObrigacaoDest).End(xlUp).Row
                    For j = 5 To lastRowDest
                        If Trim(CStr(aba.Cells(j, colObrigacaoDest).Value)) = obr Then
                            Set wsDest = aba
                            foundRow = j
                            Exit For
                        End If
                    Next j
                End If
            End If

            If foundRow > 0 Then Exit For
        Next aba

        If wsDest Is Nothing Or foundRow = 0 Then
            ' Log de erro: obrigação não encontrada - importante colocar o full back
            logWs.Cells(logRow, 1).Value = Now
            logWs.Cells(logRow, 2).Value = "(N/A)"
            logWs.Cells(logRow, 3).Value = obr
            logWs.Cells(logRow, 4).Value = uf
            logWs.Cells(logRow, 10).Value = "AVISO"
            logWs.Cells(logRow, 11).Value = "Obrigação não encontrada em nenhuma aba válida"
            logWs.Cells(logRow, 12).Value = i
            logRow = logRow + 1
            GoTo ProximoI
        End If
             
        ' ======================
        ' Verifica e desbloqueia aba se necessário
        ' ======================
        abaProtegida = wsDest.ProtectContents
        If abaProtegida Then
            On Error Resume Next
            wsDest.Unprotect Password:=SENHA_ABA
            On Error GoTo 0
        End If
        
        ' ======================
        ' Atualizar valores se houver diferença
        ' ======================
        areaAtual = Trim(CStr(wsDest.Cells(foundRow, colAreaDest).Value))
        statusAtual = Trim(CStr(wsDest.Cells(foundRow, colStatusDest).Value))
        colsAtualizadas = ""
        
        If areaNova <> "" And areaNova <> areaAtual Then
            wsDest.Cells(foundRow, colAreaDest).Value = areaNova
            colsAtualizadas = colsAtualizadas & "Área da Pendência;"
        End If
        
        If statusNovo <> "" And statusNovo <> statusAtual Then
            wsDest.Cells(foundRow, colStatusDest).Value = statusNovo
            colsAtualizadas = colsAtualizadas & "Status;"
        End If
        
        ' ======================
        ' Reprotege aba, se estava protegida
        ' ======================
        If abaProtegida Then
            On Error Resume Next
            wsDest.Protect Password:=SENHA_ABA, _
                DrawingObjects:=False, _
                Contents:=True, _
                Scenarios:=True, _
                AllowFormattingCells:=False, _
                AllowFormattingColumns:=False, _
                AllowFormattingRows:=False, _
                AllowInsertingColumns:=False, _
                AllowInsertingRows:=False, _
                AllowInsertingHyperlinks:=False, _
                AllowDeletingColumns:=False, _
                AllowDeletingRows:=False, _
                AllowSorting:=False, _
                AllowFiltering:=True, _
                AllowUsingPivotTables:=True
            On Error GoTo 0
        End If
        
        ' ======================
        ' Registrar log
        ' ======================
        logWs.Cells(logRow, 1).Value = Now
        logWs.Cells(logRow, 2).Value = wsDest.Name
        logWs.Cells(logRow, 3).Value = obr
        logWs.Cells(logRow, 4).Value = uf
        logWs.Cells(logRow, 5).Value = colsAtualizadas
        logWs.Cells(logRow, 6).Value = areaAtual
        logWs.Cells(logRow, 7).Value = areaNova
        logWs.Cells(logRow, 8).Value = statusAtual
        logWs.Cells(logRow, 9).Value = statusNovo
        
        If colsAtualizadas <> "" Then
            logWs.Cells(logRow, 10).Value = "OK"
            logWs.Cells(logRow, 11).Value = "Atualizado"
        Else
            logWs.Cells(logRow, 10).Value = "SEM_MUDANCA"
            logWs.Cells(logRow, 11).Value = "Valores iguais"
        End If
        
        logWs.Cells(logRow, 12).Value = i
        logRow = logRow + 1
        
ProximoI:
    Next i
    
    Application.StatusBar = False
    MsgBox "Atualização concluída! Verifique a aba LOG_ATUALIZACAO na base unificada.", vbInformation
End Sub


