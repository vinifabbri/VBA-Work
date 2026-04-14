Attribute VB_Name = "Módulo2"
Option Explicit

Public Sub LancarPedidoMultiploEGerarObservacoes()

    Dim wsConfig As Worksheet
    Dim wsAprov As Worksheet
    Dim wsMeta As Worksheet
    
    Dim abaAprov As String, abaMeta As String
    Dim colItem As String, colDescricao As String
    Dim colValorMO As String, colValorMAT As String
    Dim colCustoTotal As String, colConsumo As String
    Dim colResto As String, colInicioPedidos As String
    
    Dim pedidoCompra As String
    Dim linhaSubitem As Long, linhaMacro As Long
    Dim valorDigitado As Variant, valorLancado As Double
    Dim colDestino As Long
    Dim continuar As VbMsgBoxResult
    
    Dim dictValoresPorMacro As Object
    Dim arrMacros() As Long
    Dim i As Long
    Dim chave As Variant
    Dim observacaoFinal As String
    Dim bloco As String
    
    On Error GoTo TratarErro
    
    Set wsConfig = ThisWorkbook.Worksheets("CONFIG")
    
    abaAprov = LerConfig(wsConfig, "ABA_APROVACAO_MAT")
    abaMeta = LerConfig(wsConfig, "ABA_META")
    colItem = UCase(LerConfig(wsConfig, "COL_ITEM"))
    colDescricao = UCase(LerConfig(wsConfig, "COL_DESCRICAO"))
    colValorMO = UCase(LerConfig(wsConfig, "COL_VALOR_MO"))
    colValorMAT = UCase(LerConfig(wsConfig, "COL_VALOR_MAT"))
    colCustoTotal = UCase(LerConfig(wsConfig, "COL_CUSTO_TOTAL"))
    colConsumo = UCase(LerConfig(wsConfig, "COL_CONSUMO"))
    colResto = UCase(LerConfig(wsConfig, "COL_RESTO"))
    colInicioPedidos = UCase(LerConfig(wsConfig, "COL_INICIO_PEDIDOS"))
    
    If abaAprov = "" Or abaMeta = "" Or colItem = "" Or colDescricao = "" Or _
       colValorMO = "" Or colValorMAT = "" Or colCustoTotal = "" Or _
       colConsumo = "" Or colResto = "" Or colInicioPedidos = "" Then
        MsgBox "Há campos vazios na aba CONFIG.", vbExclamation
        Exit Sub
    End If
    
    Set wsAprov = ThisWorkbook.Worksheets(abaAprov)
    Set wsMeta = ThisWorkbook.Worksheets(abaMeta)
    
    pedidoCompra = Trim(InputBox("QUAL O PEDIDO DE COMPRA?", "Pedido de Compra"))
    If pedidoCompra = "" Then Exit Sub
    
    Set dictValoresPorMacro = CreateObject("Scripting.Dictionary")
    
    Do
    
        linhaSubitem = Application.InputBox( _
            Prompt:="QUAL A LINHA DO SUBITEM?", _
            Title:="Linha do Subitem", _
            Type:=1)
        
        If linhaSubitem = 0 Then Exit Do
        
        If linhaSubitem < 1 Then
            MsgBox "Linha do subitem inválida.", vbExclamation
            GoTo PerguntarContinuacao
        End If
        
        If Trim(wsAprov.Cells(linhaSubitem, ColunaParaNumero(colItem)).Text) = "" Then
            MsgBox "A linha informada não possui item/subitem na aba '" & abaAprov & "'.", vbExclamation
            GoTo PerguntarContinuacao
        End If
        
        valorDigitado = Application.InputBox( _
            Prompt:="QUAL O VALOR?", _
            Title:="Valor do Pedido", _
            Type:=1)
        
        If valorDigitado = False Then GoTo PerguntarContinuacao
        
        valorLancado = CDbl(valorDigitado)
        
        linhaMacro = Application.InputBox( _
            Prompt:="QUAL A LINHA DO MACRO?", _
            Title:="Linha do Macro", _
            Type:=1)
        
        If linhaMacro = 0 Then GoTo PerguntarContinuacao
        
        If linhaMacro < 1 Then
            MsgBox "Linha do macro inválida.", vbExclamation
            GoTo PerguntarContinuacao
        End If
        
        If Trim(wsMeta.Cells(linhaMacro, ColunaParaNumero(colDescricao)).Text) = "" Then
            MsgBox "Linha do macro sem descrição na aba '" & abaMeta & "'.", vbExclamation
            GoTo PerguntarContinuacao
        End If
        
        colDestino = ProximaColunaLivrePedido(wsAprov, linhaSubitem, ColunaParaNumero(colInicioPedidos))
        
        If colDestino = 0 Then
            MsgBox "Não foi encontrado espaço livre para Pedido e Valor na linha " & linhaSubitem & ".", vbExclamation
            GoTo PerguntarContinuacao
        End If
        
        wsAprov.Cells(linhaSubitem, colDestino).Value = pedidoCompra
        wsAprov.Cells(linhaSubitem, colDestino + 1).Value = valorLancado
        wsAprov.Cells(linhaSubitem, colDestino + 1).NumberFormat = "#,##0.00"
        
        If dictValoresPorMacro.Exists(CStr(linhaMacro)) Then
            dictValoresPorMacro(CStr(linhaMacro)) = dictValoresPorMacro(CStr(linhaMacro)) + valorLancado
        Else
            dictValoresPorMacro.Add CStr(linhaMacro), valorLancado
        End If
        
PerguntarContinuacao:
        continuar = MsgBox("Deseja lançar outro item deste mesmo pedido?", vbYesNo + vbQuestion, "Continuar?")
    
    Loop While continuar = vbYes
    
    If dictValoresPorMacro.Count = 0 Then
        MsgBox "Nenhum lançamento válido foi registrado.", vbExclamation
        Exit Sub
    End If
    
    arrMacros = DictionaryKeysToSortedLongArray(dictValoresPorMacro)
    
    observacaoFinal = ""
    
    For i = LBound(arrMacros) To UBound(arrMacros)
    
        bloco = MontarBlocoObservacao( _
            wsMeta:=wsMeta, _
            linhaMacro:=arrMacros(i), _
            valorOrcado:=CDbl(dictValoresPorMacro(CStr(arrMacros(i)))), _
            colDescricao:=colDescricao, _
            colValorMO:=colValorMO, _
            colValorMAT:=colValorMAT, _
            colCustoTotal:=colCustoTotal, _
            colConsumo:=colConsumo, _
            colResto:=colResto)
        
        If observacaoFinal <> "" Then
            observacaoFinal = observacaoFinal & vbCrLf & vbCrLf & String(95, "_") & vbCrLf & vbCrLf
        End If
        
        observacaoFinal = observacaoFinal & bloco
    Next i
    
    frmObservacao.CarregarTexto observacaoFinal
    frmObservacao.Show
    
    Exit Sub

TratarErro:
    MsgBox "Erro ao executar a macro: " & Err.Description, vbCritical

End Sub

Private Function MontarBlocoObservacao( _
    ByVal wsMeta As Worksheet, _
    ByVal linhaMacro As Long, _
    ByVal valorOrcado As Double, _
    ByVal colDescricao As String, _
    ByVal colValorMO As String, _
    ByVal colValorMAT As String, _
    ByVal colCustoTotal As String, _
    ByVal colConsumo As String, _
    ByVal colResto As String) As String
    
    Dim descricao As String
    Dim valorMacro As Double
    Dim consumo As String
    Dim acumulado As Double
    Dim saldo As Double
    Dim valorMat As Double
    Dim valorMo As Double
    
    descricao = Trim(wsMeta.Cells(linhaMacro, ColunaParaNumero(colDescricao)).Text)
    valorMacro = LerNumeroCelula(wsMeta.Cells(linhaMacro, ColunaParaNumero(colValorMAT)))
    consumo = Trim(wsMeta.Cells(linhaMacro, ColunaParaNumero(colConsumo)).Text)
    acumulado = LerNumeroCelula(wsMeta.Cells(linhaMacro, ColunaParaNumero(colCustoTotal)))
    saldo = LerNumeroCelula(wsMeta.Cells(linhaMacro, ColunaParaNumero(colResto)))
    valorMat = LerNumeroCelula(wsMeta.Cells(linhaMacro, ColunaParaNumero(colValorMAT)))
    valorMo = LerNumeroCelula(wsMeta.Cells(linhaMacro, ColunaParaNumero(colValorMO)))
    
    MontarBlocoObservacao = _
        "DESCRIÇÃO: " & descricao & vbCrLf & vbCrLf & _
        "VALOR TOTAL MACRO MATERIAL: " & FormatarMoedaBR(valorMacro) & _
        "    /    CONSUMIDO: " & consumo & vbCrLf & vbCrLf & vbCrLf & _
        "VALOR ORÇADO: " & FormatarMoedaBR(valorOrcado) & vbCrLf & _
        "ACUMULADO: " & FormatarMoedaBR(acumulado) & vbCrLf & _
        "SALDO: " & FormatarNumeroComSinal(saldo) & vbCrLf & vbCrLf & vbCrLf & _
        "VALOR TOTAL MATERIAL E MÃO DE OBRA:" & vbCrLf & vbCrLf & _
        "MAT.: " & FormatarMoedaBR(valorMat) & vbCrLf & _
        "M.O.: " & FormatarMoedaBR(valorMo)
End Function

Private Function LerConfig(ByVal ws As Worksheet, ByVal chave As String) As String
    Dim ultimaLinha As Long
    Dim i As Long
    
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To ultimaLinha
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(Trim(chave)) Then
            LerConfig = Trim(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    
    LerConfig = ""
End Function

Private Function ColunaParaNumero(ByVal letraColuna As String) As Long
    ColunaParaNumero = Range(UCase(letraColuna) & "1").Column
End Function

Private Function ProximaColunaLivrePedido(ByVal ws As Worksheet, ByVal linha As Long, ByVal colInicial As Long) As Long
    Dim c As Long
    Dim ultimaCol As Long
    
    ultimaCol = ws.Cells(linha, ws.Columns.Count).End(xlToLeft).Column
    
    If ultimaCol < colInicial Then
        ultimaCol = colInicial
    End If
    
    ultimaCol = ultimaCol + 20
    
    For c = colInicial To ultimaCol Step 2
        If CelulaVaziaReal(ws.Cells(linha, c)) And CelulaVaziaReal(ws.Cells(linha, c + 1)) Then
            ProximaColunaLivrePedido = c
            Exit Function
        End If
    Next c
    
    ProximaColunaLivrePedido = 0
End Function

Private Function CelulaVaziaReal(ByVal cel As Range) As Boolean
    Dim txt As String
    
    txt = CStr(cel.Value)
    txt = Replace(txt, Chr(160), "")
    txt = Trim(txt)
    
    CelulaVaziaReal = (txt = "")
End Function

Private Function LerNumeroCelula(ByVal cel As Range) As Double
    Dim s As String
    
    On Error GoTo falha
    
    If IsNumeric(cel.Value) Then
        LerNumeroCelula = CDbl(cel.Value)
        Exit Function
    End If
    
    s = Trim(cel.Text)
    s = Replace(s, "R$", "")
    s = Replace(s, " ", "")
    s = Replace(s, ".", "")
    s = Replace(s, ",", ".")
    
    If s = "" Or s = "-" Then
        LerNumeroCelula = 0
    ElseIf IsNumeric(s) Then
        LerNumeroCelula = CDbl(s)
    Else
        LerNumeroCelula = 0
    End If
    
    Exit Function

falha:
    LerNumeroCelula = 0
End Function

Private Function FormatarNumeroBR(ByVal valor As Double) As String
    FormatarNumeroBR = Format(valor, "#,##0.00")
End Function

Private Function FormatarMoedaBR(ByVal valor As Double) As String
    If valor < 0 Then
        FormatarMoedaBR = "-R$" & Format(Abs(valor), "#,##0.00")
    Else
        FormatarMoedaBR = "R$" & Format(valor, "#,##0.00")
    End If
End Function

Private Function FormatarNumeroComSinal(ByVal valor As Double) As String
    If valor < 0 Then
        FormatarNumeroComSinal = "-R$" & Format(Abs(valor), "#,##0.00")
    Else
        FormatarNumeroComSinal = "R$" & Format(valor, "#,##0.00")
    End If
End Function

Private Function DictionaryKeysToSortedLongArray(ByVal dict As Object) As Long()
    Dim arr() As Long
    Dim i As Long
    Dim j As Long
    Dim temp As Long
    Dim k As Variant
    
    ReDim arr(0 To dict.Count - 1)
    
    i = 0
    For Each k In dict.Keys
        arr(i) = CLng(k)
        i = i + 1
    Next k
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    DictionaryKeysToSortedLongArray = arr
End Function

