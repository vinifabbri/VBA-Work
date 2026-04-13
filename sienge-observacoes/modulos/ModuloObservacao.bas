Option Explicit

'=========================================================
' MACRO PRINCIPAL
'=========================================================
Public Sub LancarPedidoEGerarObservacao()

    Dim wsConfig As Worksheet
    Dim wsAprov As Worksheet
    Dim wsMeta As Worksheet
    
    Dim abaAprov As String
    Dim abaMeta As String
    
    Dim colItem As String
    Dim colDescricao As String
    Dim colValorMO As String
    Dim colValorMAT As String
    Dim colCustoTotal As String
    Dim colConsumo As String
    Dim colResto As String
    Dim colInicioPedidos As String
    
    Dim linhaSubitem As Long
    Dim linhaMacro As Long
    Dim pedidoCompra As String
    Dim valorDigitado As Variant
    Dim valorOrcado As Double
    Dim colDestino As Long
    
    Dim nomeMacro As String
    Dim valorMacro As Double
    Dim consumoMacro As String
    Dim custoTotal As Double
    Dim resto As Double
    Dim valorMat As Double
    Dim valorMo As Double
    
    Dim observacao As String
    
    On Error GoTo TratarErro
    
    Set wsConfig = ThisWorkbook.Worksheets("CONFIG")
    
    '=====================================================
    ' LER CONFIGURAÇÕES
    '=====================================================
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
    
    '=====================================================
    ' 1) LINHA DO SUBITEM
    '=====================================================
    linhaSubitem = Application.InputBox( _
        Prompt:="QUAL A LINHA DO SUBITEM?", _
        Title:="Linha do Subitem", _
        Type:=1)
    
    If linhaSubitem = 0 Then Exit Sub
    
    If linhaSubitem < 1 Then
        MsgBox "Linha do subitem inválida.", vbExclamation
        Exit Sub
    End If
    
    If Trim(wsAprov.Cells(linhaSubitem, ColunaParaNumero(colItem)).Text) = "" Then
        MsgBox "A linha informada não possui item/subitem na aba '" & abaAprov & "'.", vbExclamation
        Exit Sub
    End If
    
    '=====================================================
    ' 2) ENCONTRA PRÓXIMO PAR LIVRE
    '=====================================================
    colDestino = ProximaColunaLivrePedido(wsAprov, linhaSubitem, ColunaParaNumero(colInicioPedidos))
    
    If colDestino = 0 Then
        MsgBox "Não foi encontrado espaço livre para Pedido e Valor.", vbExclamation
        Exit Sub
    End If
    
    '=====================================================
    ' 3) PEDIDO
    '=====================================================
    pedidoCompra = Trim(InputBox("QUAL O PEDIDO DE COMPRA?", "Pedido de Compra"))
    If pedidoCompra = "" Then Exit Sub
    
    '=====================================================
    ' 4) VALOR
    '=====================================================
    valorDigitado = Application.InputBox( _
        Prompt:="QUAL O VALOR?", _
        Title:="Valor do Pedido", _
        Type:=1)
    
    If valorDigitado = False Then Exit Sub
    
    valorOrcado = CDbl(valorDigitado)
    
    '=====================================================
    ' 5) LINHA DO MACRO
    '=====================================================
    linhaMacro = Application.InputBox( _
        Prompt:="QUAL A LINHA DO MACRO?", _
        Title:="Linha do Macro", _
        Type:=1)
    
    If linhaMacro = 0 Then Exit Sub
    
    If linhaMacro < 1 Then
        MsgBox "Linha do macro inválida.", vbExclamation
        Exit Sub
    End If
    
    If Trim(wsMeta.Cells(linhaMacro, ColunaParaNumero(colDescricao)).Text) = "" Then
        MsgBox "Linha do macro sem descrição na aba '" & abaMeta & "'. Verifique.", vbExclamation
        Exit Sub
    End If
    
    '=====================================================
    ' 6) GRAVA NA ABA APROVAÇÃO
    '=====================================================
    wsAprov.Cells(linhaSubitem, colDestino).Value = pedidoCompra
    wsAprov.Cells(linhaSubitem, colDestino + 1).Value = valorOrcado
    wsAprov.Cells(linhaSubitem, colDestino + 1).NumberFormat = "#,##0.00"
    
    '=====================================================
    ' 7) BUSCA DADOS DO MACRO NA ABA META
    '=====================================================
    nomeMacro = Trim(wsMeta.Cells(linhaMacro, ColunaParaNumero(colDescricao)).Text)
    valorMacro = LerNumeroCelula(wsMeta.Cells(linhaMacro, ColunaParaNumero(colValorMAT)))
    consumoMacro = Trim(wsMeta.Cells(linhaMacro, ColunaParaNumero(colConsumo)).Text)
    
    custoTotal = LerNumeroCelula(wsMeta.Cells(linhaMacro, ColunaParaNumero(colCustoTotal)))
    resto = LerNumeroCelula(wsMeta.Cells(linhaMacro, ColunaParaNumero(colResto)))
    
    valorMat = LerNumeroCelula(wsMeta.Cells(linhaMacro, ColunaParaNumero(colValorMAT)))
    valorMo = LerNumeroCelula(wsMeta.Cells(linhaMacro, ColunaParaNumero(colValorMO)))
    
    '=====================================================
    ' 8) MONTA OBSERVAÇÃO
    '=====================================================
    observacao = _
        PadRight("NOME DO MACRO:", 18) & nomeMacro & vbCrLf & _
        PadRight("VALOR:", 18) & FormatarMoedaBR(valorMacro) & vbCrLf & _
        PadRight("CONSUMO:", 18) & consumoMacro & vbCrLf & vbCrLf & _
        PadRight("VALOR ORÇADO:", 18) & FormatarMoedaBR(valorOrcado) & vbCrLf & _
        PadRight("CUSTO TOTAL:", 18) & FormatarMoedaBR(custoTotal) & vbCrLf & _
        PadRight("RESTO:", 18) & FormatarNumeroComSinal(resto) & vbCrLf & vbCrLf & _
        PadRight("VALOR MAT:", 18) & FormatarMoedaBR(valorMat) & vbCrLf & _
        PadRight("VALOR MO:", 18) & FormatarMoedaBR(valorMo)
    
    frmObservacao.CarregarTexto observacao
    frmObservacao.Show
    
    Exit Sub

TratarErro:
    MsgBox "Erro ao executar a macro: " & Err.Description, vbCritical

End Sub

'=========================================================
' LÊ PARÂMETRO NA ABA CONFIG
' coluna A = nome do parâmetro
' coluna B = valor
'=========================================================
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

'=========================================================
' CONVERTE LETRA DA COLUNA EM NÚMERO
' Ex: G -> 7 | AD -> 30
'=========================================================
Private Function ColunaParaNumero(ByVal letraColuna As String) As Long
    ColunaParaNumero = Range(UCase(letraColuna) & "1").Column
End Function

'=========================================================
' PROCURA O PRIMEIRO PAR LIVRE NA LINHA
' Ex.: E/F, G/H, I/J...
' começa na coluna configurada e anda até achar o próximo par vazio
'=========================================================
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

'=========================================================
' VERIFICA SE A CÉLULA ESTÁ REALMENTE VAZIA
'=========================================================
Private Function CelulaVaziaReal(ByVal cel As Range) As Boolean
    
    Dim txt As String
    
    txt = CStr(cel.Value)
    txt = Replace(txt, Chr(160), "")
    txt = Trim(txt)
    
    CelulaVaziaReal = (txt = "")

End Function

'=========================================================
' LÊ NÚMERO DA CÉLULA
'=========================================================
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

'=========================================================
' FORMATAÇÕES
'=========================================================
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

Private Function PadRight(ByVal texto As String, ByVal tamanho As Long) As String
    If Len(texto) >= tamanho Then
        PadRight = texto
    Else
        PadRight = texto & Space(tamanho - Len(texto))
    End If
End Function


