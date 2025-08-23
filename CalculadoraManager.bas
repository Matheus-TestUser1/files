' ====================================================================
' MÓDULO CALCULADORA MANAGER - SISTEMA PDV MADEIREIRA MARIA LUZIA
' Responsável por todos os cálculos do sistema
' ====================================================================

Option Explicit

' === CALCULAR TOTAL GERAL ===
Public Function CalcularTotalGeral(lstSelecionados As MSForms.ListBox, Optional frete As Double = 0, Optional descontoAdicional As Double = 0) As Double
    On Error GoTo TratarErro
    
    Dim totalProdutos As Double
    Dim totalDescontos As Double
    Dim totalFinal As Double
    
    ' Calcular total dos produtos
    totalProdutos = ProdutoManager.CalcularTotalProdutos(lstSelecionados)
    
    ' Calcular total de descontos
    totalDescontos = DescontoManager.CalcularTotalDescontos(lstSelecionados)
    
    ' Calcular total final
    totalFinal = totalProdutos + frete - descontoAdicional
    
    CalcularTotalGeral = totalFinal
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("CalcularTotalGeral", Err)
    CalcularTotalGeral = 0
End Function

' === CALCULAR DESCONTO SOBRE TOTAL ===
Public Function CalcularDescontoSobreTotal(valorTotal As Double, percentualDesconto As Double) As Double
    On Error GoTo TratarErro
    
    If percentualDesconto < 0 Or percentualDesconto > 100 Then
        CalcularDescontoSobreTotal = 0
        Exit Function
    End If
    
    CalcularDescontoSobreTotal = valorTotal * (percentualDesconto / 100)
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("CalcularDescontoSobreTotal", Err)
    CalcularDescontoSobreTotal = 0
End Function

' === CALCULAR FRETE ===
Public Function CalcularFrete(cep As String, valorTotal As Double, peso As Double) As Double
    On Error GoTo TratarErro
    
    ' Frete grátis para pedidos acima de R$ 500
    If valorTotal >= 500 Then
        CalcularFrete = 0
        Exit Function
    End If
    
    ' Extrair CEP numérico
    Dim cepNumerico As String
    cepNumerico = Replace(Replace(cep, "-", ""), " ", "")
    
    If Len(cepNumerico) <> 8 Or Not IsNumeric(cepNumerico) Then
        ' CEP inválido, frete padrão
        CalcularFrete = 25
        Exit Function
    End If
    
    Dim cepNum As Long
    cepNum = CLng(cepNumerico)
    
    ' Tabela de frete por região (CEP de PE)
    Select Case True
        Case cepNum >= 53000000 And cepNum <= 53999999 ' Região Metropolitana
            CalcularFrete = IIf(peso <= 50, 15, 25)
            
        Case cepNum >= 50000000 And cepNum <= 52999999 ' Recife e adjacências
            CalcularFrete = IIf(peso <= 50, 20, 35)
            
        Case cepNum >= 54000000 And cepNum <= 56999999 ' Interior
            CalcularFrete = IIf(peso <= 50, 30, 50)
            
        Case Else ' Fora de PE
            CalcularFrete = IIf(peso <= 50, 40, 80)
    End Select
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("CalcularFrete", Err)
    CalcularFrete = 25 ' Frete padrão em caso de erro
End Function

' === CALCULAR PESO TOTAL ===
Public Function CalcularPesoTotal(lstSelecionados As MSForms.ListBox) As Double
    On Error GoTo TratarErro
    
    Dim pesoTotal As Double
    pesoTotal = 0
    
    ' Tabela de peso por categoria (kg por unidade)
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        Dim referencia As String
        Dim quantidade As Long
        Dim pesoPorUnidade As Double
        
        referencia = lstSelecionados.List(i, 0)
        quantidade = CLng(lstSelecionados.List(i, 4))
        
        ' Obter categoria do produto
        Dim categoria As String
        categoria = ObterCategoriaProduto(referencia)
        
        ' Definir peso por categoria
        Select Case UCase(categoria)
            Case "MADEIRA"
                pesoPorUnidade = 5 ' 5kg por peça de madeira
            Case "FERRAGEM"
                pesoPorUnidade = 0.5 ' 0.5kg por item de ferragem
            Case "TINTAS"
                pesoPorUnidade = 20 ' 20kg por lata de tinta
            Case "ADESIVOS"
                pesoPorUnidade = 1 ' 1kg por item de cola
            Case Else
                pesoPorUnidade = 2 ' Peso padrão
        End Select
        
        pesoTotal = pesoTotal + (pesoPorUnidade * quantidade)
    Next i
    
    CalcularPesoTotal = pesoTotal
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("CalcularPesoTotal", Err)
    CalcularPesoTotal = 0
End Function

' === OBTER CATEGORIA DO PRODUTO ===
Private Function ObterCategoriaProduto(referencia As String) As String
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = UCase(Trim(referencia)) Then
            ObterCategoriaProduto = ws.Cells(i, 3).Value
            Exit Function
        End If
    Next i
    
    ObterCategoriaProduto = "OUTROS"
    
    Exit Function
TratarErro:
    ObterCategoriaProduto = "OUTROS"
End Function

' === CALCULAR TROCO ===
Public Function CalcularTroco(valorTotal As Double, valorPago As Double) As Double
    On Error GoTo TratarErro
    
    If valorPago >= valorTotal Then
        CalcularTroco = valorPago - valorTotal
    Else
        CalcularTroco = 0
    End If
    
    Exit Function
TratarErro:
    CalcularTroco = 0
End Function

' === CALCULAR PARCELAS ===
Public Function CalcularParcelas(valorTotal As Double, numeroParcelas As Integer) As Double
    On Error GoTo TratarErro
    
    If numeroParcelas <= 0 Then
        CalcularParcelas = valorTotal
        Exit Function
    End If
    
    ' Sem juros até 3 parcelas
    If numeroParcelas <= 3 Then
        CalcularParcelas = valorTotal / numeroParcelas
    Else
        ' Com juros de 2% ao mês para mais de 3 parcelas
        Dim taxa As Double
        taxa = 0.02 ' 2% ao mês
        
        Dim valorComJuros As Double
        valorComJuros = valorTotal * (1 + (taxa * (numeroParcelas - 3)))
        
        CalcularParcelas = valorComJuros / numeroParcelas
    End If
    
    Exit Function
TratarErro:
    CalcularParcelas = valorTotal
End Function

' === CALCULAR MARGEM DE LUCRO ===
Public Function CalcularMargemLucro(lstSelecionados As MSForms.ListBox) As Double
    On Error GoTo TratarErro
    
    Dim totalCusto As Double
    Dim totalVenda As Double
    
    totalCusto = 0
    totalVenda = 0
    
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        Dim referencia As String
        Dim quantidade As Long
        Dim precoCusto As Double
        Dim valorTotalItem As Double
        
        referencia = lstSelecionados.List(i, 0)
        quantidade = CLng(lstSelecionados.List(i, 4))
        
        ' Obter preço de custo
        precoCusto = ObterPrecoCusto(referencia)
        
        ' Obter valor total do item
        Dim valorTexto As String
        valorTexto = lstSelecionados.List(i, 6)
        valorTotalItem = ErrorHandler.ConverterTextoParaValor(valorTexto)
        
        totalCusto = totalCusto + (precoCusto * quantidade)
        totalVenda = totalVenda + valorTotalItem
    Next i
    
    If totalCusto > 0 Then
        CalcularMargemLucro = ((totalVenda - totalCusto) / totalVenda) * 100
    Else
        CalcularMargemLucro = 0
    End If
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("CalcularMargemLucro", Err)
    CalcularMargemLucro = 0
End Function

' === OBTER PREÇO DE CUSTO ===
Private Function ObterPrecoCusto(referencia As String) As Double
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = UCase(Trim(referencia)) Then
            ObterPrecoCusto = CDbl(ws.Cells(i, 5).Value) ' Coluna E = Preco_Custo
            Exit Function
        End If
    Next i
    
    ObterPrecoCusto = 0
    
    Exit Function
TratarErro:
    ObterPrecoCusto = 0
End Function

' === CALCULAR DESCONTO POR FORMA DE PAGAMENTO ===
Public Function CalcularDescontoPagamento(valorTotal As Double, formaPagamento As String) As Double
    On Error GoTo TratarErro
    
    Select Case True
        Case InStr(UCase(formaPagamento), "À VISTA") > 0
            CalcularDescontoPagamento = valorTotal * 0.05 ' 5% à vista
            
        Case InStr(UCase(formaPagamento), "PIX") > 0
            CalcularDescontoPagamento = valorTotal * 0.03 ' 3% PIX
            
        Case Else
            CalcularDescontoPagamento = 0 ' Sem desconto
    End Select
    
    Exit Function
TratarErro:
    CalcularDescontoPagamento = 0
End Function

' === GERAR RESUMO FINANCEIRO ===
Public Function GerarResumoFinanceiro(lstSelecionados As MSForms.ListBox, formaPagamento As String, frete As Double) As String
    On Error GoTo TratarErro
    
    Dim resumo As String
    Dim subtotal As Double
    Dim totalDescontos As Double
    Dim descontoPagamento As Double
    Dim totalFinal As Double
    Dim margemLucro As Double
    
    ' Calcular valores
    subtotal = ProdutoManager.CalcularTotalProdutos(lstSelecionados)
    totalDescontos = DescontoManager.CalcularTotalDescontos(lstSelecionados)
    descontoPagamento = CalcularDescontoPagamento(subtotal, formaPagamento)
    totalFinal = subtotal + frete - descontoPagamento
    margemLucro = CalcularMargemLucro(lstSelecionados)
    
    ' Montar resumo
    resumo = "💰 RESUMO FINANCEIRO" & vbCrLf & vbCrLf
    resumo = resumo & "Subtotal (produtos): " & Format(subtotal, "R$ #,##0.00") & vbCrLf
    resumo = resumo & "Descontos por item: " & Format(totalDescontos, "R$ #,##0.00") & vbCrLf
    resumo = resumo & "Desconto pagamento: " & Format(descontoPagamento, "R$ #,##0.00") & " (" & formaPagamento & ")" & vbCrLf
    resumo = resumo & "Frete: " & Format(frete, "R$ #,##0.00") & vbCrLf
    resumo = resumo & String(30, "-") & vbCrLf
    resumo = resumo & "TOTAL FINAL: " & Format(totalFinal, "R$ #,##0.00") & vbCrLf
    resumo = resumo & "Margem de lucro: " & Format(margemLucro, "0.00") & "%" & vbCrLf
    resumo = resumo & "Produtos: " & lstSelecionados.ListCount & " itens"
    
    GerarResumoFinanceiro = resumo
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarResumoFinanceiro", Err)
    GerarResumoFinanceiro = "Erro ao gerar resumo"
End Function

' === VALIDAR VALORES MONETÁRIOS ===
Public Function ValidarValorMonetario(valor As String) As Boolean
    On Error GoTo TratarErro
    
    ' Remover formatação
    valor = Replace(Replace(valor, "R$", ""), " ", "")
    valor = Replace(valor, ".", "")
    valor = Replace(valor, ",", ".")
    
    ' Verificar se é numérico e positivo
    If IsNumeric(valor) Then
        ValidarValorMonetario = (CDbl(valor) >= 0)
    Else
        ValidarValorMonetario = False
    End If
    
    Exit Function
TratarErro:
    ValidarValorMonetario = False
End Function

' === ARREDONDAR VALOR ===
Public Function ArredondarValor(valor As Double, casasDecimais As Integer) As Double
    On Error GoTo TratarErro
    
    Dim multiplicador As Double
    multiplicador = 10 ^ casasDecimais
    
    ArredondarValor = Round(valor * multiplicador) / multiplicador
    
    Exit Function
TratarErro:
    ArredondarValor = valor
End Function

' === CALCULAR PORCENTAGEM ===
Public Function CalcularPorcentagem(valor As Double, total As Double) As Double
    On Error GoTo TratarErro
    
    If total = 0 Then
        CalcularPorcentagem = 0
    Else
        CalcularPorcentagem = (valor / total) * 100
    End If
    
    Exit Function
TratarErro:
    CalcularPorcentagem = 0
End Function

' === CONVERTER MOEDA PARA EXTENSO ===
Public Function ConverterMoedaParaExtenso(valor As Double) As String
    On Error GoTo TratarErro
    
    ' Função básica para converter valor para extenso
    ' Esta é uma versão simplificada
    
    Dim reais As Long
    Dim centavos As Long
    
    reais = Int(valor)
    centavos = Round((valor - reais) * 100)
    
    Dim textoReais As String
    Dim textoCentavos As String
    
    ' Converter reais
    If reais = 0 Then
        textoReais = "zero reais"
    ElseIf reais = 1 Then
        textoReais = "um real"
    Else
        textoReais = CStr(reais) & " reais"
    End If
    
    ' Converter centavos
    If centavos = 0 Then
        textoCentavos = ""
    ElseIf centavos = 1 Then
        textoCentavos = " e um centavo"
    Else
        textoCentavos = " e " & CStr(centavos) & " centavos"
    End If
    
    ConverterMoedaParaExtenso = textoReais & textoCentavos
    
    Exit Function
TratarErro:
    ConverterMoedaParaExtenso = "valor inválido"
End Function

' === CALCULAR COMISSÃO VENDEDOR ===
Public Function CalcularComissaoVendedor(valorVenda As Double, percentualComissao As Double) As Double
    On Error GoTo TratarErro
    
    If percentualComissao < 0 Or percentualComissao > 50 Then
        CalcularComissaoVendedor = 0
        Exit Function
    End If
    
    CalcularComissaoVendedor = valorVenda * (percentualComissao / 100)
    
    Exit Function
TratarErro:
    CalcularComissaoVendedor = 0
End Function

' === CALCULAR IMPOSTOS ===
Public Function CalcularImpostos(valorVenda As Double, regime As String) As Double
    On Error GoTo TratarErro
    
    Select Case UCase(regime)
        Case "SIMPLES NACIONAL"
            ' Aproximação para comércio no Simples Nacional
            CalcularImpostos = valorVenda * 0.06 ' 6%
            
        Case "LUCRO PRESUMIDO"
            CalcularImpostos = valorVenda * 0.12 ' 12%
            
        Case "LUCRO REAL"
            CalcularImpostos = valorVenda * 0.15 ' 15%
            
        Case Else
            CalcularImpostos = valorVenda * 0.06 ' Padrão Simples
    End Select
    
    Exit Function
TratarErro:
    CalcularImpostos = 0
End Function

' === ATUALIZAR TOTAIS EM TEMPO REAL ===
Public Sub AtualizarTotais(frm As Object, lstSelecionados As MSForms.ListBox)
    On Error GoTo TratarErro
    
    ' Verificar se o formulário tem os labels necessários
    Dim temLabelTotal As Boolean
    Dim temLabelDescontos As Boolean
    Dim temLabelItens As Boolean
    
    On Error Resume Next
    temLabelTotal = Not (frm.Controls("lblTotal") Is Nothing)
    temLabelDescontos = Not (frm.Controls("lblTotalDescontos") Is Nothing)
    temLabelItens = Not (frm.Controls("lblTotalItens") Is Nothing)
    On Error GoTo TratarErro
    
    ' Calcular valores
    Dim totalProdutos As Double
    Dim totalDescontos As Double
    Dim totalItens As Long
    
    totalProdutos = ProdutoManager.CalcularTotalProdutos(lstSelecionados)
    totalDescontos = DescontoManager.CalcularTotalDescontos(lstSelecionados)
    totalItens = lstSelecionados.ListCount
    
    ' Atualizar labels se existirem
    If temLabelTotal Then
        frm.Controls("lblTotal").Caption = "Total: " & Format(totalProdutos, "R$ #,##0.00")
    End If
    
    If temLabelDescontos Then
        frm.Controls("lblTotalDescontos").Caption = "Descontos: " & Format(totalDescontos, "R$ #,##0.00")
    End If
    
    If temLabelItens Then
        frm.Controls("lblTotalItens").Caption = "Itens: " & totalItens
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AtualizarTotais", Err)
End Sub

' === SIMULAR FINANCIAMENTO ===
Public Function SimularFinanciamento(valorTotal As Double, numeroParcelas As Integer, taxaJuros As Double) As String
    On Error GoTo TratarErro
    
    If numeroParcelas <= 0 Or taxaJuros < 0 Then
        SimularFinanciamento = "Parâmetros inválidos"
        Exit Function
    End If
    
    Dim valorParcela As Double
    Dim valorTotalComJuros As Double
    Dim totalJuros As Double
    
    ' Calcular usando juros compostos
    If taxaJuros = 0 Then
        valorParcela = valorTotal / numeroParcelas
        valorTotalComJuros = valorTotal
        totalJuros = 0
    Else
        Dim taxaMensal As Double
        taxaMensal = taxaJuros / 100
        
        valorParcela = valorTotal * (taxaMensal * (1 + taxaMensal) ^ numeroParcelas) / ((1 + taxaMensal) ^ numeroParcelas - 1)
        valorTotalComJuros = valorParcela * numeroParcelas
        totalJuros = valorTotalComJuros - valorTotal
    End If
    
    ' Montar simulação
    Dim simulacao As String
    simulacao = "💳 SIMULAÇÃO DE FINANCIAMENTO" & vbCrLf & vbCrLf
    simulacao = simulacao & "Valor à vista: " & Format(valorTotal, "R$ #,##0.00") & vbCrLf
    simulacao = simulacao & "Parcelas: " & numeroParcelas & "x de " & Format(valorParcela, "R$ #,##0.00") & vbCrLf
    simulacao = simulacao & "Taxa de juros: " & Format(taxaJuros, "0.00") & "% a.m." & vbCrLf
    simulacao = simulacao & "Total com juros: " & Format(valorTotalComJuros, "R$ #,##0.00") & vbCrLf
    simulacao = simulacao & "Total de juros: " & Format(totalJuros, "R$ #,##0.00") & vbCrLf
    simulacao = simulacao & "Economia à vista: " & Format(totalJuros, "R$ #,##0.00")
    
    SimularFinanciamento = simulacao
    
    Exit Function
TratarErro:
    SimularFinanciamento = "Erro na simulação"
End Function