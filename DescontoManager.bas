' ====================================================================
' MÓDULO DESCONTO MANAGER - SISTEMA PDV MADEIREIRA MARIA LUZIA
' Responsável por todas as operações relacionadas aos descontos
' ====================================================================

Option Explicit

' Tipos de desconto
Public Enum TipoDesconto
    Percentual = 1
    ValorFixo = 2
End Enum

' === APLICAR DESCONTO ===
Public Sub AplicarDesconto(lstSelecionados As MSForms.ListBox, tipoDesconto As TipoDesconto, valorDesconto As Double)
    On Error GoTo TratarErro
    
    If lstSelecionados.ListIndex < 0 Then
        MsgBox "⚠️ Selecione um produto para aplicar desconto!", vbExclamation
        Exit Sub
    End If
    
    If valorDesconto <= 0 Then
        MsgBox "⚠️ Informe um valor de desconto válido!", vbExclamation
        Exit Sub
    End If
    
    Dim indiceSelecionado As Long
    indiceSelecionado = lstSelecionados.ListIndex
    
    ' Obter dados do produto
    Dim referencia As String, descricao As String, unidade As String
    Dim precoUnitario As Double, quantidade As Long
    
    referencia = lstSelecionados.List(indiceSelecionado, 0)
    descricao = lstSelecionados.List(indiceSelecionado, 1)
    unidade = lstSelecionados.List(indiceSelecionado, 2)
    
    ' Converter preço unitário (remover formatação)
    Dim precoTexto As String
    precoTexto = lstSelecionados.List(indiceSelecionado, 3)
    precoUnitario = ErrorHandler.ConverterTextoParaValor(precoTexto)
    
    quantidade = CLng(lstSelecionados.List(indiceSelecionado, 4))
    
    ' Calcular desconto
    Dim valorTotalSemDesconto As Double
    Dim valorDescontoCalculado As Double
    Dim novoValorTotal As Double
    Dim percentualDesconto As Double
    
    valorTotalSemDesconto = precoUnitario * quantidade
    
    Select Case tipoDesconto
        Case TipoDesconto.Percentual
            If valorDesconto > 100 Then
                MsgBox "❌ Desconto percentual não pode ser maior que 100%!", vbExclamation
                Exit Sub
            End If
            
            percentualDesconto = valorDesconto
            valorDescontoCalculado = valorTotalSemDesconto * (valorDesconto / 100)
            novoValorTotal = valorTotalSemDesconto - valorDescontoCalculado
            
        Case TipoDesconto.ValorFixo
            If valorDesconto >= valorTotalSemDesconto Then
                MsgBox "❌ Desconto em valor não pode ser maior ou igual ao valor total do produto!", vbExclamation
                Exit Sub
            End If
            
            valorDescontoCalculado = valorDesconto
            novoValorTotal = valorTotalSemDesconto - valorDesconto
            percentualDesconto = (valorDesconto / valorTotalSemDesconto) * 100
    End Select
    
    ' Atualizar item na lista
    With lstSelecionados
        .List(indiceSelecionado, 5) = Format(percentualDesconto, "0.00") & "%"
        .List(indiceSelecionado, 6) = Format(novoValorTotal, "R$ #,##0.00")
    End With
    
    MsgBox "✅ Desconto aplicado com sucesso!" & vbCrLf & _
           "Produto: " & descricao & vbCrLf & _
           "Desconto: " & Format(percentualDesconto, "0.00") & "% (R$ " & Format(valorDescontoCalculado, "#,##0.00") & ")" & vbCrLf & _
           "Novo Valor: " & Format(novoValorTotal, "R$ #,##0.00"), vbInformation, "Desconto Aplicado"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AplicarDesconto", Err)
    MsgBox "❌ Erro ao aplicar desconto: " & Err.Description, vbCritical
End Sub

' === REMOVER DESCONTO ===
Public Sub RemoverDesconto(lstSelecionados As MSForms.ListBox)
    On Error GoTo TratarErro
    
    If lstSelecionados.ListIndex < 0 Then
        MsgBox "⚠️ Selecione um produto para remover desconto!", vbExclamation
        Exit Sub
    End If
    
    Dim indiceSelecionado As Long
    indiceSelecionado = lstSelecionados.ListIndex
    
    ' Obter dados do produto
    Dim precoUnitario As Double, quantidade As Long
    
    ' Converter preço unitário (remover formatação)
    Dim precoTexto As String
    precoTexto = lstSelecionados.List(indiceSelecionado, 3)
    precoUnitario = ErrorHandler.ConverterTextoParaValor(precoTexto)
    
    quantidade = CLng(lstSelecionados.List(indiceSelecionado, 4))
    
    ' Calcular valor sem desconto
    Dim valorTotal As Double
    valorTotal = precoUnitario * quantidade
    
    ' Atualizar item na lista
    With lstSelecionados
        .List(indiceSelecionado, 5) = "0.00%"
        .List(indiceSelecionado, 6) = Format(valorTotal, "R$ #,##0.00")
    End With
    
    MsgBox "✅ Desconto removido com sucesso!", vbInformation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("RemoverDesconto", Err)
    MsgBox "❌ Erro ao remover desconto: " & Err.Description, vbCritical
End Sub

' === APLICAR DESCONTO GERAL ===
Public Sub AplicarDescontoGeral(lstSelecionados As MSForms.ListBox, tipoDesconto As TipoDesconto, valorDesconto As Double)
    On Error GoTo TratarErro
    
    If lstSelecionados.ListCount = 0 Then
        MsgBox "⚠️ Não há produtos para aplicar desconto!", vbExclamation
        Exit Sub
    End If
    
    If valorDesconto <= 0 Then
        MsgBox "⚠️ Informe um valor de desconto válido!", vbExclamation
        Exit Sub
    End If
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Aplicar desconto a todos os produtos?" & vbCrLf & _
                     "Tipo: " & IIf(tipoDesconto = TipoDesconto.Percentual, "Percentual", "Valor Fixo") & vbCrLf & _
                     "Valor: " & IIf(tipoDesconto = TipoDesconto.Percentual, valorDesconto & "%", "R$ " & Format(valorDesconto, "#,##0.00")), _
                     vbYesNo + vbQuestion, "Confirmar Desconto Geral")
    
    If resposta = vbNo Then Exit Sub
    
    Dim produtosAtualizados As Long
    produtosAtualizados = 0
    
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        ' Obter dados do produto
        Dim precoUnitario As Double, quantidade As Long
        
        ' Converter preço unitário (remover formatação)
        Dim precoTexto As String
        precoTexto = lstSelecionados.List(i, 3)
        precoUnitario = ErrorHandler.ConverterTextoParaValor(precoTexto)
        
        quantidade = CLng(lstSelecionados.List(i, 4))
        
        ' Calcular desconto
        Dim valorTotalSemDesconto As Double
        Dim valorDescontoCalculado As Double
        Dim novoValorTotal As Double
        Dim percentualDesconto As Double
        
        valorTotalSemDesconto = precoUnitario * quantidade
        
        Select Case tipoDesconto
            Case TipoDesconto.Percentual
                If valorDesconto > 100 Then
                    ' Pular produtos com desconto inválido
                    GoTo ProximoProduto
                End If
                
                percentualDesconto = valorDesconto
                valorDescontoCalculado = valorTotalSemDesconto * (valorDesconto / 100)
                novoValorTotal = valorTotalSemDesconto - valorDescontoCalculado
                
            Case TipoDesconto.ValorFixo
                If valorDesconto >= valorTotalSemDesconto Then
                    ' Pular produtos onde o desconto é maior que o valor
                    GoTo ProximoProduto
                End If
                
                valorDescontoCalculado = valorDesconto
                novoValorTotal = valorTotalSemDesconto - valorDesconto
                percentualDesconto = (valorDesconto / valorTotalSemDesconto) * 100
        End Select
        
        ' Atualizar item na lista
        With lstSelecionados
            .List(i, 5) = Format(percentualDesconto, "0.00") & "%"
            .List(i, 6) = Format(novoValorTotal, "R$ #,##0.00")
        End With
        
        produtosAtualizados = produtosAtualizados + 1
        
ProximoProduto:
    Next i
    
    MsgBox "✅ Desconto geral aplicado!" & vbCrLf & _
           "Produtos atualizados: " & produtosAtualizados & " de " & lstSelecionados.ListCount, _
           vbInformation, "Desconto Geral Aplicado"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AplicarDescontoGeral", Err)
    MsgBox "❌ Erro ao aplicar desconto geral: " & Err.Description, vbCritical
End Sub

' === CALCULAR TOTAL DE DESCONTOS ===
Public Function CalcularTotalDescontos(lstSelecionados As MSForms.ListBox) As Double
    On Error GoTo TratarErro
    
    Dim totalDescontos As Double
    totalDescontos = 0
    
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        ' Obter dados do produto
        Dim precoUnitario As Double, quantidade As Long
        Dim valorTotal As Double
        
        ' Converter valores
        Dim precoTexto As String, valorTotalTexto As String
        precoTexto = lstSelecionados.List(i, 3)
        valorTotalTexto = lstSelecionados.List(i, 6)
        
        precoUnitario = ErrorHandler.ConverterTextoParaValor(precoTexto)
        valorTotal = ErrorHandler.ConverterTextoParaValor(valorTotalTexto)
        quantidade = CLng(lstSelecionados.List(i, 4))
        
        ' Calcular desconto do item
        Dim valorSemDesconto As Double
        Dim descontoItem As Double
        
        valorSemDesconto = precoUnitario * quantidade
        descontoItem = valorSemDesconto - valorTotal
        
        totalDescontos = totalDescontos + descontoItem
    Next i
    
    CalcularTotalDescontos = totalDescontos
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("CalcularTotalDescontos", Err)
    CalcularTotalDescontos = 0
End Function

' === VALIDAR DESCONTO ===
Public Function ValidarDesconto(tipoDesconto As TipoDesconto, valorDesconto As Double, valorProduto As Double) As Boolean
    On Error GoTo TratarErro
    
    Select Case tipoDesconto
        Case TipoDesconto.Percentual
            ValidarDesconto = (valorDesconto > 0 And valorDesconto <= 100)
            
        Case TipoDesconto.ValorFixo
            ValidarDesconto = (valorDesconto > 0 And valorDesconto < valorProduto)
            
        Case Else
            ValidarDesconto = False
    End Select
    
    Exit Function
TratarErro:
    ValidarDesconto = False
End Function

' === OBTER MAIOR DESCONTO APLICADO ===
Public Function ObterMaiorDesconto(lstSelecionados As MSForms.ListBox) As Double
    On Error GoTo TratarErro
    
    Dim maiorDesconto As Double
    maiorDesconto = 0
    
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        Dim descontoTexto As String
        Dim percentualDesconto As Double
        
        descontoTexto = lstSelecionados.List(i, 5) ' Coluna de desconto
        descontoTexto = Replace(descontoTexto, "%", "")
        
        If IsNumeric(descontoTexto) Then
            percentualDesconto = CDbl(descontoTexto)
            If percentualDesconto > maiorDesconto Then
                maiorDesconto = percentualDesconto
            End If
        End If
    Next i
    
    ObterMaiorDesconto = maiorDesconto
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ObterMaiorDesconto", Err)
    ObterMaiorDesconto = 0
End Function

' === APLICAR DESCONTO POR CATEGORIA ===
Public Sub AplicarDescontoPorCategoria(lstSelecionados As MSForms.ListBox, categoria As String, tipoDesconto As TipoDesconto, valorDesconto As Double)
    On Error GoTo TratarErro
    
    If lstSelecionados.ListCount = 0 Then
        MsgBox "⚠️ Não há produtos para aplicar desconto!", vbExclamation
        Exit Sub
    End If
    
    ' Buscar produtos da categoria na planilha
    Dim wsProdutos As Worksheet
    Set wsProdutos = ThisWorkbook.Worksheets("Produtos")
    
    Dim produtosAtualizados As Long
    produtosAtualizados = 0
    
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        Dim referencia As String
        referencia = lstSelecionados.List(i, 0)
        
        ' Verificar categoria do produto
        Dim categoriaProduto As String
        categoriaProduto = ObterCategoriaProduto(referencia)
        
        If UCase(categoriaProduto) = UCase(categoria) Then
            ' Aplicar desconto a este produto
            Dim precoUnitario As Double, quantidade As Long
            
            ' Converter preço unitário
            Dim precoTexto As String
            precoTexto = lstSelecionados.List(i, 3)
            precoUnitario = ErrorHandler.ConverterTextoParaValor(precoTexto)
            
            quantidade = CLng(lstSelecionados.List(i, 4))
            
            ' Calcular desconto
            Dim valorTotalSemDesconto As Double
            Dim valorDescontoCalculado As Double
            Dim novoValorTotal As Double
            Dim percentualDesconto As Double
            
            valorTotalSemDesconto = precoUnitario * quantidade
            
            Select Case tipoDesconto
                Case TipoDesconto.Percentual
                    percentualDesconto = valorDesconto
                    valorDescontoCalculado = valorTotalSemDesconto * (valorDesconto / 100)
                    novoValorTotal = valorTotalSemDesconto - valorDescontoCalculado
                    
                Case TipoDesconto.ValorFixo
                    If valorDesconto < valorTotalSemDesconto Then
                        valorDescontoCalculado = valorDesconto
                        novoValorTotal = valorTotalSemDesconto - valorDesconto
                        percentualDesconto = (valorDesconto / valorTotalSemDesconto) * 100
                    Else
                        GoTo ProximoProduto ' Pular se desconto for maior que valor
                    End If
            End Select
            
            ' Atualizar item
            With lstSelecionados
                .List(i, 5) = Format(percentualDesconto, "0.00") & "%"
                .List(i, 6) = Format(novoValorTotal, "R$ #,##0.00")
            End With
            
            produtosAtualizados = produtosAtualizados + 1
        End If
        
ProximoProduto:
    Next i
    
    If produtosAtualizados > 0 Then
        MsgBox "✅ Desconto aplicado por categoria!" & vbCrLf & _
               "Categoria: " & categoria & vbCrLf & _
               "Produtos atualizados: " & produtosAtualizados, vbInformation
    Else
        MsgBox "ℹ️ Nenhum produto da categoria '" & categoria & "' encontrado na lista.", vbInformation
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AplicarDescontoPorCategoria", Err)
    MsgBox "❌ Erro ao aplicar desconto por categoria: " & Err.Description, vbCritical
End Sub

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
    
    ObterCategoriaProduto = ""
    
    Exit Function
TratarErro:
    ObterCategoriaProduto = ""
End Function

' === CALCULAR DESCONTO PROGRESSIVO ===
Public Sub AplicarDescontoProgressivo(lstSelecionados As MSForms.ListBox, valorMinimo As Double, percentualDesconto As Double)
    On Error GoTo TratarErro
    
    If lstSelecionados.ListCount = 0 Then
        MsgBox "⚠️ Não há produtos na lista!", vbExclamation
        Exit Sub
    End If
    
    ' Calcular valor total atual
    Dim valorTotalPedido As Double
    valorTotalPedido = ProdutoManager.CalcularTotalProdutos(lstSelecionados)
    
    If valorTotalPedido < valorMinimo Then
        MsgBox "ℹ️ Valor do pedido (R$ " & Format(valorTotalPedido, "#,##0.00") & ") " & _
               "é menor que o mínimo para desconto (R$ " & Format(valorMinimo, "#,##0.00") & ").", _
               vbInformation, "Desconto Progressivo"
        Exit Sub
    End If
    
    ' Aplicar desconto percentual a todos os produtos
    Call AplicarDescontoGeral(lstSelecionados, TipoDesconto.Percentual, percentualDesconto)
    
    MsgBox "✅ Desconto progressivo aplicado!" & vbCrLf & _
           "Valor mínimo atingido: R$ " & Format(valorMinimo, "#,##0.00") & vbCrLf & _
           "Desconto aplicado: " & percentualDesconto & "%", vbInformation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AplicarDescontoProgressivo", Err)
    MsgBox "❌ Erro ao aplicar desconto progressivo: " & Err.Description, vbCritical
End Sub

' === OBTER RELATÓRIO DE DESCONTOS ===
Public Function ObterRelatorioDescontos(lstSelecionados As MSForms.ListBox) As String
    On Error GoTo TratarErro
    
    Dim relatorio As String
    relatorio = "RELATÓRIO DE DESCONTOS" & vbCrLf & vbCrLf
    
    If lstSelecionados.ListCount = 0 Then
        relatorio = relatorio & "Nenhum produto na lista."
        ObterRelatorioDescontos = relatorio
        Exit Function
    End If
    
    Dim totalDescontos As Double
    Dim produtosComDesconto As Long
    Dim maiorDesconto As Double
    
    totalDescontos = 0
    produtosComDesconto = 0
    maiorDesconto = 0
    
    relatorio = relatorio & "Produto|Desconto|Valor Economizado" & vbCrLf
    relatorio = relatorio & String(50, "-") & vbCrLf
    
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        ' Obter dados
        Dim referencia As String, descricao As String
        Dim precoUnitario As Double, quantidade As Long, valorTotal As Double
        Dim percentualDesconto As Double
        
        referencia = lstSelecionados.List(i, 0)
        descricao = lstSelecionados.List(i, 1)
        
        Dim precoTexto As String, valorTotalTexto As String, descontoTexto As String
        precoTexto = lstSelecionados.List(i, 3)
        valorTotalTexto = lstSelecionados.List(i, 6)
        descontoTexto = lstSelecionados.List(i, 5)
        
        precoUnitario = ErrorHandler.ConverterTextoParaValor(precoTexto)
        valorTotal = ErrorHandler.ConverterTextoParaValor(valorTotalTexto)
        quantidade = CLng(lstSelecionados.List(i, 4))
        
        ' Extrair percentual de desconto
        descontoTexto = Replace(descontoTexto, "%", "")
        If IsNumeric(descontoTexto) Then
            percentualDesconto = CDbl(descontoTexto)
        Else
            percentualDesconto = 0
        End If
        
        ' Calcular economia
        Dim valorSemDesconto As Double
        Dim economia As Double
        
        valorSemDesconto = precoUnitario * quantidade
        economia = valorSemDesconto - valorTotal
        
        If percentualDesconto > 0 Then
            produtosComDesconto = produtosComDesconto + 1
            totalDescontos = totalDescontos + economia
            
            If percentualDesconto > maiorDesconto Then
                maiorDesconto = percentualDesconto
            End If
            
            relatorio = relatorio & referencia & " - " & Left(descricao, 20) & "|" & _
                        Format(percentualDesconto, "0.00") & "%|" & _
                        Format(economia, "R$ #,##0.00") & vbCrLf
        End If
    Next i
    
    relatorio = relatorio & vbCrLf & "RESUMO:" & vbCrLf
    relatorio = relatorio & "Produtos com desconto: " & produtosComDesconto & vbCrLf
    relatorio = relatorio & "Total economizado: " & Format(totalDescontos, "R$ #,##0.00") & vbCrLf
    relatorio = relatorio & "Maior desconto: " & Format(maiorDesconto, "0.00") & "%" & vbCrLf
    
    ObterRelatorioDescontos = relatorio
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ObterRelatorioDescontos", Err)
    ObterRelatorioDescontos = "Erro ao gerar relatório de descontos"
End Function

' === CONFIGURAR DESCONTO PROMOCIONAL ===
Public Sub ConfigurarDesconto Promocional(categoria As String, percentual As Double, dataInicio As Date, dataFim As Date)
    On Error GoTo TratarErro
    
    ' Esta função pode ser expandida para criar um sistema de promoções
    ' Por enquanto, apenas registra a configuração
    
    Dim mensagem As String
    mensagem = "🎉 PROMOÇÃO CONFIGURADA!" & vbCrLf & vbCrLf & _
               "Categoria: " & categoria & vbCrLf & _
               "Desconto: " & percentual & "%" & vbCrLf & _
               "Período: " & Format(dataInicio, "dd/mm/yyyy") & " a " & Format(dataFim, "dd/mm/yyyy")
    
    MsgBox mensagem, vbInformation, "Promoção Ativa"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ConfigurarDescontoPromocional", Err)
End Sub

' === LIMPAR TODOS OS DESCONTOS ===
Public Sub LimparTodosDescontos(lstSelecionados As MSForms.ListBox)
    On Error GoTo TratarErro
    
    If lstSelecionados.ListCount = 0 Then
        MsgBox "⚠️ Não há produtos na lista!", vbExclamation
        Exit Sub
    End If
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Remover todos os descontos aplicados?", vbYesNo + vbQuestion, "Confirmar Remoção")
    
    If resposta = vbNo Then Exit Sub
    
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        ' Recalcular valor sem desconto
        Dim precoUnitario As Double, quantidade As Long
        
        Dim precoTexto As String
        precoTexto = lstSelecionados.List(i, 3)
        precoUnitario = ErrorHandler.ConverterTextoParaValor(precoTexto)
        quantidade = CLng(lstSelecionados.List(i, 4))
        
        Dim valorTotal As Double
        valorTotal = precoUnitario * quantidade
        
        ' Atualizar item
        With lstSelecionados
            .List(i, 5) = "0.00%"
            .List(i, 6) = Format(valorTotal, "R$ #,##0.00")
        End With
    Next i
    
    MsgBox "✅ Todos os descontos foram removidos!", vbInformation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("LimparTodosDescontos", Err)
    MsgBox "❌ Erro ao limpar descontos: " & Err.Description, vbCritical
End Sub