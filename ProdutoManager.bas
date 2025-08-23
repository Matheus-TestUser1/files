' ====================================================================
' MÓDULO PRODUTO MANAGER - SISTEMA PDV MADEIREIRA MARIA LUZIA
' Responsável por todas as operações relacionadas aos produtos
' ====================================================================

Option Explicit

' === CARREGAR PRODUTOS ===
Public Sub CarregarProdutos(lstProdutos As MSForms.ListBox, Optional filtro As String = "")
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    
    lstProdutos.Clear
    
    ' Configurar colunas do ListBox
    lstProdutos.ColumnCount = 7
    lstProdutos.ColumnWidths = "80;200;80;50;80;80;60"
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If ultimaLinha < 2 Then
        MsgBox "⚠️ Nenhum produto cadastrado!", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If Trim(CStr(ws.Cells(i, 2).Value)) <> "" Then
            ' Aplicar filtro se especificado
            If filtro = "" Or _
               InStr(1, UCase(ws.Cells(i, 1).Value), UCase(filtro)) > 0 Or _
               InStr(1, UCase(ws.Cells(i, 2).Value), UCase(filtro)) > 0 Or _
               InStr(1, UCase(ws.Cells(i, 3).Value), UCase(filtro)) > 0 Then
                
                lstProdutos.AddItem
                With lstProdutos
                    .List(.ListCount - 1, 0) = ws.Cells(i, 1).Value ' Referência
                    .List(.ListCount - 1, 1) = ws.Cells(i, 2).Value ' Descrição
                    .List(.ListCount - 1, 2) = ws.Cells(i, 3).Value ' Categoria
                    .List(.ListCount - 1, 3) = ws.Cells(i, 4).Value ' Unidade
                    .List(.ListCount - 1, 4) = Format(ws.Cells(i, 6).Value, "R$ #,##0.00") ' Preço Venda
                    .List(.ListCount - 1, 5) = ws.Cells(i, 7).Value ' Estoque
                    .List(.ListCount - 1, 6) = ws.Cells(i, 6).Value ' Preço numérico (oculto)
                End With
            End If
        End If
    Next i
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("CarregarProdutos", Err)
    MsgBox "❌ Erro ao carregar produtos: " & Err.Description, vbCritical
End Sub

' === PESQUISAR PRODUTOS ===
Public Sub PesquisarProdutos(lstProdutos As MSForms.ListBox, txtPesquisa As MSForms.TextBox)
    On Error GoTo TratarErro
    
    Dim termoPesquisa As String
    termoPesquisa = Trim(txtPesquisa.Text)
    
    If Len(termoPesquisa) < 2 Then
        ' Se menos de 2 caracteres, carregar todos
        Call CarregarProdutos(lstProdutos)
    Else
        ' Carregar com filtro
        Call CarregarProdutos(lstProdutos, termoPesquisa)
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("PesquisarProdutos", Err)
End Sub

' === OBTER DADOS DO PRODUTO SELECIONADO ===
Public Function ObterProdutoSelecionado(lstProdutos As MSForms.ListBox) As String
    On Error GoTo TratarErro
    
    If lstProdutos.ListIndex < 0 Then
        ObterProdutoSelecionado = ""
        Exit Function
    End If
    
    With lstProdutos
        ObterProdutoSelecionado = .List(.ListIndex, 0) & "|" & _    ' Referência
                                 .List(.ListIndex, 1) & "|" & _    ' Descrição
                                 .List(.ListIndex, 2) & "|" & _    ' Categoria
                                 .List(.ListIndex, 3) & "|" & _    ' Unidade
                                 .List(.ListIndex, 6) & "|" & _    ' Preço (valor numérico)
                                 .List(.ListIndex, 5)             ' Estoque
    End With
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ObterProdutoSelecionado", Err)
    ObterProdutoSelecionado = ""
End Function

' === ADICIONAR PRODUTO À LISTA DE SELECIONADOS ===
Public Sub AdicionarProdutoSelecionado(lstSelecionados As MSForms.ListBox, dadosProduto As String, quantidade As Long, Optional desconto As Double = 0)
    On Error GoTo TratarErro
    
    If dadosProduto = "" Or quantidade <= 0 Then
        MsgBox "⚠️ Selecione um produto e informe a quantidade!", vbExclamation
        Exit Sub
    End If
    
    Dim produto() As String
    produto = Split(dadosProduto, "|")
    
    If UBound(produto) < 5 Then
        MsgBox "❌ Dados do produto inválidos!", vbExclamation
        Exit Sub
    End If
    
    ' Verificar estoque
    Dim estoqueDisponivel As Long
    estoqueDisponivel = CLng(produto(5))
    
    If quantidade > estoqueDisponivel Then
        MsgBox "⚠️ Quantidade solicitada (" & quantidade & ") maior que o estoque disponível (" & estoqueDisponivel & ")!", _
               vbExclamation, "Estoque Insuficiente"
        Exit Sub
    End If
    
    ' Calcular valores
    Dim precoUnitario As Double
    Dim valorDesconto As Double
    Dim valorTotal As Double
    
    precoUnitario = CDbl(produto(4))
    valorDesconto = (precoUnitario * quantidade) * (desconto / 100)
    valorTotal = (precoUnitario * quantidade) - valorDesconto
    
    ' Configurar colunas do ListBox se necessário
    If lstSelecionados.ColumnCount = 0 Then
        lstSelecionados.ColumnCount = 7
        lstSelecionados.ColumnWidths = "80;200;50;80;60;60;100"
    End If
    
    ' Verificar se produto já existe na lista
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        If lstSelecionados.List(i, 0) = produto(0) Then ' Mesma referência
            ' Atualizar quantidade e valores
            Dim novaQuantidade As Long
            novaQuantidade = CLng(lstSelecionados.List(i, 4)) + quantidade
            
            ' Verificar estoque novamente
            If novaQuantidade > estoqueDisponivel Then
                MsgBox "⚠️ Quantidade total (" & novaQuantidade & ") maior que o estoque disponível (" & estoqueDisponivel & ")!", _
                       vbExclamation, "Estoque Insuficiente"
                Exit Sub
            End If
            
            ' Atualizar valores
            valorTotal = (precoUnitario * novaQuantidade) - ((precoUnitario * novaQuantidade) * (desconto / 100))
            
            lstSelecionados.List(i, 4) = novaQuantidade
            lstSelecionados.List(i, 5) = Format(desconto, "0.00") & "%"
            lstSelecionados.List(i, 6) = Format(valorTotal, "R$ #,##0.00")
            
            MsgBox "✅ Quantidade atualizada para " & novaQuantidade & " unidades!", vbInformation
            Exit Sub
        End If
    Next i
    
    ' Adicionar novo produto
    lstSelecionados.AddItem
    With lstSelecionados
        .List(.ListCount - 1, 0) = produto(0) ' Referência
        .List(.ListCount - 1, 1) = produto(1) ' Descrição
        .List(.ListCount - 1, 2) = produto(3) ' Unidade
        .List(.ListCount - 1, 3) = Format(precoUnitario, "R$ #,##0.00") ' Preço unitário
        .List(.ListCount - 1, 4) = quantidade ' Quantidade
        .List(.ListCount - 1, 5) = Format(desconto, "0.00") & "%" ' Desconto
        .List(.ListCount - 1, 6) = Format(valorTotal, "R$ #,##0.00") ' Valor Total
    End With
    
    MsgBox "✅ Produto adicionado com sucesso!" & vbCrLf & _
           "Produto: " & produto(1) & vbCrLf & _
           "Quantidade: " & quantidade & vbCrLf & _
           "Valor Total: " & Format(valorTotal, "R$ #,##0.00"), vbInformation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AdicionarProdutoSelecionado", Err)
    MsgBox "❌ Erro ao adicionar produto: " & Err.Description, vbCritical
End Sub

' === REMOVER PRODUTO SELECIONADO ===
Public Sub RemoverProdutoSelecionado(lstSelecionados As MSForms.ListBox)
    On Error GoTo TratarErro
    
    If lstSelecionados.ListIndex < 0 Then
        MsgBox "⚠️ Selecione um produto para remover!", vbExclamation
        Exit Sub
    End If
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Remover produto selecionado?", vbYesNo + vbQuestion, "Confirmar Remoção")
    
    If resposta = vbYes Then
        lstSelecionados.RemoveItem lstSelecionados.ListIndex
        MsgBox "✅ Produto removido!", vbInformation
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("RemoverProdutoSelecionado", Err)
End Sub

' === CALCULAR TOTAL DOS PRODUTOS ===
Public Function CalcularTotalProdutos(lstSelecionados As MSForms.ListBox) As Double
    On Error GoTo TratarErro
    
    Dim total As Double
    total = 0
    
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        Dim valorTexto As String
        valorTexto = lstSelecionados.List(i, 6) ' Valor Total
        
        ' Remover formatação monetária
        valorTexto = Replace(Replace(valorTexto, "R$", ""), " ", "")
        valorTexto = Replace(valorTexto, ".", "")
        valorTexto = Replace(valorTexto, ",", ".")
        
        total = total + CDbl(valorTexto)
    Next i
    
    CalcularTotalProdutos = total
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("CalcularTotalProdutos", Err)
    CalcularTotalProdutos = 0
End Function

' === CADASTRAR NOVO PRODUTO ===
Public Sub CadastrarNovoProduto()
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    
    ' Solicitar dados do produto
    Dim referencia As String, descricao As String, categoria As String
    Dim unidade As String, precoCusto As String, precoVenda As String, estoque As String
    
    referencia = InputBox("Digite a referência do produto:", "Cadastro de Produto")
    If referencia = "" Then Exit Sub
    
    ' Verificar se referência já existe
    If VerificarReferenciaExiste(referencia) Then
        MsgBox "❌ Referência já existe! Use uma referência diferente.", vbExclamation
        Exit Sub
    End If
    
    descricao = InputBox("Digite a descrição do produto:", "Cadastro de Produto")
    If descricao = "" Then Exit Sub
    
    categoria = InputBox("Digite a categoria (ex: Madeira, Ferragem, Tintas):", "Cadastro de Produto")
    unidade = InputBox("Digite a unidade (ex: UN, KG, M, M²):", "Cadastro de Produto")
    precoCusto = InputBox("Digite o preço de custo:", "Cadastro de Produto")
    precoVenda = InputBox("Digite o preço de venda:", "Cadastro de Produto")
    estoque = InputBox("Digite a quantidade em estoque:", "Cadastro de Produto")
    
    ' Validar valores numéricos
    If Not IsNumeric(precoCusto) Or Not IsNumeric(precoVenda) Or Not IsNumeric(estoque) Then
        MsgBox "❌ Valores de preço e estoque devem ser numéricos!", vbExclamation
        Exit Sub
    End If
    
    ' Inserir dados
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim novaLinha As Long
    novaLinha = ultimaLinha + 1
    
    With ws
        .Cells(novaLinha, 1).Value = UCase(referencia)
        .Cells(novaLinha, 2).Value = descricao
        .Cells(novaLinha, 3).Value = categoria
        .Cells(novaLinha, 4).Value = UCase(unidade)
        .Cells(novaLinha, 5).Value = CDbl(precoCusto)
        .Cells(novaLinha, 6).Value = CDbl(precoVenda)
        .Cells(novaLinha, 7).Value = CLng(estoque)
    End With
    
    ' Formatar valores
    ws.Range("E" & novaLinha & ":F" & novaLinha).NumberFormat = "R$ #,##0.00"
    
    MsgBox "✅ Produto cadastrado com sucesso!" & vbCrLf & _
           "Referência: " & referencia & vbCrLf & _
           "Descrição: " & descricao, vbInformation, "Cadastro Concluído"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("CadastrarNovoProduto", Err)
    MsgBox "❌ Erro ao cadastrar produto: " & Err.Description, vbCritical
End Sub

' === VERIFICAR SE REFERÊNCIA EXISTE ===
Private Function VerificarReferenciaExiste(referencia As String) As Boolean
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = UCase(Trim(referencia)) Then
            VerificarReferenciaExiste = True
            Exit Function
        End If
    Next i
    
    VerificarReferenciaExiste = False
    
    Exit Function
TratarErro:
    VerificarReferenciaExiste = False
End Function

' === ATUALIZAR ESTOQUE ===
Public Sub AtualizarEstoque(referencia As String, quantidade As Long, operacao As String)
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = UCase(Trim(referencia)) Then
            Dim estoqueAtual As Long
            estoqueAtual = CLng(ws.Cells(i, 7).Value)
            
            Select Case UCase(operacao)
                Case "SAIDA", "VENDA"
                    If estoqueAtual >= quantidade Then
                        ws.Cells(i, 7).Value = estoqueAtual - quantidade
                    Else
                        MsgBox "❌ Estoque insuficiente!" & vbCrLf & _
                               "Disponível: " & estoqueAtual & vbCrLf & _
                               "Solicitado: " & quantidade, vbExclamation
                        Exit Sub
                    End If
                    
                Case "ENTRADA", "COMPRA"
                    ws.Cells(i, 7).Value = estoqueAtual + quantidade
                    
                Case "AJUSTE"
                    ws.Cells(i, 7).Value = quantidade
                    
                Case Else
                    MsgBox "❌ Operação inválida! Use: SAIDA, ENTRADA ou AJUSTE", vbExclamation
                    Exit Sub
            End Select
            
            Exit Sub
        End If
    Next i
    
    MsgBox "❌ Produto não encontrado: " & referencia, vbExclamation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AtualizarEstoque", Err)
    MsgBox "❌ Erro ao atualizar estoque: " & Err.Description, vbCritical
End Sub

' === OBTER PREÇO DO PRODUTO ===
Public Function ObterPrecoProduto(referencia As String) As Double
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = UCase(Trim(referencia)) Then
            ObterPrecoProduto = CDbl(ws.Cells(i, 6).Value) ' Preço de venda
            Exit Function
        End If
    Next i
    
    ObterPrecoProduto = 0
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ObterPrecoProduto", Err)
    ObterPrecoProduto = 0
End Function

' === OBTER ESTOQUE DO PRODUTO ===
Public Function ObterEstoqueProduto(referencia As String) As Long
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = UCase(Trim(referencia)) Then
            ObterEstoqueProduto = CLng(ws.Cells(i, 7).Value)
            Exit Function
        End If
    Next i
    
    ObterEstoqueProduto = 0
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ObterEstoqueProduto", Err)
    ObterEstoqueProduto = 0
End Function

' === EDITAR PRODUTO ===
Public Sub EditarProduto(referencia As String)
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    
    ' Buscar produto
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim linhaProduto As Long
    linhaProduto = 0
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = UCase(Trim(referencia)) Then
            linhaProduto = i
            Exit For
        End If
    Next i
    
    If linhaProduto = 0 Then
        MsgBox "❌ Produto não encontrado!", vbExclamation
        Exit Sub
    End If
    
    ' Solicitar novos dados
    Dim descricao As String, categoria As String, unidade As String
    Dim precoCusto As String, precoVenda As String, estoque As String
    
    descricao = InputBox("Descrição:", "Editar Produto", ws.Cells(linhaProduto, 2).Value)
    If descricao = "" Then Exit Sub
    
    categoria = InputBox("Categoria:", "Editar Produto", ws.Cells(linhaProduto, 3).Value)
    unidade = InputBox("Unidade:", "Editar Produto", ws.Cells(linhaProduto, 4).Value)
    precoCusto = InputBox("Preço de Custo:", "Editar Produto", ws.Cells(linhaProduto, 5).Value)
    precoVenda = InputBox("Preço de Venda:", "Editar Produto", ws.Cells(linhaProduto, 6).Value)
    estoque = InputBox("Estoque:", "Editar Produto", ws.Cells(linhaProduto, 7).Value)
    
    ' Validar valores numéricos
    If Not IsNumeric(precoCusto) Or Not IsNumeric(precoVenda) Or Not IsNumeric(estoque) Then
        MsgBox "❌ Valores de preço e estoque devem ser numéricos!", vbExclamation
        Exit Sub
    End If
    
    ' Atualizar dados
    With ws
        .Cells(linhaProduto, 2).Value = descricao
        .Cells(linhaProduto, 3).Value = categoria
        .Cells(linhaProduto, 4).Value = UCase(unidade)
        .Cells(linhaProduto, 5).Value = CDbl(precoCusto)
        .Cells(linhaProduto, 6).Value = CDbl(precoVenda)
        .Cells(linhaProduto, 7).Value = CLng(estoque)
    End With
    
    MsgBox "✅ Produto atualizado com sucesso!", vbInformation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("EditarProduto", Err)
    MsgBox "❌ Erro ao editar produto: " & Err.Description, vbCritical
End Sub

' === LISTAR PRODUTOS COM ESTOQUE BAIXO ===
Public Function ListarProdutosEstoqueBaixo(Optional limiteEstoque As Long = 10) As String
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim resultado As String
    resultado = "PRODUTOS COM ESTOQUE BAIXO (≤" & limiteEstoque & "):" & vbCrLf & vbCrLf
    resultado = resultado & "Ref|Descrição|Estoque" & vbCrLf
    
    Dim encontrados As Boolean
    encontrados = False
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If CLng(ws.Cells(i, 7).Value) <= limiteEstoque Then
            resultado = resultado & ws.Cells(i, 1).Value & "|" & _
                       ws.Cells(i, 2).Value & "|" & _
                       ws.Cells(i, 7).Value & vbCrLf
            encontrados = True
        End If
    Next i
    
    If Not encontrados Then
        resultado = resultado & "Nenhum produto com estoque baixo encontrado!"
    End If
    
    ListarProdutosEstoqueBaixo = resultado
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ListarProdutosEstoqueBaixo", Err)
    ListarProdutosEstoqueBaixo = "Erro ao verificar estoque"
End Function