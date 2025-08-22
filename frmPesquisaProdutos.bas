Option Explicit
' === CÓDIGO REFINADO - VALORES MONETÁRIOS COM SISTEMA DE DESCONTO ===

Private Const QTD_MINIMA As Long = 1, QTD_MAXIMA As Long = 9999

' CORREÇÃO 1: As declarações WithEvents abaixo foram removidas por serem desnecessárias
' Private WithEvents btnMais As MSForms.CommandButton
' Private WithEvents btnMenos As MSForms.CommandButton

Private Sub btnEnviarParaProdutos_Click()
    On Error GoTo TratarErro
    
    If Me.lstSelecionados.ListCount = 0 Then
        MsgBox "?? Nenhum produto selecionado!", vbExclamation
        Exit Sub
    End If
    
    ' Calcular totais
    Dim TotalItensLocal As Long
    Dim ValorTotalLocal As Currency
    Dim i As Long
    
    TotalItensLocal = Me.lstSelecionados.ListCount
    ValorTotalLocal = 0
    
    ' Calcular valor total usando a função TextoParaMoeda existente
    For i = 0 To Me.lstSelecionados.ListCount - 1
        ValorTotalLocal = ValorTotalLocal + TextoParaMoeda(CStr(Me.lstSelecionados.List(i, 6)))
    Next i
    
    ' Obter data/hora atual e usuário
    Dim dataHoraAtual As String
    Dim usuarioAtual As String
    dataHoraAtual = Format(Now, "yyyy-mm-dd hh:mm:ss")
    usuarioAtual = "Matheus-TestUser1"
    
    ' Confirmar envio
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("?? Enviar " & TotalItensLocal & " produto(s) para frmPDVPrincipal?" & vbCrLf & vbCrLf & _
                     "?? Valor Total: " & Format(ValorTotalLocal, "R$ #,##0.00") & vbCrLf & vbCrLf & _
                     "?? Destino: frmPDVPrincipal.produtosv2" & vbCrLf & _
                     "?? Usuário: " & usuarioAtual & vbCrLf & _
                     "?? Data/Hora: " & dataHoraAtual, vbYesNo + vbQuestion, "Confirmar Envio")
    
    If resposta = vbYes Then
        ' Tentar conectar com frmPDVPrincipal
        Dim formPrincipal As Object
        Dim produtosv2 As Object
        Dim conectouSucesso As Boolean
        conectouSucesso = False
        
        ' Primeira tentativa: buscar na coleção UserForms
        Dim frm As Object
        On Error Resume Next
        For Each frm In VBA.UserForms
            If TypeName(frm) = "frmPDVPrincipal" Then
                Set formPrincipal = frm
                Exit For
            End If
        Next frm
        
        ' Segunda tentativa: acesso direto se não encontrou
        If formPrincipal Is Nothing Then
            Set formPrincipal = frmPDVPrincipal
        End If
        
        ' Tentar acessar o controle produtosv2
        If Not formPrincipal Is Nothing Then
            Set produtosv2 = formPrincipal.Controls("produtosv2")
            If produtosv2 Is Nothing Then
                ' Fallback para produtosv1 se produtosv2 não existir
                Set produtosv2 = formPrincipal.Controls("produtosv1")
            End If
            
            If Not produtosv2 Is Nothing Then
                conectouSucesso = True
            End If
        End If
        On Error GoTo TratarErro
        
        If conectouSucesso Then
            ' ? SUCESSO - Configurar e enviar dados
            Dim enviados As Long
            enviados = 0
            
            ' Configurar a listbox de destino
            On Error Resume Next
            With produtosv2
                .ColumnCount = 7
                .ColumnWidths = "60;140;30;60;40;80;60"
            End With
            On Error GoTo TratarErro
            
            ' Enviar cada item individualmente com verificação
            For i = 0 To Me.lstSelecionados.ListCount - 1
                On Error Resume Next
                
                ' Adicionar nova linha
                produtosv2.AddItem
                
                ' Verificar se a linha foi criada e popular dados
                If produtosv2.ListCount > 0 Then
                    Dim ultimaLinha As Long
                    ultimaLinha = produtosv2.ListCount - 1
                    
                    ' Popular dados coluna por coluna
                    produtosv2.List(ultimaLinha, 0) = CStr(Me.lstSelecionados.List(i, 0)) ' Referencia
                    produtosv2.List(ultimaLinha, 1) = CStr(Me.lstSelecionados.List(i, 1)) ' Descrição
                    produtosv2.List(ultimaLinha, 2) = CStr(Me.lstSelecionados.List(i, 2)) ' Uni
                    produtosv2.List(ultimaLinha, 3) = CStr(Me.lstSelecionados.List(i, 3)) ' Valor
                    produtosv2.List(ultimaLinha, 4) = CStr(Me.lstSelecionados.List(i, 4)) ' Quant.
                    produtosv2.List(ultimaLinha, 5) = CStr(Me.lstSelecionados.List(i, 5)) ' Desc.
                    produtosv2.List(ultimaLinha, 6) = CStr(Me.lstSelecionados.List(i, 6)) ' Valor Total
                    
                    enviados = enviados + 1
                End If
                
                On Error GoTo TratarErro
            Next i
            
            ' ?? SUCESSO!
            Dim nomeDestino As String
            nomeDestino = IIf(produtosv2.Name = "produtosv2", "produtosv2", "produtosv1")
            
            MsgBox "?? PRODUTOS ENVIADOS COM SUCESSO!" & vbCrLf & vbCrLf & _
                   "? Conectado com: frmPDVPrincipal" & vbCrLf & _
                   "?? Produtos enviados: " & enviados & " de " & TotalItensLocal & vbCrLf & _
                   "?? Valor total: " & Format(ValorTotalLocal, "R$ #,##0.00") & vbCrLf & _
                   "?? ListBox " & nomeDestino & " atualizada" & vbCrLf & vbCrLf & _
                   "?? " & usuarioAtual & " | ?? " & dataHoraAtual, vbInformation, "Envio Concluído!"
            
            Unload Me
    End If
    
    Exit Sub
    
TratarErro:
    MsgBox "? ERRO: " & Err.Description & vbCrLf & vbCrLf & _
           "?? " & usuarioAtual & " | ?? " & Format(Now, "yyyy-mm-dd hh:mm:ss"), vbExclamation, "Erro no Envio"
           End If
           
End Sub

' Função auxiliar para debug - adicione temporariamente se precisar testar
Private Sub DebugEnvio()
    Debug.Print "=== DEBUG ENVIO ==="
    Debug.Print "Total de itens: " & Me.lstSelecionados.ListCount
    
    Dim i As Long
    For i = 0 To Me.lstSelecionados.ListCount - 1
        Debug.Print "Item " & i & ":"
        Debug.Print "  Col 0: " & Me.lstSelecionados.List(i, 0)
        Debug.Print "  Col 1: " & Me.lstSelecionados.List(i, 1)
        Debug.Print "  Col 2: " & Me.lstSelecionados.List(i, 2)
        Debug.Print "  Col 3: " & Me.lstSelecionados.List(i, 3)
        Debug.Print "  Col 4: " & Me.lstSelecionados.List(i, 4)
        Debug.Print "  Col 5: " & Me.lstSelecionados.List(i, 5)
        Debug.Print "  Col 6: " & Me.lstSelecionados.List(i, 6)
    Next i
End Sub

Private Sub btnAplicarDesconto_Click()
    If Me.lstSelecionados.ListIndex < 0 Then
        MsgBox "Selecione um produto para aplicar desconto.", vbExclamation
        Exit Sub
    End If
    
    Dim valorDesconto As Double
    Dim tipoDesconto As String
    
    If Me.optDescontoPercentual.Value = True Then
        tipoDesconto = "percentual"
        valorDesconto = CDbl(Me.txtDesconto.Value)
        
        If valorDesconto < 0 Or valorDesconto > 100 Then
            MsgBox "Desconto percentual deve estar entre 0% e 100%.", vbExclamation
            Exit Sub
        End If
    Else
        tipoDesconto = "valor"
        valorDesconto = TextoParaMoeda(Me.txtDesconto.Value)
        
        If valorDesconto < 0 Then
            MsgBox "Valor de desconto não pode ser negativo.", vbExclamation
            Exit Sub
        End If
    End If
    
    Call AplicarDescontoItem(Me.lstSelecionados.ListIndex, valorDesconto, tipoDesconto)
    
    Me.txtDesconto.Value = ""
    Call AtualizarResumo
End Sub

Private Sub AplicarDescontoItem(indiceItem As Long, valorDesconto As Double, tipoDesconto As String)
    Dim precoUnitario As Currency
    Dim quantidade As Long
    Dim descontoAplicado As Currency
    Dim novoTotal As Currency
    
    precoUnitario = TextoParaMoeda(CStr(Me.lstSelecionados.List(indiceItem, 3)))
    quantidade = CLng(Me.lstSelecionados.List(indiceItem, 4))
    
    If tipoDesconto = "percentual" Then
        descontoAplicado = (precoUnitario * quantidade) * (valorDesconto / 100)
    Else
        descontoAplicado = valorDesconto
    End If
    
    novoTotal = (precoUnitario * quantidade) - descontoAplicado
    
    If novoTotal < 0 Then
        novoTotal = 0
        descontoAplicado = precoUnitario * quantidade
    End If
    
    Me.lstSelecionados.List(indiceItem, 5) = Format(descontoAplicado, "R$ #,##0.00")
    Me.lstSelecionados.List(indiceItem, 6) = Format(novoTotal, "R$ #,##0.00")
    
    Call AtualizarResumo
    
    If tipoDesconto = "percentual" Then
        MsgBox "Desconto de " & valorDesconto & "% aplicado!" & vbCrLf & _
               "Desconto: " & Format(descontoAplicado, "R$ #,##0.00"), vbInformation
    Else
        MsgBox "Desconto de " & Format(valorDesconto, "R$ #,##0.00") & " aplicado!", vbInformation
    End If
End Sub

Private Sub btnRemoverDesconto_Click()
    If Me.lstSelecionados.ListIndex < 0 Then
        MsgBox "Selecione um produto para remover desconto.", vbExclamation
        Exit Sub
    End If
    
    Dim indiceItem As Long
    indiceItem = Me.lstSelecionados.ListIndex
    
    Dim precoUnitario As Currency
    Dim quantidade As Long
    Dim totalSemDesconto As Currency
    
    precoUnitario = TextoParaMoeda(CStr(Me.lstSelecionados.List(indiceItem, 3)))
    quantidade = CLng(Me.lstSelecionados.List(indiceItem, 4))
    totalSemDesconto = precoUnitario * quantidade
    
    Me.lstSelecionados.List(indiceItem, 5) = "R$ 0,00"
    Me.lstSelecionados.List(indiceItem, 6) = Format(totalSemDesconto, "R$ #,##0.00")
    
    Call AtualizarResumo
    MsgBox "Desconto removido com sucesso!", vbInformation
End Sub

Private Sub AtualizarResumo()
    Dim totalFinal As Currency
    Dim totalDescontos As Currency
    Dim quantidadeItens As Long
    Dim i As Long
    
    totalFinal = 0
    totalDescontos = 0
    quantidadeItens = Me.lstSelecionados.ListCount
    
    On Error Resume Next
    For i = 0 To quantidadeItens - 1
        Dim valorTotalItem As Currency
        Dim descontoItem As Currency
        
        ' Pega o valor total da coluna 6 de cada linha
        valorTotalItem = TextoParaMoeda(CStr(Me.lstSelecionados.List(i, 6)))
        ' Pega o desconto da coluna 5 de cada linha
        descontoItem = TextoParaMoeda(CStr(Me.lstSelecionados.List(i, 5)))
        
        totalDescontos = totalDescontos + descontoItem
        totalFinal = totalFinal + valorTotalItem
    Next i
    On Error GoTo 0
    
    Me.lblTotal.caption = "Total: " & Format(totalFinal, "R$ #,##0.00")
    Me.lblTotalItens.caption = "Itens Selecionados: " & quantidadeItens
End Sub
Private Sub btnAdicionar_Click()
    If Me.lstProdutos.ListIndex < 0 Then
        MsgBox "Selecione um produto.", vbExclamation
        Exit Sub
    End If

    Dim i As Long, idProduto As String
    idProduto = Me.lstProdutos.List(Me.lstProdutos.ListIndex, 0)

    For i = 0 To Me.lstSelecionados.ListCount - 1
        If CStr(Me.lstSelecionados.List(i, 0)) = idProduto Then
            MsgBox "Produto já adicionado.", vbInformation
            Exit Sub
        End If
    Next i

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")

    Dim precoUnitario As Currency
    Dim linhaAtual As Long

    ' Encontra a linha do produto e pega o PREÇO DE VENDA da coluna G
    For linhaAtual = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If CStr(ws.Cells(linhaAtual, 1).Value) = idProduto Then
            precoUnitario = ws.Cells(linhaAtual, 7).Value ' Preço de Venda da Coluna G
            Exit For
        End If
    Next linhaAtual

    Dim qtd As Long
    qtd = CLng(Me.txtQuantidade.Value)

    Dim totalItem As Currency
    totalItem = precoUnitario * qtd

    ' Adiciona os dados como uma NOVA LINHA na ListBox visível
    With Me.lstSelecionados
        .AddItem
        .List(.ListCount - 1, 0) = idProduto
        .List(.ListCount - 1, 1) = Me.lstProdutos.List(Me.lstProdutos.ListIndex, 1)
        .List(.ListCount - 1, 2) = "UN"
        .List(.ListCount - 1, 3) = Format(precoUnitario, "R$ #,##0.00")
        .List(.ListCount - 1, 4) = qtd
        .List(.ListCount - 1, 5) = "R$ 0,00"
        .List(.ListCount - 1, 6) = Format(totalItem, "R$ #,##0.00")
    End With

    Call AtualizarResumo
    Me.txtQuantidade.Value = 1
End Sub




Private Sub btnRemover_Click()
    If Me.lstSelecionados.ListIndex >= 0 Then
        Me.lstSelecionados.RemoveItem Me.lstSelecionados.ListIndex
        Call AtualizarResumo
    Else
        MsgBox "Selecione um item para remover.", vbExclamation
    End If
End Sub

Private Sub optDescontoPercentual_Click()
    On Error Resume Next
    If Me.optDescontoPercentual.Value = True Then
        Me.lblDesconto.caption = "Desconto (%):"
        Me.txtDesconto.Value = ""
    End If
    On Error GoTo 0
End Sub

Private Sub optDescontoValor_Click()
    On Error Resume Next
    If Me.optDescontoValor.Value = True Then
        Me.lblDesconto.caption = "Desconto (R$):"
        Me.txtDesconto.Value = ""
    End If
    On Error GoTo 0
End Sub

Private Sub lblTotal_Click()

End Sub

' CORREÇÃO 2: Os eventos vazios abaixo foram removidos por não terem utilidade.
' Private Sub frameAcoes_Click()
' End Sub
'
' Private Sub frameSelecionados_Click()
' End Sub
'
' Private Sub lblTotal_Click()
' End Sub

Private Sub lstProdutos_Click()
    ' Nenhuma ação necessária ao clicar na lista de produtos de origem.
End Sub

Private Sub lstSelecionados_Click()
    ' O evento _Change já chama a atualização. Manter esta chamada é opcional, mas não prejudica.
    Call AtualizarResumo
End Sub

Private Sub lstSelecionados_Change()
    ' Evento principal para garantir que os totais estejam sempre sincronizados.
    Call AtualizarResumo
End Sub

Private Sub txtPesquisa_Change()
    Dim termoBusca As String
    termoBusca = UCase(Trim(Me.txtPesquisa.Value))
    If Len(termoBusca) >= 2 Then
        Call PesquisarProdutos(termoBusca)
    ElseIf Len(termoBusca) = 0 Then
        Call CarregarTodosProdutos
    End If
End Sub

Private Sub btnMais_Click()
    If CLng(Me.txtQuantidade.Value) < QTD_MAXIMA Then
        Me.txtQuantidade.Value = CLng(Me.txtQuantidade.Value) + 1
    End If
End Sub

Private Sub btnMenos_Click()
    If CLng(Me.txtQuantidade.Value) > QTD_MINIMA Then
        Me.txtQuantidade.Value = CLng(Me.txtQuantidade.Value) - 1
    End If
End Sub

Private Sub btnLimparPesquisa_Click()
    Me.txtPesquisa.Value = ""
    Me.txtPesquisa.SetFocus
End Sub

Private Sub CarregarTodosProdutos()
    On Error GoTo TratarErro
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Me.lstProdutos.Clear
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If Trim(CStr(ws.Cells(i, 1).Value)) <> "" Then
            With Me.lstProdutos
                .AddItem
                .List(.ListCount - 1, 0) = ws.Cells(i, 1).Value
                .List(.ListCount - 1, 1) = ws.Cells(i, 2).Value
                .List(.ListCount - 1, 2) = ws.Cells(i, 3).Value
                .List(.ListCount - 1, 3) = ws.Cells(i, 4).Value
                .List(.ListCount - 1, 4) = Format(ws.Cells(i, 6).Value, "R$ #,##0.00")
            End With
        End If
    Next i
    Exit Sub
TratarErro:
    MsgBox "Erro ao carregar produtos: " & Err.Description, vbCritical
End Sub

Private Sub PesquisarProdutos(termo As String)
    On Error GoTo TratarErro
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Produtos")
    Me.lstProdutos.Clear
    
    Dim i As Long
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If UCase(CStr(ws.Cells(i, 1).Value)) Like "*" & termo & "*" Or _
           UCase(CStr(ws.Cells(i, 2).Value)) Like "*" & termo & "*" Then
            With Me.lstProdutos
                .AddItem
                .List(.ListCount - 1, 0) = ws.Cells(i, 1).Value
                .List(.ListCount - 1, 1) = ws.Cells(i, 2).Value
                .List(.ListCount - 1, 2) = ws.Cells(i, 3).Value
                .List(.ListCount - 1, 3) = ws.Cells(i, 4).Value
                .List(.ListCount - 1, 4) = Format(ws.Cells(i, 6).Value, "R$ #,##0.00")
            End With
        End If
    Next i
    Exit Sub
TratarErro:
    MsgBox "Erro na pesquisa: " & Err.Description, vbCritical
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo TratarErro
    Me.txtQuantidade.Value = QTD_MINIMA
    Call ConfigurarListas
    
    ' CORREÇÃO 1: As linhas abaixo foram removidas por serem desnecessárias
    ' Set btnMais = Me.Controls("btnMais")
    ' Set btnMenos = Me.Controls("btnMenos")
    
    On Error Resume Next
    ' Me.optDescontoPercentual.Value = True ' Descomente se quiser que seja o padrão
    ' Me.lblDesconto.caption = "Desconto (%):" ' Descomente se quiser que seja o padrão
    On Error GoTo 0
    
    Call CarregarTodosProdutos
    Call AtualizarResumo
    Me.txtPesquisa.SetFocus
    Exit Sub
TratarErro:
    MsgBox "Erro ao iniciar: " & Err.Description, vbCritical
End Sub

Private Sub ConfigurarListas()
    With lstProdutos
        .ColumnCount = 5
        .ColumnWidths = "60;250;100;80;60"
    End With
    With lstSelecionados
        .ColumnCount = 7
        .ColumnWidths = "60;140;30;60;40;80;60"
    End With
End Sub

Private Function TextoParaMoeda(texto As String) As Currency
    On Error Resume Next
    
    Dim temp As String
    temp = Trim(CStr(texto))
    
    If temp = "" Then
        TextoParaMoeda = 0
        Exit Function
    End If
    
    temp = Replace(temp, "R$", "")
    temp = Replace(temp, " ", "")
    
    If InStr(temp, ".") > 0 And InStr(temp, ",") > 0 Then
        temp = Replace(temp, ".", "")
        temp = Replace(temp, ",", ".")
    ElseIf InStr(temp, ",") > 0 Then ' Tratamento mais genérico para a vírgula
        temp = Replace(temp, ",", ".")
    End If
    
    If IsNumeric(temp) Then
        TextoParaMoeda = CCur(temp)
    Else
        TextoParaMoeda = 0
    End If
    
    On Error GoTo 0
End Function

