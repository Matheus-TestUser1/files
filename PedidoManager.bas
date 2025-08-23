' ====================================================================
' MÓDULO PEDIDO MANAGER - SISTEMA PDV MADEIREIRA MARIA LUZIA
' Responsável por todas as operações relacionadas aos pedidos
' ====================================================================

Option Explicit

' === GERAR NOVO PEDIDO ===
Public Function GerarNovoPedido(dadosCliente As String, lstSelecionados As MSForms.ListBox, formaPagamento As String) As String
    On Error GoTo TratarErro
    
    ' Obter próximo número
    Dim numeroPedido As String
    numeroPedido = ObterProximoNumeroPedido()
    
    ' Criar nova planilha do pedido
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    Dim ws As Worksheet
    Set ws = CriarPlanilhaPedido(nomePlanilha)
    
    ' Preencher dados do pedido
    Call PreencherDadosPedido(ws, numeroPedido, dadosCliente, lstSelecionados, formaPagamento)
    
    ' Atualizar estatísticas
    Call AtualizarEstatisticasPedidos
    
    ' Atualizar estoque
    Call AtualizarEstoquePedido(lstSelecionados)
    
    GerarNovoPedido = numeroPedido
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarNovoPedido", Err)
    GerarNovoPedido = ""
End Function

' === OBTER PRÓXIMO NÚMERO DE PEDIDO ===
Private Function ObterProximoNumeroPedido() As String
    On Error GoTo TratarErro
    
    Dim wsControle As Worksheet
    Set wsControle = ThisWorkbook.Worksheets("Controle")
    
    Dim ultimoPedido As Long
    ultimoPedido = wsControle.Cells(2, 1).Value
    ultimoPedido = ultimoPedido + 1
    
    ' Atualizar controle
    wsControle.Cells(2, 1).Value = ultimoPedido
    
    ObterProximoNumeroPedido = Format(ultimoPedido, "00000")
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ObterProximoNumeroPedido", Err)
    ObterProximoNumeroPedido = Format(Now(), "hhmmss")
End Function

' === CRIAR PLANILHA DO PEDIDO ===
Private Function CriarPlanilhaPedido(nomePlanilha As String) As Worksheet
    On Error GoTo TratarErro
    
    ' Copiar template
    Dim wsTemplate As Worksheet
    Dim wsNovo As Worksheet
    
    Set wsTemplate = ThisWorkbook.Worksheets("Template_Pedido")
    wsTemplate.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    
    Set wsNovo = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    wsNovo.Name = nomePlanilha
    
    Set CriarPlanilhaPedido = wsNovo
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("CriarPlanilhaPedido", Err)
    Set CriarPlanilhaPedido = Nothing
End Function

' === PREENCHER DADOS DO PEDIDO ===
Private Sub PreencherDadosPedido(ws As Worksheet, numeroPedido As String, dadosCliente As String, lstSelecionados As MSForms.ListBox, formaPagamento As String)
    On Error GoTo TratarErro
    
    ' Preencher número do pedido
    ws.Range("B5").Value = "PEDIDO DE VENDA " & numeroPedido
    ws.Range("J5").Value = Format(Now(), "dd/mm/yyyy")
    
    ' Preencher dados do cliente
    If dadosCliente <> "" Then
        Dim cliente() As String
        cliente = Split(dadosCliente, "|")
        
        If UBound(cliente) >= 7 Then
            ws.Range("C7").Value = cliente(1) ' Nome
            ws.Range("C8").Value = cliente(3) ' Endereço
            ws.Range("C9").Value = cliente(2) ' CPF/CNPJ
            ws.Range("G9").Value = cliente(4) ' Cidade
            ws.Range("J9").Value = cliente(5) ' UF
            ws.Range("L9").Value = cliente(6) ' CEP
        End If
    End If
    
    ' Preencher produtos
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        If i <= 10 Then ' Máximo 11 linhas de produtos (11-21)
            Dim linha As Long
            linha = 11 + i
            
            ws.Cells(linha, 3).Value = lstSelecionados.List(i, 0) ' Referência
            ws.Cells(linha, 4).Value = lstSelecionados.List(i, 1) ' Descrição
            ws.Cells(linha, 8).Value = lstSelecionados.List(i, 2) ' Unidade
            ws.Cells(linha, 9).Value = ErrorHandler.ConverterTextoParaValor(lstSelecionados.List(i, 3)) ' Valor
            ws.Cells(linha, 10).Value = lstSelecionados.List(i, 4) ' Quantidade
            ws.Cells(linha, 11).Value = lstSelecionados.List(i, 5) ' Desconto
            ws.Cells(linha, 12).Value = ErrorHandler.ConverterTextoParaValor(lstSelecionados.List(i, 6)) ' Total
        End If
    Next i
    
    ' Calcular totais
    Dim totalProdutos As Double
    Dim totalDescontos As Double
    Dim totalFinal As Double
    
    totalProdutos = ProdutoManager.CalcularTotalProdutos(lstSelecionados)
    totalDescontos = DescontoManager.CalcularTotalDescontos(lstSelecionados)
    totalFinal = totalProdutos
    
    ' Aplicar desconto por forma de pagamento
    Dim descontoPagamento As Double
    descontoPagamento = CalculadoraManager.CalcularDescontoPagamento(totalProdutos, formaPagamento)
    totalFinal = totalFinal - descontoPagamento
    
    ' Preencher totais
    ws.Range("K22").Value = totalProdutos ' Valor produtos
    ws.Range("K23").Value = 0 ' Frete
    ws.Range("K24").Value = totalDescontos + descontoPagamento ' Valor desconto
    ws.Range("K25").Value = totalFinal ' Valor total
    
    ' Preencher informações adicionais
    ws.Range("C23").Value = Environ("USERNAME") ' Vendedor
    ws.Range("G23").Value = "BALCÃO" ' Situação
    ws.Range("C24").Value = formaPagamento ' Condições de pagamento
    ws.Range("G24").Value = Format(Now() + 1, "dd/mm/yyyy") ' Entrega
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("PreencherDadosPedido", Err)
End Sub

' === ATUALIZAR ESTOQUE APÓS PEDIDO ===
Private Sub AtualizarEstoquePedido(lstSelecionados As MSForms.ListBox)
    On Error GoTo TratarErro
    
    Dim i As Long
    For i = 0 To lstSelecionados.ListCount - 1
        Dim referencia As String
        Dim quantidade As Long
        
        referencia = lstSelecionados.List(i, 0)
        quantidade = CLng(lstSelecionados.List(i, 4))
        
        ' Dar baixa no estoque
        Call ProdutoManager.AtualizarEstoque(referencia, quantidade, "SAIDA")
    Next i
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AtualizarEstoquePedido", Err)
End Sub

' === ATUALIZAR ESTATÍSTICAS ===
Private Sub AtualizarEstatisticasPedidos()
    On Error GoTo TratarErro
    
    Dim wsControle As Worksheet
    Set wsControle = ThisWorkbook.Worksheets("Controle")
    
    ' Contar pedidos (planilhas que começam com "Pedido_")
    Dim totalPedidos As Long
    totalPedidos = 0
    
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Left(ThisWorkbook.Worksheets(i).Name, 7) = "Pedido_" Then
            totalPedidos = totalPedidos + 1
        End If
    Next i
    
    ' Atualizar controle
    wsControle.Cells(2, 1).Value = totalPedidos
    
    ' Atualizar Dashboard se existir
    On Error Resume Next
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    If Not wsDashboard Is Nothing Then
        wsDashboard.Range("B8:D10").Value = "📋 TOTAL DE PEDIDOS" & vbCrLf & vbCrLf & totalPedidos
    End If
    On Error GoTo TratarErro
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AtualizarEstatisticasPedidos", Err)
End Sub

' === LISTAR PEDIDOS ===
Public Function ListarPedidos(Optional filtroStatus As String = "") As String
    On Error GoTo TratarErro
    
    Dim resultado As String
    resultado = "LISTA DE PEDIDOS" & vbCrLf & vbCrLf
    resultado = resultado & "Nº|Cliente|Data|Valor|Status" & vbCrLf
    resultado = resultado & String(50, "-") & vbCrLf
    
    Dim totalEncontrados As Long
    totalEncontrados = 0
    
    ' Percorrer todas as planilhas de pedidos
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(i)
        
        If Left(ws.Name, 7) = "Pedido_" Then
            Dim numeroPedido As String
            Dim nomeCliente As String
            Dim dataPedido As String
            Dim valorTotal As String
            Dim status As String
            
            ' Extrair dados do pedido
            numeroPedido = Replace(ws.Name, "Pedido_", "")
            
            On Error Resume Next
            nomeCliente = ws.Range("C7").Value
            dataPedido = ws.Range("J5").Value
            valorTotal = Format(ws.Range("K25").Value, "R$ #,##0.00")
            status = IIf(ws.Range("G23").Value = "", "PENDENTE", ws.Range("G23").Value)
            On Error GoTo TratarErro
            
            ' Aplicar filtro se especificado
            If filtroStatus = "" Or UCase(status) = UCase(filtroStatus) Then
                resultado = resultado & numeroPedido & "|" & _
                           nomeCliente & "|" & _
                           dataPedido & "|" & _
                           valorTotal & "|" & _
                           status & vbCrLf
                totalEncontrados = totalEncontrados + 1
            End If
        End If
    Next i
    
    resultado = resultado & vbCrLf & "Total encontrados: " & totalEncontrados
    
    ListarPedidos = resultado
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ListarPedidos", Err)
    ListarPedidos = "Erro ao listar pedidos"
End Function

' === BUSCAR PEDIDO ===
Public Function BuscarPedido(numeroPedido As String) As String
    On Error GoTo TratarErro
    
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    ' Verificar se a planilha existe
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    On Error GoTo TratarErro
    
    If ws Is Nothing Then
        BuscarPedido = ""
        Exit Function
    End If
    
    ' Extrair dados do pedido
    Dim dadosPedido As String
    dadosPedido = numeroPedido & "|" & _
                  ws.Range("C7").Value & "|" & _
                  ws.Range("J5").Value & "|" & _
                  Format(ws.Range("K25").Value, "R$ #,##0.00") & "|" & _
                  IIf(ws.Range("G23").Value = "", "PENDENTE", ws.Range("G23").Value)
    
    BuscarPedido = dadosPedido
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("BuscarPedido", Err)
    BuscarPedido = ""
End Function

' === CANCELAR PEDIDO ===
Public Sub CancelarPedido(numeroPedido As String)
    On Error GoTo TratarErro
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("⚠️ Tem certeza que deseja cancelar o pedido " & numeroPedido & "?" & vbCrLf & _
                     "Esta ação irá:" & vbCrLf & _
                     "• Devolver produtos ao estoque" & vbCrLf & _
                     "• Marcar pedido como cancelado" & vbCrLf & _
                     "• Registrar no log do sistema", _
                     vbYesNo + vbQuestion, "Confirmar Cancelamento")
    
    If resposta = vbNo Then Exit Sub
    
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    ' Acessar planilha do pedido
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    
    ' Devolver produtos ao estoque
    Call DevolverProdutosAoEstoque(ws)
    
    ' Marcar como cancelado
    ws.Range("G23").Value = "CANCELADO"
    ws.Range("G24").Value = "Cancelado em " & Format(Now(), "dd/mm/yyyy hh:mm")
    
    ' Aplicar formatação de cancelamento
    With ws.Range("B2:L26")
        .Font.Strikethrough = True
        .Font.Color = RGB(128, 128, 128)
    End With
    
    ' Adicionar marca d'água de cancelamento
    ws.Range("F15").Value = "*** CANCELADO ***"
    ws.Range("F15").Font.Size = 24
    ws.Range("F15").Font.Bold = True
    ws.Range("F15").Font.Color = RGB(255, 0, 0)
    
    MsgBox "✅ Pedido " & numeroPedido & " cancelado com sucesso!" & vbCrLf & _
           "Produtos devolvidos ao estoque.", vbInformation, "Cancelamento Concluído"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("CancelarPedido", Err)
    MsgBox "❌ Erro ao cancelar pedido: " & Err.Description, vbCritical
End Sub

' === DEVOLVER PRODUTOS AO ESTOQUE ===
Private Sub DevolverProdutosAoEstoque(ws As Worksheet)
    On Error GoTo TratarErro
    
    ' Percorrer linhas de produtos (11-21)
    Dim i As Long
    For i = 11 To 21
        Dim referencia As String
        Dim quantidade As Long
        
        referencia = Trim(CStr(ws.Cells(i, 3).Value)) ' Coluna C
        quantidade = 0
        
        On Error Resume Next
        quantidade = CLng(ws.Cells(i, 10).Value) ' Coluna J - Quantidade
        On Error GoTo TratarErro
        
        If referencia <> "" And quantidade > 0 Then
            Call ProdutoManager.AtualizarEstoque(referencia, quantidade, "ENTRADA")
        End If
    Next i
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("DevolverProdutosAoEstoque", Err)
End Sub

' === IMPRIMIR PEDIDO ===
Public Sub ImprimirPedido(numeroPedido As String)
    On Error GoTo TratarErro
    
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    
    ' Configurar impressão
    With ws.PageSetup
        .PrintArea = "A1:M26"
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
    End With
    
    ' Visualizar antes de imprimir
    ws.PrintPreview
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ImprimirPedido", Err)
    MsgBox "❌ Erro ao imprimir pedido: " & Err.Description, vbCritical
End Sub

' === EXPORTAR PEDIDO PARA PDF ===
Public Sub ExportarPedidoParaPDF(numeroPedido As String)
    On Error GoTo TratarErro
    
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    
    ' Definir caminho do PDF
    Dim caminhoArquivo As String
    caminhoArquivo = ThisWorkbook.Path & "\Pedido_" & numeroPedido & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".pdf"
    
    ' Exportar para PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           Filename:=caminhoArquivo, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProps:=True, _
                           IgnorePrintAreas:=False, _
                           OpenAfterPublish:=False
    
    MsgBox "✅ Pedido exportado para PDF!" & vbCrLf & _
           "Arquivo: " & caminhoArquivo, vbInformation, "Exportação Concluída"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ExportarPedidoParaPDF", Err)
    MsgBox "❌ Erro ao exportar PDF: " & Err.Description, vbCritical
End Sub

' === DUPLICAR PEDIDO ===
Public Function DuplicarPedido(numeroPedidoOriginal As String) As String
    On Error GoTo TratarErro
    
    Dim nomePlanilhaOriginal As String
    nomePlanilhaOriginal = "Pedido_" & numeroPedidoOriginal
    
    ' Verificar se pedido original existe
    Dim wsOriginal As Worksheet
    Set wsOriginal = ThisWorkbook.Worksheets(nomePlanilhaOriginal)
    
    ' Gerar novo número
    Dim novoNumero As String
    novoNumero = ObterProximoNumeroPedido()
    
    ' Copiar planilha
    Dim nomePlanilhaNova As String
    nomePlanilhaNova = "Pedido_" & novoNumero
    
    wsOriginal.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    
    Dim wsNova As Worksheet
    Set wsNova = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    wsNova.Name = nomePlanilhaNova
    
    ' Atualizar dados do novo pedido
    wsNova.Range("B5").Value = "PEDIDO DE VENDA " & novoNumero
    wsNova.Range("J5").Value = Format(Now(), "dd/mm/yyyy")
    wsNova.Range("G23").Value = "DUPLICADO"
    wsNova.Range("G24").Value = "Duplicado de " & numeroPedidoOriginal
    
    ' Remover formatação de cancelamento se existir
    With wsNova.Range("B2:L26")
        .Font.Strikethrough = False
        .Font.Color = RGB(0, 0, 0)
    End With
    
    DuplicarPedido = novoNumero
    
    MsgBox "✅ Pedido duplicado com sucesso!" & vbCrLf & _
           "Pedido original: " & numeroPedidoOriginal & vbCrLf & _
           "Novo pedido: " & novoNumero, vbInformation, "Duplicação Concluída"
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("DuplicarPedido", Err)
    DuplicarPedido = ""
End Function

' === ALTERAR STATUS DO PEDIDO ===
Public Sub AlterarStatusPedido(numeroPedido As String, novoStatus As String)
    On Error GoTo TratarErro
    
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    
    ' Atualizar status
    ws.Range("G23").Value = UCase(novoStatus)
    ws.Range("G24").Value = "Atualizado em " & Format(Now(), "dd/mm/yyyy hh:mm")
    
    ' Aplicar formatação baseada no status
    Select Case UCase(novoStatus)
        Case "ENTREGUE"
            ws.Range("G23").Interior.Color = RGB(0, 176, 80)
            ws.Range("G23").Font.Color = RGB(255, 255, 255)
            
        Case "EM PRODUÇÃO"
            ws.Range("G23").Interior.Color = RGB(255, 192, 0)
            ws.Range("G23").Font.Color = RGB(0, 0, 0)
            
        Case "CANCELADO"
            ws.Range("G23").Interior.Color = RGB(255, 0, 0)
            ws.Range("G23").Font.Color = RGB(255, 255, 255)
            
        Case Else
            ws.Range("G23").Interior.Color = RGB(200, 200, 200)
            ws.Range("G23").Font.Color = RGB(0, 0, 0)
    End Select
    
    MsgBox "✅ Status do pedido " & numeroPedido & " alterado para: " & novoStatus, vbInformation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AlterarStatusPedido", Err)
    MsgBox "❌ Erro ao alterar status: " & Err.Description, vbCritical
End Sub

' === GERAR RELATÓRIO DE VENDAS ===
Public Function GerarRelatorioVendas(dataInicio As Date, dataFim As Date) As String
    On Error GoTo TratarErro
    
    Dim relatorio As String
    relatorio = "RELATÓRIO DE VENDAS" & vbCrLf
    relatorio = relatorio & "Período: " & Format(dataInicio, "dd/mm/yyyy") & " a " & Format(dataFim, "dd/mm/yyyy") & vbCrLf & vbCrLf
    
    Dim totalVendas As Double
    Dim quantidadePedidos As Long
    Dim ticketMedio As Double
    
    totalVendas = 0
    quantidadePedidos = 0
    
    relatorio = relatorio & "Nº Pedido|Cliente|Data|Valor" & vbCrLf
    relatorio = relatorio & String(50, "-") & vbCrLf
    
    ' Percorrer pedidos no período
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(i)
        
        If Left(ws.Name, 7) = "Pedido_" Then
            Dim dataPedido As Date
            Dim valorPedido As Double
            
            On Error Resume Next
            dataPedido = CDate(ws.Range("J5").Value)
            valorPedido = CDbl(ws.Range("K25").Value)
            On Error GoTo TratarErro
            
            If dataPedido >= dataInicio And dataPedido <= dataFim Then
                Dim numeroPedido As String
                Dim nomeCliente As String
                
                numeroPedido = Replace(ws.Name, "Pedido_", "")
                nomeCliente = ws.Range("C7").Value
                
                relatorio = relatorio & numeroPedido & "|" & _
                           nomeCliente & "|" & _
                           Format(dataPedido, "dd/mm/yyyy") & "|" & _
                           Format(valorPedido, "R$ #,##0.00") & vbCrLf
                
                totalVendas = totalVendas + valorPedido
                quantidadePedidos = quantidadePedidos + 1
            End If
        End If
    Next i
    
    ' Calcular ticket médio
    If quantidadePedidos > 0 Then
        ticketMedio = totalVendas / quantidadePedidos
    End If
    
    ' Resumo
    relatorio = relatorio & vbCrLf & "RESUMO:" & vbCrLf
    relatorio = relatorio & "Quantidade de pedidos: " & quantidadePedidos & vbCrLf
    relatorio = relatorio & "Total de vendas: " & Format(totalVendas, "R$ #,##0.00") & vbCrLf
    relatorio = relatorio & "Ticket médio: " & Format(ticketMedio, "R$ #,##0.00") & vbCrLf
    
    GerarRelatorioVendas = relatorio
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarRelatorioVendas", Err)
    GerarRelatorioVendas = "Erro ao gerar relatório"
End Function

' === ENVIAR PEDIDO POR EMAIL ===
Public Sub EnviarPedidoPorEmail(numeroPedido As String, emailDestino As String)
    On Error GoTo TratarErro
    
    ' Esta função requer configuração do Outlook
    ' Por enquanto, apenas simula o envio
    
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    
    Dim nomeCliente As String
    Dim valorTotal As Double
    
    nomeCliente = ws.Range("C7").Value
    valorTotal = ws.Range("K25").Value
    
    ' Simular envio
    MsgBox "📧 Simulação de envio por email:" & vbCrLf & vbCrLf & _
           "Para: " & emailDestino & vbCrLf & _
           "Assunto: Pedido " & numeroPedido & " - Madeireira Maria Luzia" & vbCrLf & _
           "Cliente: " & nomeCliente & vbCrLf & _
           "Valor: " & Format(valorTotal, "R$ #,##0.00") & vbCrLf & vbCrLf & _
           "⚠️ Funcionalidade de email requer configuração do Outlook", _
           vbInformation, "Envio por Email"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("EnviarPedidoPorEmail", Err)
    MsgBox "❌ Erro ao enviar email: " & Err.Description, vbCritical
End Sub

' === OBTER ÚLTIMOS PEDIDOS ===
Public Function ObterUltimosPedidos(Optional quantidade As Long = 5) As String
    On Error GoTo TratarErro
    
    Dim resultado As String
    resultado = ""
    
    Dim pedidosEncontrados As Long
    pedidosEncontrados = 0
    
    ' Percorrer planilhas em ordem reversa (mais recentes primeiro)
    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If pedidosEncontrados >= quantidade Then Exit For
        
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(i)
        
        If Left(ws.Name, 7) = "Pedido_" Then
            Dim numeroPedido As String
            Dim nomeCliente As String
            Dim dataPedido As String
            Dim valorTotal As String
            Dim status As String
            
            numeroPedido = Replace(ws.Name, "Pedido_", "")
            
            On Error Resume Next
            nomeCliente = ws.Range("C7").Value
            dataPedido = ws.Range("J5").Value
            valorTotal = Format(ws.Range("K25").Value, "R$ #,##0.00")
            status = IIf(ws.Range("G23").Value = "", "PENDENTE", ws.Range("G23").Value)
            On Error GoTo TratarErro
            
            resultado = resultado & numeroPedido & "|" & _
                       nomeCliente & "|" & _
                       dataPedido & "|" & _
                       valorTotal & "|" & _
                       status & vbCrLf
            
            pedidosEncontrados = pedidosEncontrados + 1
        End If
    Next i
    
    ObterUltimosPedidos = resultado
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ObterUltimosPedidos", Err)
    ObterUltimosPedidos = ""
End Function

' === CALCULAR ESTATÍSTICAS DE VENDAS ===
Public Sub CalcularEstatisticasVendas()
    On Error GoTo TratarErro
    
    Dim totalVendas As Double
    Dim quantidadePedidos As Long
    
    totalVendas = 0
    quantidadePedidos = 0
    
    ' Percorrer todos os pedidos
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(i)
        
        If Left(ws.Name, 7) = "Pedido_" Then
            Dim status As String
            Dim valorPedido As Double
            
            On Error Resume Next
            status = ws.Range("G23").Value
            valorPedido = CDbl(ws.Range("K25").Value)
            On Error GoTo TratarErro
            
            ' Contar apenas pedidos não cancelados
            If UCase(status) <> "CANCELADO" Then
                totalVendas = totalVendas + valorPedido
                quantidadePedidos = quantidadePedidos + 1
            End If
        End If
    Next i
    
    ' Atualizar controle
    Dim wsControle As Worksheet
    Set wsControle = ThisWorkbook.Worksheets("Controle")
    wsControle.Cells(2, 2).Value = totalVendas ' Total vendas
    
    ' Atualizar Dashboard
    On Error Resume Next
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    If Not wsDashboard Is Nothing Then
        wsDashboard.Range("B8:D10").Value = "📋 TOTAL DE PEDIDOS" & vbCrLf & vbCrLf & quantidadePedidos
        wsDashboard.Range("E8:G10").Value = "💰 TOTAL DE VENDAS" & vbCrLf & vbCrLf & Format(totalVendas, "R$ #,##0.00")
    End If
    On Error GoTo TratarErro
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("CalcularEstatisticasVendas", Err)
End Sub