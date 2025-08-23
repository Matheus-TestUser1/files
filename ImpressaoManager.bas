' ====================================================================
' MÓDULO IMPRESSÃO MANAGER - SISTEMA PDV MADEIREIRA MARIA LUZIA
' Responsável por todas as operações de impressão e exportação
' ====================================================================

Option Explicit

' === IMPRIMIR PEDIDO ===
Public Sub ImprimirPedido(numeroPedido As String)
    On Error GoTo TratarErro
    
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    ' Verificar se pedido existe
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    On Error GoTo TratarErro
    
    If ws Is Nothing Then
        MsgBox "❌ Pedido " & numeroPedido & " não encontrado!", vbExclamation
        Exit Sub
    End If
    
    ' Configurar impressão
    Call ConfigurarImpressaoPedido(ws)
    
    ' Confirmar impressão
    Dim nomeCliente As String
    Dim valorTotal As Double
    
    nomeCliente = ws.Range("C7").Value
    valorTotal = ws.Range("K25").Value
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("🖨️ Imprimir Pedido " & numeroPedido & "?" & vbCrLf & vbCrLf & _
                     "Cliente: " & nomeCliente & vbCrLf & _
                     "Valor: " & Format(valorTotal, "R$ #,##0.00") & vbCrLf & _
                     "Data: " & ws.Range("J5").Value, _
                     vbYesNo + vbQuestion, "Confirmar Impressão")
    
    If resposta = vbYes Then
        ws.PrintOut
        MsgBox "✅ Pedido " & numeroPedido & " impresso com sucesso!", vbInformation
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ImprimirPedido", Err)
    MsgBox "❌ Erro ao imprimir pedido: " & Err.Description, vbCritical
End Sub

' === CONFIGURAR IMPRESSÃO DO PEDIDO ===
Private Sub ConfigurarImpressaoPedido(ws As Worksheet)
    On Error Resume Next
    
    With ws.PageSetup
        .PrintArea = "A1:M26"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Draft = False
        .PrintQuality = 600
        .BlackAndWhite = False
    End With
    
    On Error GoTo 0
End Sub

' === EXPORTAR PARA PDF ===
Public Sub ExportarParaPDF(numeroPedido As String)
    On Error GoTo TratarErro
    
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    
    ' Definir nome do arquivo
    Dim nomeArquivo As String
    Dim caminhoCompleto As String
    
    nomeArquivo = "Pedido_" & numeroPedido & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".pdf"
    caminhoCompleto = ThisWorkbook.Path & "\" & nomeArquivo
    
    ' Configurar área de impressão
    Call ConfigurarImpressaoPedido(ws)
    
    ' Exportar para PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           Filename:=caminhoCompleto, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProps:=True, _
                           IgnorePrintAreas:=False, _
                           OpenAfterPublish:=False
    
    MsgBox "✅ Pedido exportado para PDF!" & vbCrLf & vbCrLf & _
           "📁 Arquivo: " & nomeArquivo & vbCrLf & _
           "📂 Local: " & ThisWorkbook.Path & vbCrLf & vbCrLf & _
           "Arquivo salvo com sucesso!", vbInformation, "Exportação PDF"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ExportarParaPDF", Err)
    MsgBox "❌ Erro ao exportar PDF: " & Err.Description, vbCritical
End Sub

' === VISUALIZAR IMPRESSÃO ===
Public Sub VisualizarImpressao(numeroPedido As String)
    On Error GoTo TratarErro
    
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    
    ' Configurar impressão
    Call ConfigurarImpressaoPedido(ws)
    
    ' Mostrar visualização
    ws.PrintPreview
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("VisualizarImpressao", Err)
    MsgBox "❌ Erro na visualização: " & Err.Description, vbCritical
End Sub

' === IMPRIMIR MÚLTIPLOS PEDIDOS ===
Public Sub ImprimirMultiplosPedidos(listaPedidos As String)
    On Error GoTo TratarErro
    
    If Trim(listaPedidos) = "" Then
        MsgBox "⚠️ Nenhum pedido especificado!", vbExclamation
        Exit Sub
    End If
    
    Dim pedidos() As String
    pedidos = Split(listaPedidos, ",")
    
    Dim pedidosImpressos As Long
    Dim pedidosComErro As Long
    
    pedidosImpressos = 0
    pedidosComErro = 0
    
    Dim i As Long
    For i = 0 To UBound(pedidos)
        Dim numeroPedido As String
        numeroPedido = Trim(pedidos(i))
        
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("Pedido_" & numeroPedido)
        
        If Not ws Is Nothing Then
            Call ConfigurarImpressaoPedido(ws)
            ws.PrintOut
            pedidosImpressos = pedidosImpressos + 1
        Else
            pedidosComErro = pedidosComErro + 1
        End If
        
        Set ws = Nothing
        On Error GoTo TratarErro
    Next i
    
    MsgBox "📊 Impressão em lote concluída!" & vbCrLf & vbCrLf & _
           "✅ Pedidos impressos: " & pedidosImpressos & vbCrLf & _
           "❌ Pedidos com erro: " & pedidosComErro & vbCrLf & _
           "📄 Total processados: " & (UBound(pedidos) + 1), vbInformation, "Impressão em Lote"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ImprimirMultiplosPedidos", Err)
    MsgBox "❌ Erro na impressão em lote: " & Err.Description, vbCritical
End Sub

' === CONFIGURAR IMPRESSORA ===
Public Sub ConfigurarImpressora()
    On Error GoTo TratarErro
    
    ' Mostrar diálogo de configuração da impressora
    Application.Dialogs(xlDialogPrinterSetup).Show
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ConfigurarImpressora", Err)
    MsgBox "❌ Erro ao configurar impressora: " & Err.Description, vbCritical
End Sub

' === ENVIAR POR EMAIL ===
Public Sub EnviarPorEmail(numeroPedido As String, emailDestino As String)
    On Error GoTo TratarErro
    
    ' Primeiro, exportar para PDF
    Dim nomePlanilha As String
    nomePlanilha = "Pedido_" & numeroPedido
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    
    ' Criar PDF temporário
    Dim caminhoArquivo As String
    caminhoArquivo = Environ("TEMP") & "\Pedido_" & numeroPedido & ".pdf"
    
    Call ConfigurarImpressaoPedido(ws)
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           Filename:=caminhoArquivo, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProps:=True, _
                           IgnorePrintAreas:=False, _
                           OpenAfterPublish:=False
    
    ' Tentar enviar por Outlook
    Call EnviarEmailOutlook(numeroPedido, emailDestino, caminhoArquivo)
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("EnviarPorEmail", Err)
    MsgBox "❌ Erro ao enviar email: " & Err.Description, vbCritical
End Sub

' === ENVIAR EMAIL VIA OUTLOOK ===
Private Sub EnviarEmailOutlook(numeroPedido As String, emailDestino As String, caminhoAnexo As String)
    On Error GoTo TratarErro
    
    ' Obter dados do pedido
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Pedido_" & numeroPedido)
    
    Dim nomeCliente As String
    Dim valorTotal As Double
    Dim dataPedido As String
    
    nomeCliente = ws.Range("C7").Value
    valorTotal = ws.Range("K25").Value
    dataPedido = ws.Range("J5").Value
    
    ' Criar objeto Outlook
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    ' Configurar email
    With OutlookMail
        .To = emailDestino
        .Subject = "Pedido " & numeroPedido & " - Madeireira Maria Luzia"
        .Body = "Prezado(a) " & nomeCliente & "," & vbCrLf & vbCrLf & _
                "Segue em anexo o pedido " & numeroPedido & " da Madeireira Maria Luzia." & vbCrLf & vbCrLf & _
                "Dados do pedido:" & vbCrLf & _
                "• Número: " & numeroPedido & vbCrLf & _
                "• Data: " & dataPedido & vbCrLf & _
                "• Valor Total: " & Format(valorTotal, "R$ #,##0.00") & vbCrLf & vbCrLf & _
                "Para dúvidas ou informações, entre em contato:" & vbCrLf & _
                "WhatsApp: (81) 3011-5515" & vbCrLf & _
                "Facebook: Madeireira Maria Luzia" & vbCrLf & vbCrLf & _
                "Atenciosamente," & vbCrLf & _
                "Equipe Madeireira Maria Luzia"
        
        .Attachments.Add caminhoAnexo
        .Display ' Mostrar email para revisão antes do envio
    End With
    
    MsgBox "📧 Email preparado e exibido no Outlook!" & vbCrLf & _
           "Revise e clique em 'Enviar' no Outlook.", vbInformation, "Email Preparado"
    
    Exit Sub
TratarErro:
    ' Se Outlook não estiver disponível, mostrar instruções
    MsgBox "⚠️ Não foi possível acessar o Outlook." & vbCrLf & vbCrLf & _
           "Para enviar por email:" & vbCrLf & _
           "1. Abra seu cliente de email" & vbCrLf & _
           "2. Anexe o arquivo PDF gerado" & vbCrLf & _
           "3. Envie para: " & emailDestino & vbCrLf & vbCrLf & _
           "Arquivo PDF: " & caminhoAnexo, vbInformation, "Envio Manual"
End Sub

' === IMPRIMIR RELATÓRIO ===
Public Sub ImprimirRelatorio(tipoRelatorio As String, dataInicio As Date, dataFim As Date)
    On Error GoTo TratarErro
    
    ' Criar planilha temporária para o relatório
    Dim wsRelatorio As Worksheet
    Set wsRelatorio = ThisWorkbook.Worksheets.Add
    wsRelatorio.Name = "Relatorio_Temp_" & Format(Now(), "hhmmss")
    
    ' Gerar conteúdo do relatório
    Select Case UCase(tipoRelatorio)
        Case "VENDAS"
            Call GerarRelatorioVendasParaImpressao(wsRelatorio, dataInicio, dataFim)
            
        Case "PRODUTOS"
            Call GerarRelatorioProdutosParaImpressao(wsRelatorio)
            
        Case "CLIENTES"
            Call GerarRelatorioClientesParaImpressao(wsRelatorio)
            
        Case "ESTOQUE"
            Call GerarRelatorioEstoqueParaImpressao(wsRelatorio)
            
        Case Else
            MsgBox "❌ Tipo de relatório inválido!", vbExclamation
            wsRelatorio.Delete
            Exit Sub
    End Select
    
    ' Configurar impressão do relatório
    Call ConfigurarImpressaoRelatorio(wsRelatorio)
    
    ' Visualizar antes de imprimir
    wsRelatorio.PrintPreview
    
    ' Limpar planilha temporária
    Application.DisplayAlerts = False
    wsRelatorio.Delete
    Application.DisplayAlerts = True
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ImprimirRelatorio", Err)
    MsgBox "❌ Erro ao imprimir relatório: " & Err.Description, vbCritical
End Sub

' === GERAR RELATÓRIO DE VENDAS PARA IMPRESSÃO ===
Private Sub GerarRelatorioVendasParaImpressao(ws As Worksheet, dataInicio As Date, dataFim As Date)
    On Error GoTo TratarErro
    
    ' Cabeçalho
    ws.Range("A1").Value = "RELATÓRIO DE VENDAS - MADEIREIRA MARIA LUZIA"
    ws.Range("A2").Value = "Período: " & Format(dataInicio, "dd/mm/yyyy") & " a " & Format(dataFim, "dd/mm/yyyy")
    ws.Range("A3").Value = "Gerado em: " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
    ' Formatação do cabeçalho
    With ws.Range("A1:A3")
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Range("A1").Font.Size = 16
    
    ' Cabeçalho da tabela
    ws.Range("A5").Value = "Nº Pedido"
    ws.Range("B5").Value = "Cliente"
    ws.Range("C5").Value = "Data"
    ws.Range("D5").Value = "Valor"
    ws.Range("E5").Value = "Status"
    ws.Range("F5").Value = "Forma Pagamento"
    
    With ws.Range("A5:F5")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Dados
    Dim linha As Long
    linha = 6
    
    Dim totalVendas As Double
    Dim quantidadePedidos As Long
    
    totalVendas = 0
    quantidadePedidos = 0
    
    ' Percorrer pedidos
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        Dim wsPedido As Worksheet
        Set wsPedido = ThisWorkbook.Worksheets(i)
        
        If Left(wsPedido.Name, 7) = "Pedido_" Then
            Dim dataPedido As Date
            Dim valorPedido As Double
            
            On Error Resume Next
            dataPedido = CDate(wsPedido.Range("J5").Value)
            valorPedido = CDbl(wsPedido.Range("K25").Value)
            On Error GoTo TratarErro
            
            If dataPedido >= dataInicio And dataPedido <= dataFim Then
                ws.Cells(linha, 1).Value = Replace(wsPedido.Name, "Pedido_", "")
                ws.Cells(linha, 2).Value = wsPedido.Range("C7").Value
                ws.Cells(linha, 3).Value = Format(dataPedido, "dd/mm/yyyy")
                ws.Cells(linha, 4).Value = valorPedido
                ws.Cells(linha, 5).Value = IIf(wsPedido.Range("G23").Value = "", "PENDENTE", wsPedido.Range("G23").Value)
                ws.Cells(linha, 6).Value = wsPedido.Range("C24").Value
                
                linha = linha + 1
                totalVendas = totalVendas + valorPedido
                quantidadePedidos = quantidadePedidos + 1
            End If
        End If
    Next i
    
    ' Resumo
    linha = linha + 1
    ws.Cells(linha, 1).Value = "RESUMO:"
    ws.Cells(linha, 1).Font.Bold = True
    
    linha = linha + 1
    ws.Cells(linha, 1).Value = "Quantidade de pedidos:"
    ws.Cells(linha, 2).Value = quantidadePedidos
    
    linha = linha + 1
    ws.Cells(linha, 1).Value = "Total de vendas:"
    ws.Cells(linha, 2).Value = totalVendas
    ws.Cells(linha, 2).NumberFormat = "R$ #,##0.00"
    
    linha = linha + 1
    ws.Cells(linha, 1).Value = "Ticket médio:"
    ws.Cells(linha, 2).Value = IIf(quantidadePedidos > 0, totalVendas / quantidadePedidos, 0)
    ws.Cells(linha, 2).NumberFormat = "R$ #,##0.00"
    
    ' Formatação geral
    ws.Range("D:D").NumberFormat = "R$ #,##0.00"
    ws.Columns("A:F").AutoFit
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarRelatorioVendasParaImpressao", Err)
End Sub

' === GERAR RELATÓRIO DE PRODUTOS PARA IMPRESSÃO ===
Private Sub GerarRelatorioProdutosParaImpressao(ws As Worksheet)
    On Error GoTo TratarErro
    
    ' Cabeçalho
    ws.Range("A1").Value = "RELATÓRIO DE PRODUTOS - MADEIREIRA MARIA LUZIA"
    ws.Range("A2").Value = "Gerado em: " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
    With ws.Range("A1:A2")
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Range("A1").Font.Size = 16
    
    ' Cabeçalho da tabela
    ws.Range("A4").Value = "Referência"
    ws.Range("B4").Value = "Descrição"
    ws.Range("C4").Value = "Categoria"
    ws.Range("D4").Value = "Unidade"
    ws.Range("E4").Value = "Preço Custo"
    ws.Range("F4").Value = "Preço Venda"
    ws.Range("G4").Value = "Estoque"
    ws.Range("H4").Value = "Margem %"
    
    With ws.Range("A4:H4")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Dados dos produtos
    Dim wsProdutos As Worksheet
    Set wsProdutos = ThisWorkbook.Worksheets("Produtos")
    
    Dim ultimaLinha As Long
    ultimaLinha = wsProdutos.Cells(wsProdutos.Rows.Count, "A").End(xlUp).Row
    
    Dim linha As Long
    linha = 5
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If Trim(CStr(wsProdutos.Cells(i, 2).Value)) <> "" Then
            ws.Cells(linha, 1).Value = wsProdutos.Cells(i, 1).Value ' Referência
            ws.Cells(linha, 2).Value = wsProdutos.Cells(i, 2).Value ' Descrição
            ws.Cells(linha, 3).Value = wsProdutos.Cells(i, 3).Value ' Categoria
            ws.Cells(linha, 4).Value = wsProdutos.Cells(i, 4).Value ' Unidade
            ws.Cells(linha, 5).Value = wsProdutos.Cells(i, 5).Value ' Preço Custo
            ws.Cells(linha, 6).Value = wsProdutos.Cells(i, 6).Value ' Preço Venda
            ws.Cells(linha, 7).Value = wsProdutos.Cells(i, 7).Value ' Estoque
            
            ' Calcular margem
            Dim precoCusto As Double, precoVenda As Double, margem As Double
            precoCusto = CDbl(wsProdutos.Cells(i, 5).Value)
            precoVenda = CDbl(wsProdutos.Cells(i, 6).Value)
            
            If precoCusto > 0 Then
                margem = ((precoVenda - precoCusto) / precoVenda) * 100
            Else
                margem = 0
            End If
            
            ws.Cells(linha, 8).Value = margem
            
            linha = linha + 1
        End If
    Next i
    
    ' Formatação
    ws.Range("E:F").NumberFormat = "R$ #,##0.00"
    ws.Range("H:H").NumberFormat = "0.00%"
    ws.Columns("A:H").AutoFit
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarRelatorioProdutosParaImpressao", Err)
End Sub

' === GERAR RELATÓRIO DE CLIENTES PARA IMPRESSÃO ===
Private Sub GerarRelatorioClientesParaImpressao(ws As Worksheet)
    On Error GoTo TratarErro
    
    ' Cabeçalho
    ws.Range("A1").Value = "RELATÓRIO DE CLIENTES - MADEIREIRA MARIA LUZIA"
    ws.Range("A2").Value = "Gerado em: " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
    With ws.Range("A1:A2")
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Range("A1").Font.Size = 16
    
    ' Cabeçalho da tabela
    ws.Range("A4").Value = "ID"
    ws.Range("B4").Value = "Nome/Razão Social"
    ws.Range("C4").Value = "CPF/CNPJ"
    ws.Range("D4").Value = "Cidade"
    ws.Range("E4").Value = "UF"
    ws.Range("F4").Value = "Telefone"
    
    With ws.Range("A4:F4")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Dados dos clientes
    Dim wsClientes As Worksheet
    Set wsClientes = ThisWorkbook.Worksheets("Clientes")
    
    Dim ultimaLinha As Long
    ultimaLinha = wsClientes.Cells(wsClientes.Rows.Count, "A").End(xlUp).Row
    
    Dim linha As Long
    linha = 5
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If Trim(CStr(wsClientes.Cells(i, 2).Value)) <> "" Then
            ws.Cells(linha, 1).Value = wsClientes.Cells(i, 1).Value ' ID
            ws.Cells(linha, 2).Value = wsClientes.Cells(i, 2).Value ' Nome
            ws.Cells(linha, 3).Value = wsClientes.Cells(i, 3).Value ' CPF/CNPJ
            ws.Cells(linha, 4).Value = wsClientes.Cells(i, 5).Value ' Cidade
            ws.Cells(linha, 5).Value = wsClientes.Cells(i, 6).Value ' UF
            ws.Cells(linha, 6).Value = wsClientes.Cells(i, 8).Value ' Telefone
            
            linha = linha + 1
        End If
    Next i
    
    ' Formatação
    ws.Columns("A:F").AutoFit
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarRelatorioClientesParaImpressao", Err)
End Sub

' === GERAR RELATÓRIO DE ESTOQUE PARA IMPRESSÃO ===
Private Sub GerarRelatorioEstoqueParaImpressao(ws As Worksheet)
    On Error GoTo TratarErro
    
    ' Cabeçalho
    ws.Range("A1").Value = "RELATÓRIO DE ESTOQUE - MADEIREIRA MARIA LUZIA"
    ws.Range("A2").Value = "Gerado em: " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
    With ws.Range("A1:A2")
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Range("A1").Font.Size = 16
    
    ' Cabeçalho da tabela
    ws.Range("A4").Value = "Referência"
    ws.Range("B4").Value = "Descrição"
    ws.Range("C4").Value = "Categoria"
    ws.Range("D4").Value = "Estoque"
    ws.Range("E4").Value = "Preço Venda"
    ws.Range("F4").Value = "Valor Estoque"
    ws.Range("G4").Value = "Status"
    
    With ws.Range("A4:G4")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Dados dos produtos
    Dim wsProdutos As Worksheet
    Set wsProdutos = ThisWorkbook.Worksheets("Produtos")
    
    Dim ultimaLinha As Long
    ultimaLinha = wsProdutos.Cells(wsProdutos.Rows.Count, "A").End(xlUp).Row
    
    Dim linha As Long
    linha = 5
    
    Dim valorTotalEstoque As Double
    valorTotalEstoque = 0
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If Trim(CStr(wsProdutos.Cells(i, 2).Value)) <> "" Then
            Dim estoque As Long
            Dim precoVenda As Double
            Dim valorEstoque As Double
            Dim status As String
            
            estoque = CLng(wsProdutos.Cells(i, 7).Value)
            precoVenda = CDbl(wsProdutos.Cells(i, 6).Value)
            valorEstoque = estoque * precoVenda
            
            ' Definir status do estoque
            If estoque = 0 Then
                status = "SEM ESTOQUE"
            ElseIf estoque <= 10 Then
                status = "ESTOQUE BAIXO"
            Else
                status = "OK"
            End If
            
            ws.Cells(linha, 1).Value = wsProdutos.Cells(i, 1).Value ' Referência
            ws.Cells(linha, 2).Value = wsProdutos.Cells(i, 2).Value ' Descrição
            ws.Cells(linha, 3).Value = wsProdutos.Cells(i, 3).Value ' Categoria
            ws.Cells(linha, 4).Value = estoque
            ws.Cells(linha, 5).Value = precoVenda
            ws.Cells(linha, 6).Value = valorEstoque
            ws.Cells(linha, 7).Value = status
            
            ' Colorir linha baseado no status
            Select Case status
                Case "SEM ESTOQUE"
                    ws.Range("A" & linha & ":G" & linha).Interior.Color = RGB(255, 200, 200)
                Case "ESTOQUE BAIXO"
                    ws.Range("A" & linha & ":G" & linha).Interior.Color = RGB(255, 255, 200)
            End Select
            
            valorTotalEstoque = valorTotalEstoque + valorEstoque
            linha = linha + 1
        End If
    Next i
    
    ' Resumo
    linha = linha + 1
    ws.Cells(linha, 1).Value = "VALOR TOTAL DO ESTOQUE:"
    ws.Cells(linha, 1).Font.Bold = True
    ws.Cells(linha, 6).Value = valorTotalEstoque
    ws.Cells(linha, 6).Font.Bold = True
    
    ' Formatação
    ws.Range("E:F").NumberFormat = "R$ #,##0.00"
    ws.Columns("A:G").AutoFit
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarRelatorioEstoqueParaImpressao", Err)
End Sub

' === CONFIGURAR IMPRESSÃO DE RELATÓRIO ===
Private Sub ConfigurarImpressaoRelatorio(ws As Worksheet)
    On Error Resume Next
    
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .PrintHeadings = False
        .PrintGridlines = True
        .CenterHorizontally = True
    End With
    
    On Error GoTo 0
End Sub

' === EXPORTAR RELATÓRIO PARA PDF ===
Public Sub ExportarRelatorioParaPDF(tipoRelatorio As String, dataInicio As Date, dataFim As Date)
    On Error GoTo TratarErro
    
    ' Criar planilha temporária
    Dim wsRelatorio As Worksheet
    Set wsRelatorio = ThisWorkbook.Worksheets.Add
    wsRelatorio.Name = "Relatorio_Temp_" & Format(Now(), "hhmmss")
    
    ' Gerar conteúdo
    Select Case UCase(tipoRelatorio)
        Case "VENDAS"
            Call GerarRelatorioVendasParaImpressao(wsRelatorio, dataInicio, dataFim)
        Case "PRODUTOS"
            Call GerarRelatorioProdutosParaImpressao(wsRelatorio)
        Case "CLIENTES"
            Call GerarRelatorioClientesParaImpressao(wsRelatorio)
        Case "ESTOQUE"
            Call GerarRelatorioEstoqueParaImpressao(wsRelatorio)
    End Select
    
    ' Configurar impressão
    Call ConfigurarImpressaoRelatorio(wsRelatorio)
    
    ' Exportar para PDF
    Dim nomeArquivo As String
    nomeArquivo = ThisWorkbook.Path & "\Relatorio_" & tipoRelatorio & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".pdf"
    
    wsRelatorio.ExportAsFixedFormat Type:=xlTypePDF, _
                                   Filename:=nomeArquivo, _
                                   Quality:=xlQualityStandard, _
                                   IncludeDocProps:=True, _
                                   IgnorePrintAreas:=False, _
                                   OpenAfterPublish:=False
    
    ' Limpar planilha temporária
    Application.DisplayAlerts = False
    wsRelatorio.Delete
    Application.DisplayAlerts = True
    
    MsgBox "✅ Relatório exportado para PDF!" & vbCrLf & _
           "Arquivo: " & nomeArquivo, vbInformation, "Exportação Concluída"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ExportarRelatorioParaPDF", Err)
    MsgBox "❌ Erro ao exportar relatório: " & Err.Description, vbCritical
End Sub