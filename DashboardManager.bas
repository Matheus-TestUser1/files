' ====================================================================
' MÓDULO DASHBOARD MANAGER - SISTEMA PDV MADEIREIRA MARIA LUZIA
' Responsável por gerenciar o dashboard e estatísticas
' ====================================================================

Option Explicit

' === ABRIR PDV ===
Public Sub AbrirPDV()
    On Error GoTo TratarErro
    
    ' Verificar se o formulário principal existe
    On Error Resume Next
    frmPDVMadeireiraML.Show
    On Error GoTo TratarErro
    
    ' Se não existir, tentar o formulário atual
    If Err.Number <> 0 Then
        On Error Resume Next
        frmPDVPrincipal.Show
        On Error GoTo TratarErro
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AbrirPDV", Err)
    MsgBox "❌ Erro ao abrir PDV: " & Err.Description, vbCritical
End Sub

' === ABRIR GESTÃO DE CLIENTES ===
Public Sub AbrirGestaoClientes()
    On Error GoTo TratarErro
    
    On Error Resume Next
    frmGestaoClientes.Show
    On Error GoTo TratarErro
    
    If Err.Number <> 0 Then
        ' Se não existir, abrir formulário simples de cadastro
        Call AbrirCadastroClienteSimples
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AbrirGestaoClientes", Err)
    MsgBox "❌ Erro ao abrir gestão de clientes: " & Err.Description, vbCritical
End Sub

' === ABRIR CADASTRO SIMPLES DE CLIENTE ===
Private Sub AbrirCadastroClienteSimples()
    On Error GoTo TratarErro
    
    ' Formulário simples usando InputBox
    Dim nomeCliente As String
    Dim cpfCnpj As String
    Dim endereco As String
    Dim cidade As String
    Dim telefone As String
    
    nomeCliente = InputBox("Nome/Razão Social do cliente:", "Cadastro de Cliente")
    If nomeCliente = "" Then Exit Sub
    
    cpfCnpj = InputBox("CPF/CNPJ:", "Cadastro de Cliente")
    endereco = InputBox("Endereço:", "Cadastro de Cliente")
    cidade = InputBox("Cidade:", "Cadastro de Cliente", "Abreu e Lima")
    telefone = InputBox("Telefone:", "Cadastro de Cliente")
    
    ' Cadastrar cliente
    Call ClienteManager.CadastrarNovoCliente(nomeCliente, cpfCnpj, endereco, cidade, "PE", "", telefone)
    
    MsgBox "✅ Cliente cadastrado com sucesso!", vbInformation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AbrirCadastroClienteSimples", Err)
End Sub

' === ABRIR GESTÃO DE PRODUTOS ===
Public Sub AbrirGestaoProdutos()
    On Error GoTo TratarErro
    
    On Error Resume Next
    frmPesquisaProdutos.Show
    On Error GoTo TratarErro
    
    If Err.Number <> 0 Then
        ' Se não existir, abrir a planilha de produtos
        ThisWorkbook.Worksheets("Produtos").Activate
        MsgBox "📋 Planilha de produtos aberta para edição direta.", vbInformation
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AbrirGestaoProdutos", Err)
    MsgBox "❌ Erro ao abrir gestão de produtos: " & Err.Description, vbCritical
End Sub

' === VER PEDIDOS ===
Public Sub VerPedidos()
    On Error GoTo TratarErro
    
    Dim listaPedidos As String
    listaPedidos = PedidoManager.ListarPedidos()
    
    If listaPedidos <> "" Then
        ' Criar formulário simples para exibir pedidos
        Call ExibirListaPedidos(listaPedidos)
    Else
        MsgBox "⚠️ Nenhum pedido encontrado.", vbInformation
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("VerPedidos", Err)
    MsgBox "❌ Erro ao listar pedidos: " & Err.Description, vbCritical
End Sub

' === EXIBIR LISTA DE PEDIDOS ===
Private Sub ExibirListaPedidos(listaPedidos As String)
    On Error GoTo TratarErro
    
    ' Criar planilha temporária para exibir pedidos
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "Lista_Pedidos_" & Format(Now(), "hhmmss")
    
    ' Cabeçalho
    ws.Range("A1").Value = "LISTA DE PEDIDOS - MADEIREIRA MARIA LUZIA"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    
    ws.Range("A2").Value = "Gerado em: " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
    ' Cabeçalho da tabela
    ws.Range("A4").Value = "Nº Pedido"
    ws.Range("B4").Value = "Cliente"
    ws.Range("C4").Value = "Data"
    ws.Range("D4").Value = "Valor"
    ws.Range("E4").Value = "Status"
    
    With ws.Range("A4:E4")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Processar dados
    Dim linhas() As String
    linhas = Split(listaPedidos, vbCrLf)
    
    Dim linha As Long
    linha = 5
    
    Dim i As Long
    For i = 3 To UBound(linhas) - 2 ' Pular cabeçalho e rodapé
        If Trim(linhas(i)) <> "" And InStr(linhas(i), "|") > 0 Then
            Dim dados() As String
            dados = Split(linhas(i), "|")
            
            If UBound(dados) >= 4 Then
                ws.Cells(linha, 1).Value = dados(0) ' Número
                ws.Cells(linha, 2).Value = dados(1) ' Cliente
                ws.Cells(linha, 3).Value = dados(2) ' Data
                ws.Cells(linha, 4).Value = dados(3) ' Valor
                ws.Cells(linha, 5).Value = dados(4) ' Status
                
                linha = linha + 1
            End If
        End If
    Next i
    
    ' Formatação
    ws.Range("D:D").NumberFormat = "R$ #,##0.00"
    ws.Columns("A:E").AutoFit
    
    ' Ativar planilha
    ws.Activate
    
    MsgBox "📋 Lista de pedidos exibida na nova planilha.", vbInformation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ExibirListaPedidos", Err)
End Sub

' === GERAR RELATÓRIOS ===
Public Sub GerarRelatorios()
    On Error GoTo TratarErro
    
    ' Menu de relatórios
    Dim opcao As String
    opcao = InputBox("Escolha o tipo de relatório:" & vbCrLf & vbCrLf & _
                    "1 = Vendas" & vbCrLf & _
                    "2 = Produtos" & vbCrLf & _
                    "3 = Clientes" & vbCrLf & _
                    "4 = Estoque" & vbCrLf & _
                    "5 = Financeiro", "Gerar Relatórios", "1")
    
    If opcao = "" Then Exit Sub
    
    Select Case opcao
        Case "1"
            Call GerarRelatorioVendas
        Case "2"
            Call ImpressaoManager.ImprimirRelatorio("PRODUTOS", Date - 30, Date)
        Case "3"
            Call ImpressaoManager.ImprimirRelatorio("CLIENTES", Date - 30, Date)
        Case "4"
            Call ImpressaoManager.ImprimirRelatorio("ESTOQUE", Date - 30, Date)
        Case "5"
            Call GerarRelatorioFinanceiro
        Case Else
            MsgBox "⚠️ Opção inválida!", vbExclamation
    End Select
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarRelatorios", Err)
    MsgBox "❌ Erro ao gerar relatórios: " & Err.Description, vbCritical
End Sub

' === GERAR RELATÓRIO DE VENDAS ===
Private Sub GerarRelatorioVendas()
    On Error GoTo TratarErro
    
    ' Solicitar período
    Dim dataInicio As String
    Dim dataFim As String
    
    dataInicio = InputBox("Data de início (dd/mm/yyyy):", "Relatório de Vendas", Format(Date - 30, "dd/mm/yyyy"))
    If dataInicio = "" Then Exit Sub
    
    dataFim = InputBox("Data de fim (dd/mm/yyyy):", "Relatório de Vendas", Format(Date, "dd/mm/yyyy"))
    If dataFim = "" Then Exit Sub
    
    ' Validar datas
    If Not IsDate(dataInicio) Or Not IsDate(dataFim) Then
        MsgBox "❌ Datas inválidas!", vbExclamation
        Exit Sub
    End If
    
    ' Gerar relatório
    Call ImpressaoManager.ImprimirRelatorio("VENDAS", CDate(dataInicio), CDate(dataFim))
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarRelatorioVendas", Err)
End Sub

' === GERAR RELATÓRIO FINANCEIRO ===
Private Sub GerarRelatorioFinanceiro()
    On Error GoTo TratarErro
    
    ' Calcular estatísticas financeiras
    Dim totalVendas As Double
    Dim totalPedidos As Long
    Dim ticketMedio As Double
    Dim margemMedia As Double
    
    totalVendas = 0
    totalPedidos = 0
    margemMedia = 0
    
    ' Percorrer todos os pedidos
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Left(ThisWorkbook.Worksheets(i).Name, 7) = "Pedido_" Then
            On Error Resume Next
            Dim valorPedido As Double
            valorPedido = CDbl(ThisWorkbook.Worksheets(i).Range("K25").Value)
            
            If valorPedido > 0 Then
                totalVendas = totalVendas + valorPedido
                totalPedidos = totalPedidos + 1
            End If
            On Error GoTo TratarErro
        End If
    Next i
    
    If totalPedidos > 0 Then
        ticketMedio = totalVendas / totalPedidos
    End If
    
    ' Exibir relatório
    Dim relatorio As String
    relatorio = "💰 RELATÓRIO FINANCEIRO" & vbCrLf & vbCrLf
    relatorio = relatorio & "📊 RESUMO GERAL:" & vbCrLf
    relatorio = relatorio & "Total de vendas: " & Format(totalVendas, "R$ #,##0.00") & vbCrLf
    relatorio = relatorio & "Quantidade de pedidos: " & totalPedidos & vbCrLf
    relatorio = relatorio & "Ticket médio: " & Format(ticketMedio, "R$ #,##0.00") & vbCrLf & vbCrLf
    relatorio = relatorio & "📈 PERFORMANCE:" & vbCrLf
    relatorio = relatorio & "Vendas hoje: " & Format(CalcularVendasDia(Date), "R$ #,##0.00") & vbCrLf
    relatorio = relatorio & "Vendas esta semana: " & Format(CalcularVendasSemana(), "R$ #,##0.00") & vbCrLf
    relatorio = relatorio & "Vendas este mês: " & Format(CalcularVendasMes(), "R$ #,##0.00")
    
    MsgBox relatorio, vbInformation, "Relatório Financeiro"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarRelatorioFinanceiro", Err)
End Sub

' === CALCULAR VENDAS DO DIA ===
Private Function CalcularVendasDia(data As Date) As Double
    On Error GoTo TratarErro
    
    Dim total As Double
    total = 0
    
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Left(ThisWorkbook.Worksheets(i).Name, 7) = "Pedido_" Then
            On Error Resume Next
            Dim dataPedido As Date
            dataPedido = CDate(ThisWorkbook.Worksheets(i).Range("J5").Value)
            
            If DateValue(dataPedido) = DateValue(data) Then
                total = total + CDbl(ThisWorkbook.Worksheets(i).Range("K25").Value)
            End If
            On Error GoTo TratarErro
        End If
    Next i
    
    CalcularVendasDia = total
    
    Exit Function
TratarErro:
    CalcularVendasDia = 0
End Function

' === CALCULAR VENDAS DA SEMANA ===
Private Function CalcularVendasSemana() As Double
    On Error GoTo TratarErro
    
    Dim inicioSemana As Date
    Dim fimSemana As Date
    
    ' Calcular início e fim da semana
    inicioSemana = Date - Weekday(Date) + 1
    fimSemana = inicioSemana + 6
    
    Dim total As Double
    total = 0
    
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Left(ThisWorkbook.Worksheets(i).Name, 7) = "Pedido_" Then
            On Error Resume Next
            Dim dataPedido As Date
            dataPedido = CDate(ThisWorkbook.Worksheets(i).Range("J5").Value)
            
            If dataPedido >= inicioSemana And dataPedido <= fimSemana Then
                total = total + CDbl(ThisWorkbook.Worksheets(i).Range("K25").Value)
            End If
            On Error GoTo TratarErro
        End If
    Next i
    
    CalcularVendasSemana = total
    
    Exit Function
TratarErro:
    CalcularVendasSemana = 0
End Function

' === CALCULAR VENDAS DO MÊS ===
Private Function CalcularVendasMes() As Double
    On Error GoTo TratarErro
    
    Dim inicioMes As Date
    Dim fimMes As Date
    
    inicioMes = DateSerial(Year(Date), Month(Date), 1)
    fimMes = DateSerial(Year(Date), Month(Date) + 1, 0)
    
    Dim total As Double
    total = 0
    
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Left(ThisWorkbook.Worksheets(i).Name, 7) = "Pedido_" Then
            On Error Resume Next
            Dim dataPedido As Date
            dataPedido = CDate(ThisWorkbook.Worksheets(i).Range("J5").Value)
            
            If dataPedido >= inicioMes And dataPedido <= fimMes Then
                total = total + CDbl(ThisWorkbook.Worksheets(i).Range("K25").Value)
            End If
            On Error GoTo TratarErro
        End If
    Next i
    
    CalcularVendasMes = total
    
    Exit Function
TratarErro:
    CalcularVendasMes = 0
End Function

' === ABRIR CONFIGURAÇÕES ===
Public Sub AbrirConfiguracoes()
    On Error GoTo TratarErro
    
    ' Menu de configurações
    Dim opcao As String
    opcao = InputBox("Escolha uma opção de configuração:" & vbCrLf & vbCrLf & _
                    "1 = Backup do sistema" & vbCrLf & _
                    "2 = Verificar integridade" & vbCrLf & _
                    "3 = Exportar log de erros" & vbCrLf & _
                    "4 = Limpar dados temporários" & vbCrLf & _
                    "5 = Informações do sistema" & vbCrLf & _
                    "6 = Resetar sistema", "Configurações", "1")
    
    If opcao = "" Then Exit Sub
    
    Select Case opcao
        Case "1"
            Call UtilsManager.CriarBackup
        Case "2"
            Dim relatorio As String
            relatorio = UtilsManager.VerificarIntegridadeSistema()
            MsgBox relatorio, vbInformation, "Verificação de Integridade"
        Case "3"
            Call UtilsManager.ExportarLogErros
        Case "4"
            Call UtilsManager.LimparDadosTemporarios
        Case "5"
            Dim info As String
            info = UtilsManager.ObterInformacoesSistema()
            MsgBox info, vbInformation, "Informações do Sistema"
        Case "6"
            Call UtilsManager.ResetarSistema
        Case Else
            MsgBox "⚠️ Opção inválida!", vbExclamation
    End Select
    
    ' Atualizar dashboard após configurações
    Call AtualizarEstatisticasDashboard
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AbrirConfiguracoes", Err)
    MsgBox "❌ Erro nas configurações: " & Err.Description, vbCritical
End Sub

' === ATUALIZAR ESTATÍSTICAS DO DASHBOARD ===
Public Sub AtualizarEstatisticasDashboard()
    On Error GoTo TratarErro
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    
    ' Calcular estatísticas
    Dim totalPedidos As Long
    Dim totalVendas As Double
    Dim totalClientes As Long
    
    ' Contar pedidos e calcular vendas
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Left(ThisWorkbook.Worksheets(i).Name, 7) = "Pedido_" Then
            totalPedidos = totalPedidos + 1
            
            On Error Resume Next
            Dim status As String
            Dim valorPedido As Double
            
            status = ThisWorkbook.Worksheets(i).Range("G23").Value
            valorPedido = CDbl(ThisWorkbook.Worksheets(i).Range("K25").Value)
            
            ' Somar apenas pedidos não cancelados
            If UCase(status) <> "CANCELADO" Then
                totalVendas = totalVendas + valorPedido
            End If
            On Error GoTo TratarErro
        End If
    Next i
    
    ' Contar clientes
    On Error Resume Next
    Dim wsClientes As Worksheet
    Set wsClientes = ThisWorkbook.Worksheets("Clientes")
    If Not wsClientes Is Nothing Then
        totalClientes = wsClientes.Cells(wsClientes.Rows.Count, "A").End(xlUp).Row - 1
        If totalClientes < 0 Then totalClientes = 0
    End If
    On Error GoTo TratarErro
    
    ' Atualizar cards do Dashboard
    wsDashboard.Range("B6:D8").Value = "📋 Total de Pedidos" & vbCrLf & vbCrLf & totalPedidos
    wsDashboard.Range("E6:G8").Value = "💰 Total de Vendas" & vbCrLf & vbCrLf & Format(totalVendas, "R$ #,##0.00")
    wsDashboard.Range("H6:J8").Value = "👥 Total de Clientes" & vbCrLf & vbCrLf & totalClientes
    
    ' Atualizar últimos pedidos
    Call AtualizarUltimosPedidos(wsDashboard)
    
    ' Atualizar data da última atualização
    wsDashboard.Range("B30").Value = "Última atualização: " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & " - Sistema PDV Madeireira Maria Luzia"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AtualizarEstatisticasDashboard", Err)
End Sub

' === ATUALIZAR ÚLTIMOS PEDIDOS ===
Private Sub AtualizarUltimosPedidos(wsDashboard As Worksheet)
    On Error GoTo TratarErro
    
    ' Limpar área dos últimos pedidos
    wsDashboard.Range("B23:J27").ClearContents
    
    ' Obter últimos pedidos
    Dim ultimosPedidos As String
    ultimosPedidos = PedidoManager.ObterUltimosPedidos(5)
    
    If ultimosPedidos <> "" Then
        Dim pedidos() As String
        pedidos = Split(ultimosPedidos, vbCrLf)
        
        Dim linha As Long
        linha = 23
        
        Dim i As Long
        For i = 0 To UBound(pedidos) - 1
            If i < 5 And Trim(pedidos(i)) <> "" Then
                Dim dadosPedido() As String
                dadosPedido = Split(pedidos(i), "|")
                
                If UBound(dadosPedido) >= 4 Then
                    wsDashboard.Cells(linha, 2).Value = dadosPedido(0) ' Número
                    wsDashboard.Cells(linha, 3).Value = Left(dadosPedido(1), 20) ' Cliente (limitado)
                    wsDashboard.Cells(linha, 5).Value = dadosPedido(2) ' Data
                    wsDashboard.Cells(linha, 6).Value = dadosPedido(3) ' Valor
                    wsDashboard.Cells(linha, 7).Value = dadosPedido(4) ' Status
                    
                    ' Botão de ações (simulado com texto)
                    wsDashboard.Cells(linha, 8).Value = "Ver/Imprimir"
                    
                    linha = linha + 1
                End If
            End If
        Next i
    Else
        wsDashboard.Range("B23").Value = "Nenhum pedido encontrado"
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AtualizarUltimosPedidos", Err)
End Sub

' === CONFIGURAR BOTÕES DO DASHBOARD ===
Public Sub ConfigurarBotoesDashboard()
    On Error GoTo TratarErro
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    
    ' Criar botões usando formas (shapes) se não existirem
    Call CriarBotaoDashboard(wsDashboard, "B12:D14", "ABRIR PDV", "AbrirPDV")
    Call CriarBotaoDashboard(wsDashboard, "E12:G14", "CADASTRAR CLIENTE", "AbrirGestaoClientes")
    Call CriarBotaoDashboard(wsDashboard, "H12:J14", "CADASTRAR PRODUTO", "AbrirGestaoProdutos")
    Call CriarBotaoDashboard(wsDashboard, "B16:D18", "VER PEDIDOS", "VerPedidos")
    Call CriarBotaoDashboard(wsDashboard, "E16:G18", "RELATÓRIOS", "GerarRelatorios")
    Call CriarBotaoDashboard(wsDashboard, "H16:J18", "CONFIGURAÇÕES", "AbrirConfiguracoes")
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ConfigurarBotoesDashboard", Err)
End Sub

' === CRIAR BOTÃO DO DASHBOARD ===
Private Sub CriarBotaoDashboard(ws As Worksheet, intervalo As String, texto As String, macro As String)
    On Error Resume Next
    
    ' Remover botão existente se houver
    Dim shape As Shape
    For Each shape In ws.Shapes
        If shape.Name = "btn" & Replace(texto, " ", "") Then
            shape.Delete
            Exit For
        End If
    Next shape
    
    ' Criar novo botão
    Set shape = ws.Shapes.AddShape(msoShapeRectangle, ws.Range(intervalo).Left, ws.Range(intervalo).Top, _
                                  ws.Range(intervalo).Width, ws.Range(intervalo).Height)
    
    With shape
        .Name = "btn" & Replace(texto, " ", "")
        .TextFrame.Characters.Text = texto
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.ColorIndex = 2 ' Branco
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        
        ' Cores baseadas no tipo de botão
        Select Case texto
            Case "ABRIR PDV"
                .Fill.ForeColor.RGB = RGB(0, 176, 80)
            Case "CADASTRAR CLIENTE"
                .Fill.ForeColor.RGB = RGB(0, 112, 192)
            Case "CADASTRAR PRODUTO"
                .Fill.ForeColor.RGB = RGB(112, 48, 160)
            Case "VER PEDIDOS"
                .Fill.ForeColor.RGB = RGB(255, 192, 0)
                .TextFrame.Characters.Font.ColorIndex = 1 ' Preto
            Case "RELATÓRIOS"
                .Fill.ForeColor.RGB = RGB(192, 0, 0)
            Case "CONFIGURAÇÕES"
                .Fill.ForeColor.RGB = RGB(128, 128, 128)
        End Select
        
        .OnAction = "DashboardManager." & macro
    End With
    
    On Error GoTo 0
End Sub

' === ABRIR DASHBOARD ===
Public Sub AbrirDashboard()
    On Error GoTo TratarErro
    
    ' Atualizar estatísticas
    Call AtualizarEstatisticasDashboard
    
    ' Ativar planilha do Dashboard
    ThisWorkbook.Worksheets("Dashboard").Activate
    
    ' Ajustar zoom
    Application.ActiveWindow.Zoom = 85
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AbrirDashboard", Err)
End Sub

' === CRIAR GRÁFICO DE VENDAS ===
Public Sub CriarGraficoVendas()
    On Error GoTo TratarErro
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    
    ' Dados para o gráfico (últimos 7 dias)
    Dim dados(1 To 7, 1 To 2) As Variant
    
    Dim i As Long
    For i = 1 To 7
        Dim dataAtual As Date
        dataAtual = Date - (7 - i)
        
        dados(i, 1) = Format(dataAtual, "dd/mm")
        dados(i, 2) = CalcularVendasDia(dataAtual)
    Next i
    
    ' Criar área de dados temporária
    Dim rangeTemp As Range
    Set rangeTemp = wsDashboard.Range("L2:M8")
    
    ' Preencher dados
    For i = 1 To 7
        rangeTemp.Cells(i, 1).Value = dados(i, 1)
        rangeTemp.Cells(i, 2).Value = dados(i, 2)
    Next i
    
    ' Remover gráfico existente
    Dim chart As ChartObject
    For Each chart In wsDashboard.ChartObjects
        If chart.Name = "GraficoVendas" Then
            chart.Delete
            Exit For
        End If
    Next chart
    
    ' Criar novo gráfico
    Set chart = wsDashboard.ChartObjects.Add(Left:=wsDashboard.Range("L10").Left, _
                                            Top:=wsDashboard.Range("L10").Top, _
                                            Width:=300, Height:=200)
    
    With chart
        .Name = "GraficoVendas"
        .Chart.SetSourceData rangeTemp
        .Chart.ChartType = xlColumnClustered
        .Chart.HasTitle = True
        .Chart.ChartTitle.Text = "Vendas dos Últimos 7 Dias"
        .Chart.HasLegend = False
    End With
    
    ' Limpar dados temporários
    rangeTemp.ClearContents
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("CriarGraficoVendas", Err)
End Sub

' === MONITORAR ESTOQUE BAIXO ===
Public Sub MonitorarEstoqueBaixo()
    On Error GoTo TratarErro
    
    Dim produtosBaixoEstoque As String
    produtosBaixoEstoque = ProdutoManager.ListarProdutosEstoqueBaixo(10) ' Limite de 10 unidades
    
    If produtosBaixoEstoque <> "" Then
        ' Exibir alerta no Dashboard
        Dim wsDashboard As Worksheet
        Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
        
        ' Criar área de alerta
        wsDashboard.Range("L12:T18").Interior.Color = RGB(255, 200, 200)
        wsDashboard.Range("L12").Value = "⚠️ ALERTA DE ESTOQUE BAIXO"
        wsDashboard.Range("L12").Font.Bold = True
        wsDashboard.Range("L12").Font.Color = RGB(255, 0, 0)
        
        wsDashboard.Range("L13").Value = "Produtos com estoque baixo:"
        wsDashboard.Range("L14").Value = produtosBaixoEstoque
        
        ' Mostrar notificação
        MsgBox "⚠️ ATENÇÃO: Produtos com estoque baixo!" & vbCrLf & vbCrLf & _
               produtosBaixoEstoque & vbCrLf & vbCrLf & _
               "Verifique o estoque e faça reposição.", vbExclamation, "Alerta de Estoque"
    Else
        ' Limpar área de alerta se não há problemas
        Dim wsDashboard As Worksheet
        Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
        
        wsDashboard.Range("L12:T18").Interior.Color = RGB(240, 240, 240)
        wsDashboard.Range("L12").Value = "✅ Estoque OK"
        wsDashboard.Range("L12").Font.Color = RGB(0, 176, 80)
        wsDashboard.Range("L13:L18").ClearContents
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("MonitorarEstoqueBaixo", Err)
End Sub

' === GERAR RESUMO DIÁRIO ===
Public Sub GerarResumoDiario()
    On Error GoTo TratarErro
    
    Dim vendasHoje As Double
    Dim pedidosHoje As Long
    Dim ticketMedioHoje As Double
    
    vendasHoje = CalcularVendasDia(Date)
    pedidosHoje = 0
    
    ' Contar pedidos de hoje
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Left(ThisWorkbook.Worksheets(i).Name, 7) = "Pedido_" Then
            On Error Resume Next
            Dim dataPedido As Date
            dataPedido = CDate(ThisWorkbook.Worksheets(i).Range("J5").Value)
            
            If DateValue(dataPedido) = DateValue(Date) Then
                pedidosHoje = pedidosHoje + 1
            End If
            On Error GoTo TratarErro
        End If
    Next i
    
    If pedidosHoje > 0 Then
        ticketMedioHoje = vendasHoje / pedidosHoje
    End If
    
    ' Exibir resumo
    Dim resumo As String
    resumo = "📊 RESUMO DO DIA - " & Format(Date, "dd/mm/yyyy") & vbCrLf & vbCrLf
    resumo = resumo & "💰 Vendas hoje: " & Format(vendasHoje, "R$ #,##0.00") & vbCrLf
    resumo = resumo & "📋 Pedidos hoje: " & pedidosHoje & vbCrLf
    resumo = resumo & "🎯 Ticket médio: " & Format(ticketMedioHoje, "R$ #,##0.00") & vbCrLf & vbCrLf
    resumo = resumo & "📈 Comparativo:" & vbCrLf
    resumo = resumo & "Esta semana: " & Format(CalcularVendasSemana(), "R$ #,##0.00") & vbCrLf
    resumo = resumo & "Este mês: " & Format(CalcularVendasMes(), "R$ #,##0.00")
    
    MsgBox resumo, vbInformation, "Resumo Diário"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarResumoDiario", Err)
End Sub

' === INICIALIZAR DASHBOARD ===
Public Sub InicializarDashboard()
    On Error GoTo TratarErro
    
    ' Atualizar todas as estatísticas
    Call AtualizarEstatisticasDashboard
    
    ' Configurar botões
    Call ConfigurarBotoesDashboard
    
    ' Monitorar estoque
    Call MonitorarEstoqueBaixo
    
    ' Criar gráfico de vendas
    Call CriarGraficoVendas
    
    ' Abrir Dashboard
    Call AbrirDashboard
    
    MsgBox "✅ Dashboard inicializado com sucesso!" & vbCrLf & _
           "Todas as funcionalidades estão ativas.", vbInformation, "Dashboard Pronto"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("InicializarDashboard", Err)
End Sub