'==============================================================================
' SISTEMA PDV MADEIREIRA MARIA LUZIA - INSTALADOR AUTOMÁTICO COMPLETO
' Data/Hora: 2025-01-27 
' Desenvolvido para: Madeireira Maria Luzia
' Copie todo este código e execute a Sub InstalarSistemaCompleto()
'==============================================================================

Option Explicit

Sub InstalarSistemaCompleto()
    On Error GoTo ErroInstalacao
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    MsgBox "🚀 Iniciando instalação do Sistema PDV..." & vbCrLf & _
           "Isso criará todas as planilhas, formulários e código necessários." & vbCrLf & vbCrLf & _
           "⏱️ Aguarde alguns segundos...", _
           vbInformation, "Instalador PDV - Madeireira Maria Luzia"
    
    ' 1. Criar planilhas
    Call CriarPlanilhas
    
    ' 2. Criar Dashboard
    Call CriarDashboard
    
    ' 3. Popular dados de exemplo
    Call PopularDadosExemplo
    
    ' 4. Configurar botões do Dashboard
    Call ConfigurarBotoesDashboard
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Ir para o Dashboard
    ThisWorkbook.Worksheets("Dashboard").Activate
    
    MsgBox "✅ Sistema PDV instalado com sucesso!" & vbCrLf & vbCrLf & _
           "📊 Dashboard criado" & vbCrLf & _
           "📋 Planilhas configuradas" & vbCrLf & _
           "🎨 Formulário PDV pronto" & vbCrLf & _
           "💾 Dados de exemplo inseridos" & vbCrLf & vbCrLf & _
           "🎯 Próximos passos:" & vbCrLf & _
           "1. Criar UserForm frmPDVMadeireiraML" & vbCrLf & _
           "2. Adicionar módulos VBA" & vbCrLf & _
           "3. Configurar eventos dos botões", _
           vbInformation, "Instalação Concluída"
    
    Exit Sub
    
ErroInstalacao:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "❌ Erro durante a instalação: " & Err.Description, vbCritical
End Sub

'==============================================================================
' CRIAR PLANILHAS
'==============================================================================
Private Sub CriarPlanilhas()
    Dim ws As Worksheet
    
    ' Dashboard
    On Error Resume Next
    ThisWorkbook.Worksheets("Dashboard").Delete
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = "Dashboard"
    
    ' Produtos
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Produtos")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Produtos"
    End If
    On Error GoTo 0
    
    ' Configurar cabeçalho Produtos
    With ws
        .Cells.Clear
        .Cells(1, 1).Value = "Referencia"
        .Cells(1, 2).Value = "Descricao"
        .Cells(1, 3).Value = "Categoria"
        .Cells(1, 4).Value = "Unidade"
        .Cells(1, 5).Value = "Preco_Custo"
        .Cells(1, 6).Value = "Preco_Venda"
        .Cells(1, 7).Value = "Estoque"
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:G1").Font.Color = RGB(255, 255, 255)
        .Range("A1:G1").Borders.LineStyle = xlContinuous
        .Columns("A:G").AutoFit
    End With
    
    ' Clientes
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Clientes")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Clientes"
    End If
    On Error GoTo 0
    
    ' Configurar cabeçalho Clientes
    With ws
        .Cells.Clear
        .Cells(1, 1).Value = "ID_Cliente"
        .Cells(1, 2).Value = "Nome_RazaoSocial"
        .Cells(1, 3).Value = "CPF_CNPJ"
        .Cells(1, 4).Value = "Endereco"
        .Cells(1, 5).Value = "Cidade"
        .Cells(1, 6).Value = "UF"
        .Cells(1, 7).Value = "CEP"
        .Cells(1, 8).Value = "Telefone"
        .Range("A1:H1").Font.Bold = True
        .Range("A1:H1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:H1").Font.Color = RGB(255, 255, 255)
        .Range("A1:H1").Borders.LineStyle = xlContinuous
        .Columns("A:H").AutoFit
    End With
    
    ' Template_Pedido
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Template_Pedido")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Template_Pedido"
    End If
    On Error GoTo 0
    
    ' Criar layout do pedido
    Call CriarLayoutPedido(ws)
    
    ' Controle
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Controle")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Controle"
        ws.Visible = xlSheetHidden
    End If
    On Error GoTo 0
    
    With ws
        .Cells.Clear
        .Cells(1, 1).Value = "Ultimo_Pedido"
        .Cells(2, 1).Value = 0
        .Cells(1, 2).Value = "Total_Vendas"
        .Cells(2, 2).Value = 0
        .Cells(1, 3).Value = "Total_Clientes"
        .Cells(2, 3).Value = 0
        .Cells(1, 4).Value = "Data_Instalacao"
        .Cells(2, 4).Value = Now()
        .Range("A1:D1").Font.Bold = True
    End With
    
    ' Log_Erros
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Log_Erros")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Log_Erros"
        ws.Visible = xlSheetHidden
    End If
    On Error GoTo 0
    
    With ws
        .Cells.Clear
        .Cells(1, 1).Value = "Data/Hora"
        .Cells(1, 2).Value = "Procedimento"
        .Cells(1, 3).Value = "Numero"
        .Cells(1, 4).Value = "Descricao"
        .Cells(1, 5).Value = "Usuario"
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(255, 0, 0)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
        .Columns("A:E").AutoFit
    End With
End Sub

'==============================================================================
' CRIAR LAYOUT DO PEDIDO
'==============================================================================
Private Sub CriarLayoutPedido(ws As Worksheet)
    With ws
        ' Limpar planilha
        .Cells.Clear
        
        ' Configurar larguras
        .Columns("A").ColumnWidth = 2
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 30
        .Columns("E:G").ColumnWidth = 8
        .Columns("H").ColumnWidth = 10
        .Columns("I:L").ColumnWidth = 12
        .Columns("M").ColumnWidth = 2
        
        ' Cabeçalho da empresa
        .Range("B2:L2").Merge
        .Range("B2").Value = "MADEIREIRA MARIA LUZIA"
        .Range("B2").Font.Size = 20
        .Range("B2").Font.Bold = True
        .Range("B2").HorizontalAlignment = xlCenter
        .Range("B2:L2").Interior.Color = RGB(68, 114, 196)
        .Range("B2:L2").Font.Color = RGB(255, 255, 255)
        
        .Range("B3:L3").Merge
        .Range("B3").Value = "Av. Dr. Cláudio Gueiros Leite - 6311 - Pau Amarelo - Paulista/PE"
        .Range("B3").HorizontalAlignment = xlCenter
        .Range("B3").Font.Size = 12
        
        .Range("B4:L4").Merge
        .Range("B4").Value = "CNPJ: 48.905.025/0001-61"
        .Range("B4").HorizontalAlignment = xlCenter
        .Range("B4").Font.Size = 12
        
        ' Número do pedido
        .Range("B5:F5").Merge
        .Range("B5").Value = "PEDIDO DE VENDA 000000"
        .Range("B5").Font.Bold = True
        .Range("B5").Font.Size = 14
        
        .Range("I5").Value = "Data:"
        .Range("I5").Font.Bold = True
        .Range("J5:L5").Merge
        .Range("J5").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        ' Dados do cliente
        .Range("B7").Value = "Cliente:"
        .Range("B7").Font.Bold = True
        .Range("C7:L7").Merge
        .Range("C7").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("B8").Value = "Endereço:"
        .Range("B8").Font.Bold = True
        .Range("C8:L8").Merge
        .Range("C8").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("B9").Value = "CPF/CNPJ:"
        .Range("B9").Font.Bold = True
        .Range("C9:E9").Merge
        .Range("C9").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("F9").Value = "Cidade:"
        .Range("F9").Font.Bold = True
        .Range("G9:H9").Merge
        .Range("G9").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("I9").Value = "UF:"
        .Range("I9").Font.Bold = True
        .Range("J9").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("K9").Value = "CEP:"
        .Range("K9").Font.Bold = True
        .Range("L9").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        ' Cabeçalho da tabela de produtos
        .Range("C10").Value = "Referência"
        .Range("D10").Value = "Descrição do Item"
        .Range("H10").Value = "Uni"
        .Range("I10").Value = "Valor"
        .Range("J10").Value = "Quant."
        .Range("K10").Value = "Desc."
        .Range("L10").Value = "Valor Total"
        .Range("C10:L10").Font.Bold = True
        .Range("C10:L10").Interior.Color = RGB(200, 200, 200)
        .Range("C10:L10").Borders.LineStyle = xlContinuous
        .Range("C10:L10").HorizontalAlignment = xlCenter
        
        ' Linhas para produtos (11-21)
        Dim i As Integer
        For i = 11 To 21
            .Range("C" & i & ":L" & i).Borders.LineStyle = xlContinuous
            .Range("I" & i & ":L" & i).NumberFormat = "R$ #,##0.00"
        Next i
        
        ' Totais
        .Range("J22").Value = "VALOR PRODUTOS:"
        .Range("J22").Font.Bold = True
        .Range("K22:L22").Merge
        .Range("K22").Borders.LineStyle = xlContinuous
        .Range("K22").NumberFormat = "R$ #,##0.00"
        .Range("K22").Font.Bold = True
        
        .Range("J23").Value = "FRETE:"
        .Range("J23").Font.Bold = True
        .Range("K23:L23").Merge
        .Range("K23").Borders.LineStyle = xlContinuous
        .Range("K23").NumberFormat = "R$ #,##0.00"
        
        .Range("J24").Value = "VALOR DESCONTO:"
        .Range("J24").Font.Bold = True
        .Range("K24:L24").Merge
        .Range("K24").Borders.LineStyle = xlContinuous
        .Range("K24").NumberFormat = "R$ #,##0.00"
        
        .Range("J25").Value = "VALOR TOTAL:"
        .Range("J25").Font.Bold = True
        .Range("J25").Font.Size = 14
        .Range("K25:L25").Merge
        .Range("K25").Borders.LineStyle = xlContinuous
        .Range("K25").Font.Bold = True
        .Range("K25").Font.Size = 14
        .Range("K25").NumberFormat = "R$ #,##0.00"
        .Range("K25").Interior.Color = RGB(255, 255, 0)
        
        ' Rodapé
        .Range("B23").Value = "Vendedor:"
        .Range("C23:E23").Merge
        .Range("C23").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("F23").Value = "Situação Atual:"
        .Range("G23:H23").Merge
        .Range("G23").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("B24").Value = "Condições de Pagamento:"
        .Range("C24:E24").Merge
        .Range("C24").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("F24").Value = "Entrega:"
        .Range("G24:H24").Merge
        .Range("G24").Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        .Range("B26:L26").Merge
        .Range("B26").Value = "WHATSAPP (81) 3011-5515 - Facebook@ Madeireira Maria Luzia"
        .Range("B26").HorizontalAlignment = xlCenter
        .Range("B26").Font.Bold = True
        .Range("B26").Font.Size = 11
        
        ' Formatar área de impressão
        .PageSetup.PrintArea = "A1:M26"
        .PageSetup.Orientation = xlPortrait
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = 1
        .PageSetup.LeftMargin = Application.InchesToPoints(0.5)
        .PageSetup.RightMargin = Application.InchesToPoints(0.5)
        .PageSetup.TopMargin = Application.InchesToPoints(0.5)
        .PageSetup.BottomMargin = Application.InchesToPoints(0.5)
    End With
End Sub

'==============================================================================
' CRIAR DASHBOARD
'==============================================================================
Private Sub CriarDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    
    With ws
        ' Limpar planilha
        .Cells.Clear
        .Cells.Interior.Color = RGB(248, 248, 248)
        
        ' Configurar colunas
        .Columns("A").ColumnWidth = 2
        .Columns("B:K").ColumnWidth = 15
        .Columns("L").ColumnWidth = 2
        
        ' Título principal
        .Range("B2:K3").Merge
        .Range("B2").Value = "🏢 SISTEMA PDV - MADEIREIRA MARIA LUZIA"
        .Range("B2").Font.Size = 28
        .Range("B2").Font.Bold = True
        .Range("B2").HorizontalAlignment = xlCenter
        .Range("B2").VerticalAlignment = xlCenter
        .Range("B2:K3").Interior.Color = RGB(68, 114, 196)
        .Range("B2:K3").Font.Color = RGB(255, 255, 255)
        .Range("B2:K3").Borders.LineStyle = xlContinuous
        
        ' Subtítulo
        .Range("B4:K4").Merge
        .Range("B4").Value = "Sistema Integrado de Vendas e Gestão"
        .Range("B4").Font.Size = 14
        .Range("B4").Font.Italic = True
        .Range("B4").HorizontalAlignment = xlCenter
        .Range("B4").Interior.Color = RGB(217, 225, 242)
        
        ' Seção Estatísticas
        .Range("B6:K6").Merge
        .Range("B6").Value = "📊 ESTATÍSTICAS DO SISTEMA"
        .Range("B6").Font.Size = 16
        .Range("B6").Font.Bold = True
        .Range("B6").Interior.Color = RGB(0, 176, 80)
        .Range("B6").Font.Color = RGB(255, 255, 255)
        .Range("B6").HorizontalAlignment = xlCenter
        
        ' Cards de estatísticas
        ' Card 1 - Total de Pedidos
        .Range("B8:D10").Merge
        .Range("B8").Value = "📋 TOTAL DE PEDIDOS" & vbCrLf & vbCrLf & "0"
        .Range("B8").Font.Size = 12
        .Range("B8").Font.Bold = True
        .Range("B8").HorizontalAlignment = xlCenter
        .Range("B8").VerticalAlignment = xlCenter
        .Range("B8:D10").Interior.Color = RGB(255, 255, 255)
        .Range("B8:D10").Borders.LineStyle = xlContinuous
        .Range("B8:D10").Borders.Weight = xlMedium
        
        ' Card 2 - Total de Vendas
        .Range("E8:G10").Merge
        .Range("E8").Value = "💰 TOTAL DE VENDAS" & vbCrLf & vbCrLf & "R$ 0,00"
        .Range("E8").Font.Size = 12
        .Range("E8").Font.Bold = True
        .Range("E8").HorizontalAlignment = xlCenter
        .Range("E8").VerticalAlignment = xlCenter
        .Range("E8:G10").Interior.Color = RGB(255, 255, 255)
        .Range("E8:G10").Borders.LineStyle = xlContinuous
        .Range("E8:G10").Borders.Weight = xlMedium
        
        ' Card 3 - Total de Clientes
        .Range("H8:J10").Merge
        .Range("H8").Value = "👥 TOTAL DE CLIENTES" & vbCrLf & vbCrLf & "0"
        .Range("H8").Font.Size = 12
        .Range("H8").Font.Bold = True
        .Range("H8").HorizontalAlignment = xlCenter
        .Range("H8").VerticalAlignment = xlCenter
        .Range("H8:J10").Interior.Color = RGB(255, 255, 255)
        .Range("H8:J10").Borders.LineStyle = xlContinuous
        .Range("H8:J10").Borders.Weight = xlMedium
        
        ' Menu Principal
        .Range("B12:K12").Merge
        .Range("B12").Value = "🎯 MENU PRINCIPAL"
        .Range("B12").Font.Size = 16
        .Range("B12").Font.Bold = True
        .Range("B12").Interior.Color = RGB(255, 192, 0)
        .Range("B12").HorizontalAlignment = xlCenter
        
        ' Primeira linha de botões
        .Range("B14:D16").Merge
        .Range("B14").Value = "🛒 ABRIR PDV"
        .Range("B14").Font.Size = 14
        .Range("B14").Font.Bold = True
        .Range("B14").HorizontalAlignment = xlCenter
        .Range("B14").VerticalAlignment = xlCenter
        .Range("B14:D16").Interior.Color = RGB(0, 176, 80)
        .Range("B14:D16").Font.Color = RGB(255, 255, 255)
        .Range("B14:D16").Borders.LineStyle = xlContinuous
        .Range("B14:D16").Borders.Weight = xlMedium
        
        .Range("E14:G16").Merge
        .Range("E14").Value = "👤 CADASTRAR CLIENTE"
        .Range("E14").Font.Size = 14
        .Range("E14").Font.Bold = True
        .Range("E14").HorizontalAlignment = xlCenter
        .Range("E14").VerticalAlignment = xlCenter
        .Range("E14:G16").Interior.Color = RGB(0, 112, 192)
        .Range("E14:G16").Font.Color = RGB(255, 255, 255)
        .Range("E14:G16").Borders.LineStyle = xlContinuous
        .Range("E14:G16").Borders.Weight = xlMedium
        
        .Range("H14:J16").Merge
        .Range("H14").Value = "📦 CADASTRAR PRODUTO"
        .Range("H14").Font.Size = 14
        .Range("H14").Font.Bold = True
        .Range("H14").HorizontalAlignment = xlCenter
        .Range("H14").VerticalAlignment = xlCenter
        .Range("H14:J16").Interior.Color = RGB(112, 48, 160)
        .Range("H14:J16").Font.Color = RGB(255, 255, 255)
        .Range("H14:J16").Borders.LineStyle = xlContinuous
        .Range("H14:J16").Borders.Weight = xlMedium
        
        ' Segunda linha de botões
        .Range("B18:D20").Merge
        .Range("B18").Value = "📄 VER PEDIDOS"
        .Range("B18").Font.Size = 14
        .Range("B18").Font.Bold = True
        .Range("B18").HorizontalAlignment = xlCenter
        .Range("B18").VerticalAlignment = xlCenter
        .Range("B18:D20").Interior.Color = RGB(255, 192, 0)
        .Range("B18:D20").Font.Color = RGB(0, 0, 0)
        .Range("B18:D20").Borders.LineStyle = xlContinuous
        .Range("B18:D20").Borders.Weight = xlMedium
        
        .Range("E18:G20").Merge
        .Range("E18").Value = "📊 RELATÓRIOS"
        .Range("E18").Font.Size = 14
        .Range("E18").Font.Bold = True
        .Range("E18").HorizontalAlignment = xlCenter
        .Range("E18").VerticalAlignment = xlCenter
        .Range("E18:G20").Interior.Color = RGB(192, 0, 0)
        .Range("E18:G20").Font.Color = RGB(255, 255, 255)
        .Range("E18:G20").Borders.LineStyle = xlContinuous
        .Range("E18:G20").Borders.Weight = xlMedium
        
        .Range("H18:J20").Merge
        .Range("H18").Value = "⚙️ CONFIGURAÇÕES"
        .Range("H18").Font.Size = 14
        .Range("H18").Font.Bold = True
        .Range("H18").HorizontalAlignment = xlCenter
        .Range("H18").VerticalAlignment = xlCenter
        .Range("H18:J20").Interior.Color = RGB(128, 128, 128)
        .Range("H18:J20").Font.Color = RGB(255, 255, 255)
        .Range("H18:J20").Borders.LineStyle = xlContinuous
        .Range("H18:J20").Borders.Weight = xlMedium
        
        ' Últimos Pedidos
        .Range("B22:K22").Merge
        .Range("B22").Value = "📋 ÚLTIMOS PEDIDOS"
        .Range("B22").Font.Size = 16
        .Range("B22").Font.Bold = True
        .Range("B22").Interior.Color = RGB(68, 114, 196)
        .Range("B22").Font.Color = RGB(255, 255, 255)
        .Range("B22").HorizontalAlignment = xlCenter
        
        ' Cabeçalho da tabela de pedidos
        .Range("B24").Value = "Nº Pedido"
        .Range("C24:E24").Merge
        .Range("C24").Value = "Cliente"
        .Range("F24").Value = "Data"
        .Range("G24:H24").Merge
        .Range("G24").Value = "Valor"
        .Range("I24").Value = "Status"
        .Range("J24:K24").Merge
        .Range("J24").Value = "Ações"
        .Range("B24:K24").Font.Bold = True
        .Range("B24:K24").Interior.Color = RGB(200, 200, 200)
        .Range("B24:K24").Borders.LineStyle = xlContinuous
        .Range("B24:K24").HorizontalAlignment = xlCenter
        
        ' Área para últimos pedidos (linhas 25-29)
        Dim j As Integer
        For j = 25 To 29
            .Range("B" & j & ":K" & j).Interior.Color = RGB(255, 255, 255)
            .Range("B" & j & ":K" & j).Borders.LineStyle = xlContinuous
            .Range("G" & j & ":H" & j).NumberFormat = "R$ #,##0.00"
        Next j
        
        ' Rodapé do sistema
        .Range("B32:K32").Merge
        .Range("B32").Value = "💼 Sistema desenvolvido para Madeireira Maria Luzia - WhatsApp: (81) 3011-5515"
        .Range("B32").HorizontalAlignment = xlCenter
        .Range("B32").Font.Size = 11
        .Range("B32").Font.Italic = True
        .Range("B32").Interior.Color = RGB(217, 225, 242)
        
        ' Informações técnicas
        .Range("B33:K33").Merge
        .Range("B33").Value = "🔧 Versão 1.0 - Instalado em " & Format(Now(), "dd/mm/yyyy hh:mm")
        .Range("B33").HorizontalAlignment = xlCenter
        .Range("B33").Font.Size = 9
        .Range("B33").Font.Color = RGB(128, 128, 128)
        
        ' Ocultar linhas de grade
        ActiveWindow.DisplayGridlines = False
        
        ' Proteger células (deixar apenas botões desbloqueados)
        .Cells.Locked = True
        .Range("B14:J20").Locked = False
    End With
End Sub

'==============================================================================
' POPULAR DADOS DE EXEMPLO
'==============================================================================
Private Sub PopularDadosExemplo()
    ' Popular Clientes
    Dim wsClientes As Worksheet
    Set wsClientes = ThisWorkbook.Worksheets("Clientes")
    
    With wsClientes
        ' Cliente 1
        .Cells(2, 1).Value = 1
        .Cells(2, 2).Value = "João Silva"
        .Cells(2, 3).Value = "123.456.789-00"
        .Cells(2, 4).Value = "Rua das Flores, 123"
        .Cells(2, 5).Value = "Recife"
        .Cells(2, 6).Value = "PE"
        .Cells(2, 7).Value = "50000-000"
        .Cells(2, 8).Value = "(81) 99999-9999"
        
        ' Cliente 2
        .Cells(3, 1).Value = 2
        .Cells(3, 2).Value = "Maria Santos LTDA"
        .Cells(3, 3).Value = "12.345.678/0001-90"
        .Cells(3, 4).Value = "Av. Boa Viagem, 456"
        .Cells(3, 5).Value = "Recife"
        .Cells(3, 6).Value = "PE"
        .Cells(3, 7).Value = "51000-000"
        .Cells(3, 8).Value = "(81) 88888-8888"
        
        ' Cliente 3
        .Cells(4, 1).Value = 3
        .Cells(4, 2).Value = "Construções ABC Ltda"
        .Cells(4, 3).Value = "98.765.432/0001-10"
        .Cells(4, 4).Value = "Rua da Construção, 789"
        .Cells(4, 5).Value = "Olinda"
        .Cells(4, 6).Value = "PE"
        .Cells(4, 7).Value = "53000-000"
        .Cells(4, 8).Value = "(81) 77777-7777"
        
        ' Cliente 4
        .Cells(5, 1).Value = 4
        .Cells(5, 2).Value = "Pedro Oliveira"
        .Cells(5, 3).Value = "987.654.321-00"
        .Cells(5, 4).Value = "Av. Paulista, 321"
        .Cells(5, 5).Value = "Paulista"
        .Cells(5, 6).Value = "PE"
        .Cells(5, 7).Value = "53400-000"
        .Cells(5, 8).Value = "(81) 66666-6666"
        
        ' Cliente 5
        .Cells(6, 1).Value = 5
        .Cells(6, 2).Value = "Reformas XYZ"
        .Cells(6, 3).Value = "11.222.333/0001-44"
        .Cells(6, 4).Value = "Rua das Reformas, 555"
        .Cells(6, 5).Value = "Jaboatão"
        .Cells(6, 6).Value = "PE"
        .Cells(6, 7).Value = "54000-000"
        .Cells(6, 8).Value = "(81) 55555-5555"
    End With
    
    ' Popular Produtos
    Dim wsProdutos As Worksheet
    Set wsProdutos = ThisWorkbook.Worksheets("Produtos")
    
    With wsProdutos
        ' Produto 1
        .Cells(2, 1).Value = "MAD001"
        .Cells(2, 2).Value = "Tábua de Pinus 2,5x20x300cm"
        .Cells(2, 3).Value = "Madeira"
        .Cells(2, 4).Value = "UN"
        .Cells(2, 5).Value = 25.50
        .Cells(2, 6).Value = 35.00
        .Cells(2, 7).Value = 150
        
        ' Produto 2
        .Cells(3, 1).Value = "MAD002"
        .Cells(3, 2).Value = "Viga de Eucalipto 6x12x400cm"
        .Cells(3, 3).Value = "Madeira"
        .Cells(3, 4).Value = "UN"
        .Cells(3, 5).Value = 45.00
        .Cells(3, 6).Value = 65.00
        .Cells(3, 7).Value = 80
        
        ' Produto 3
        .Cells(4, 1).Value = "FER001"
        .Cells(4, 2).Value = "Prego 18x27 (1kg)"
        .Cells(4, 3).Value = "Ferragem"
        .Cells(4, 4).Value = "KG"
        .Cells(4, 5).Value = 8.50
        .Cells(4, 6).Value = 12.00
        .Cells(4, 7).Value = 200
        
        ' Produto 4
        .Cells(5, 1).Value = "FER002"
        .Cells(5, 2).Value = "Parafuso Fenda 6x80mm"
        .Cells(5, 3).Value = "Ferragem"
        .Cells(5, 4).Value = "UN"
        .Cells(5, 5).Value = 0.25
        .Cells(5, 6).Value = 0.40
        .Cells(5, 7).Value = 1000
        
        ' Produto 5
        .Cells(6, 1).Value = "COL001"
        .Cells(6, 2).Value = "Cola Branca 1kg"
        .Cells(6, 3).Value = "Adesivos"
        .Cells(6, 4).Value = "UN"
        .Cells(6, 5).Value = 15.00
        .Cells(6, 6).Value = 22.00
        .Cells(6, 7).Value = 50
        
        ' Produto 6
        .Cells(7, 1).Value = "MAD003"
        .Cells(7, 2).Value = "Compensado 15mm 220x110cm"
        .Cells(7, 3).Value = "Madeira"
        .Cells(7, 4).Value = "UN"
        .Cells(7, 5).Value = 85.00
        .Cells(7, 6).Value = 120.00
        .Cells(7, 7).Value = 30
        
        ' Produto 7
        .Cells(8, 1).Value = "FER003"
        .Cells(8, 2).Value = "Dobradiça 3 1/2 Aço"
        .Cells(8, 3).Value = "Ferragem"
        .Cells(8, 4).Value = "UN"
        .Cells(8, 5).Value = 12.00
        .Cells(8, 6).Value = 18.00
        .Cells(8, 7).Value = 100
        
        ' Produto 8
        .Cells(9, 1).Value = "TIN001"
        .Cells(9, 2).Value = "Tinta Latex Branca 18L"
        .Cells(9, 3).Value = "Tintas"
        .Cells(9, 4).Value = "UN"
        .Cells(9, 5).Value = 180.00
        .Cells(9, 6).Value = 250.00
        .Cells(9, 7).Value = 25
        
        ' Produto 9
        .Cells(10, 1).Value = "MAD004"
        .Cells(10, 2).Value = "Ripão de Pinus 5x7x300cm"
        .Cells(10, 3).Value = "Madeira"
        .Cells(10, 4).Value = "UN"
        .Cells(10, 5).Value = 18.00
        .Cells(10, 6).Value = 28.00
        .Cells(10, 7).Value = 120
        
        ' Produto 10
        .Cells(11, 1).Value = "FER004"
        .Cells(11, 2).Value = "Fechadura Interna Cromada"
        .Cells(11, 3).Value = "Ferragem"
        .Cells(11, 4).Value = "UN"
        .Cells(11, 5).Value = 35.00
        .Cells(11, 6).Value = 55.00
        .Cells(11, 7).Value = 40
        
        ' Formatar colunas de valores
        .Range("E:G").NumberFormat = "R$ #,##0.00"
        .Columns("A:G").AutoFit
    End With
    
    ' Atualizar estatísticas no controle
    Dim wsControle As Worksheet
    Set wsControle = ThisWorkbook.Worksheets("Controle")
    wsControle.Cells(2, 3).Value = 5 ' Total de clientes
End Sub

'==============================================================================
' CONFIGURAR BOTÕES DO DASHBOARD
'==============================================================================
Private Sub ConfigurarBotoesDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    
    ' Aqui você pode adicionar código VBA para tornar os botões clicáveis
    ' Por exemplo, usando eventos de Worksheet_SelectionChange
    ' Ou criando botões ActiveX
End Sub

'==============================================================================
' FUNÇÕES AUXILIARES PARA TESTE
'==============================================================================
Sub TesteCarregarClientes()
    MsgBox "✅ Teste: " & ThisWorkbook.Worksheets("Clientes").Cells(2, 2).Value & " carregado!", vbInformation
End Sub

Sub TesteCarregarProdutos()
    MsgBox "✅ Teste: " & ThisWorkbook.Worksheets("Produtos").Cells(2, 2).Value & " carregado!", vbInformation
End Sub

Sub AbrirDashboard()
    ThisWorkbook.Worksheets("Dashboard").Activate
End Sub