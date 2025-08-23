'====================================================================
' SISTEMA PDV MADEIREIRA MARIA LUZIA - VERSÃO INTEGRADA COMPLETA
' UserForm principal melhorado com integração total dos módulos
' Data/Hora: 2025-01-27
' Versão: INTEGRADA v2.0 - SISTEMA MODULAR COMPLETO
'====================================================================

Option Explicit

' ===== VARIÁVEIS GLOBAIS =====
Private proximoPedido As Long
Private dadosClienteAtual As String

' ===== ESTRUTURA DE DADOS =====
Private Type TipoVenda
    NumeroVenda As Long
    numeroPedido As String
    nomeCliente As String
    DataVenda As Date
    dataEntrega As Date
    FormaPagamento As String
    SubTotal As Double
    desconto As Double
    total As Double
End Type

Private VendaCorrente As TipoVenda

' ===== CONSTANTES DO SISTEMA =====
Private Const VERSAO_SISTEMA As String = "INTEGRADA v2.0"
Private Const DEV_USER As String = "Sistema PDV Modular"

' ===== VARIÁVEIS DE CORES =====
Private COR_BRANCO As Long
Private COR_VERDE_VALIDO As Long
Private COR_AMARELO_DIGITANDO As Long
Private COR_LARANJA_PARCIAL As Long
Private COR_VERMELHO_INVALIDO As Long
Private COR_AMARELO_FOCADO As Long

'====================================================================
' INICIALIZAÇÃO DO SISTEMA - VERSÃO INTEGRADA
'====================================================================
Private Sub UserForm_Initialize()
    On Error GoTo TratarErro
    
    ' Sequência de inicialização otimizada
    Call InicializarCores
    Call InicializarVariaveis
    Call CorrigirCoresFormulario
    Call ConfigurarInterface
    Call CarregarDadosIniciais
    
    ' Integração com módulos
    Call ClienteManager.CarregarClientes(Me.cmbCliente)
    Call ProdutoManager.CarregarProdutos(Me.lstProdutos)
    
    Debug.Print "🚀 SISTEMA PDV INTEGRADO INICIALIZADO | " & Format(Now, "hh:mm:ss")
    
    Exit Sub
    
TratarErro:
    Call ErrorHandler.RegistrarErro("UserForm_Initialize", Err)
    MsgBox "❌ ERRO CRÍTICO NA INICIALIZAÇÃO!" & vbCrLf & vbCrLf & _
           "Erro: " & Err.Description & vbCrLf & _
           "Contate o suporte técnico", vbCritical, "Erro Crítico"
End Sub

Private Sub InicializarVariaveis()
    Call InicializarVenda
    dadosClienteAtual = ""
End Sub

Private Sub CarregarDadosIniciais()
    Call PreencherCidades
    Call PreencherBairros
    Call PreencherFormaPagamento
End Sub

'====================================================================
' EVENTOS DE CLIENTE - INTEGRAÇÃO COM ClienteManager
'====================================================================
Private Sub cmbCliente_Change()
    On Error GoTo TratarErro
    
    If Me.cmbCliente.ListIndex >= 0 Then
        ' Usar ClienteManager para preencher dados
        Call ClienteManager.PreencherDadosCliente(Me.cmbCliente, Me)
        
        ' Armazenar dados do cliente globalmente
        dadosClienteAtual = Me.Tag
        Call TransferenciaDados.ArmazenarDadosCliente(dadosClienteAtual)
        
        ' Atualizar interface
        Call AtualizarInterfaceCliente
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("cmbCliente_Change", Err)
End Sub

Private Sub AtualizarInterfaceCliente()
    On Error Resume Next
    
    If dadosClienteAtual <> "" Then
        Dim cliente() As String
        cliente = Split(dadosClienteAtual, "|")
        
        If UBound(cliente) >= 7 Then
            ' Preencher campos se existirem no formulário atual
            If CampoExiste("txtNome") Then Me.txtNome.Value = cliente(1)
            If CampoExiste("txtCPF") Then Me.txtCPF.Value = cliente(2)
            If CampoExiste("txtEnder") Then Me.txtEnder.Value = cliente(3)
            If CampoExiste("cCidade") Then Me.cCidade.Value = cliente(4)
            If CampoExiste("txtCEP") Then Me.txtCEP.Value = cliente(6)
        End If
    End If
    
    On Error GoTo 0
End Sub

'====================================================================
' EVENTOS DE PRODUTOS - INTEGRAÇÃO COM ProdutoManager
'====================================================================
Private Sub txtPesquisa_Change()
    On Error GoTo TratarErro
    
    ' Usar ProdutoManager para filtrar produtos
    Call ProdutoManager.CarregarProdutos(Me.lstProdutos, Me.txtPesquisa.Text)
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("txtPesquisa_Change", Err)
End Sub

Private Sub btnLimpar_Click()
    On Error GoTo TratarErro
    
    Me.txtPesquisa.Value = ""
    Call ProdutoManager.CarregarProdutos(Me.lstProdutos)
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnLimpar_Click", Err)
End Sub

Private Sub btnAdicionar_Click()
    On Error GoTo TratarErro
    
    If Me.lstProdutos.ListIndex < 0 Then
        MsgBox "⚠️ Selecione um produto!", vbExclamation
        Exit Sub
    End If
    
    Dim quantidade As Long
    quantidade = 1
    
    ' Obter quantidade se campo existir
    If CampoExiste("txtQuantidade") Then
        If IsNumeric(Me.txtQuantidade.Value) And Me.txtQuantidade.Value > 0 Then
            quantidade = CLng(Me.txtQuantidade.Value)
        End If
    End If
    
    ' Usar ProdutoManager para adicionar produto
    Call ProdutoManager.AdicionarProdutoSelecionado(Me.lstProdutos, Me.produtosv1, quantidade)
    
    ' Atualizar totais usando CalculadoraManager
    Call CalculadoraManager.AtualizarTotais(Me, Me.produtosv1)
    
    ' Limpar quantidade
    If CampoExiste("txtQuantidade") Then Me.txtQuantidade.Value = "1"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnAdicionar_Click", Err)
End Sub

Private Sub btnRemover_Click()
    On Error GoTo TratarErro
    
    If Me.produtosv1.ListIndex < 0 Then
        MsgBox "⚠️ Selecione um produto para remover!", vbExclamation
        Exit Sub
    End If
    
    ' Usar ProdutoManager para remover produto
    Call ProdutoManager.RemoverProdutoSelecionado(Me.produtosv1)
    
    ' Atualizar totais
    Call CalculadoraManager.AtualizarTotais(Me, Me.produtosv1)
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnRemover_Click", Err)
End Sub

'====================================================================
' EVENTOS DE DESCONTO - INTEGRAÇÃO COM DescontoManager
'====================================================================
Private Sub btnAplicarDesconto_Click()
    On Error GoTo TratarErro
    
    If Me.produtosv1.ListIndex < 0 Then
        MsgBox "⚠️ Selecione um produto para aplicar desconto!", vbExclamation
        Exit Sub
    End If
    
    Dim valorDesconto As Double
    Dim tipoDesconto As DescontoManager.TipoDesconto
    
    ' Obter valor do desconto
    If CampoExiste("txtDesconto") Then
        If Not IsNumeric(Me.txtDesconto.Value) Or Me.txtDesconto.Value <= 0 Then
            MsgBox "⚠️ Informe um valor de desconto válido!", vbExclamation
            Exit Sub
        End If
        valorDesconto = CDbl(Me.txtDesconto.Value)
    Else
        Exit Sub
    End If
    
    ' Determinar tipo de desconto (assumindo que há radio buttons)
    If CampoExiste("optDescontoPercentual") Then
        If Me.optDescontoPercentual.Value Then
            tipoDesconto = DescontoManager.TipoDesconto.Percentual
        Else
            tipoDesconto = DescontoManager.TipoDesconto.ValorFixo
        End If
    Else
        tipoDesconto = DescontoManager.TipoDesconto.Percentual ' Padrão
    End If
    
    ' Aplicar desconto usando DescontoManager
    Call DescontoManager.AplicarDesconto(Me.produtosv1, tipoDesconto, valorDesconto)
    
    ' Atualizar totais
    Call CalculadoraManager.AtualizarTotais(Me, Me.produtosv1)
    
    ' Limpar campo de desconto
    If CampoExiste("txtDesconto") Then Me.txtDesconto.Value = ""
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnAplicarDesconto_Click", Err)
End Sub

Private Sub btnRemoverDesconto_Click()
    On Error GoTo TratarErro
    
    If Me.produtosv1.ListIndex < 0 Then
        MsgBox "⚠️ Selecione um produto para remover desconto!", vbExclamation
        Exit Sub
    End If
    
    ' Usar DescontoManager para remover desconto
    Call DescontoManager.RemoverDesconto(Me.produtosv1)
    
    ' Atualizar totais
    Call CalculadoraManager.AtualizarTotais(Me, Me.produtosv1)
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnRemoverDesconto_Click", Err)
End Sub

'====================================================================
' EVENTOS PRINCIPAIS - INTEGRAÇÃO COMPLETA
'====================================================================
Private Sub btnGerarPedido_Click()
    On Error GoTo TratarErro
    
    ' Validar dados obrigatórios usando UtilsManager
    Dim nomeCliente As String
    Dim formaPagamento As String
    
    nomeCliente = IIf(CampoExiste("txtNome"), Me.txtNome.Value, Me.cmbCliente.Value)
    formaPagamento = IIf(CampoExiste("cPagamento"), Me.cPagamento.Value, "À Vista")
    
    If Not UtilsManager.ValidarDadosObrigatorios(nomeCliente, formaPagamento, Me.produtosv1) Then
        Exit Sub
    End If
    
    ' Gerar pedido usando PedidoManager
    Dim numeroPedido As String
    numeroPedido = PedidoManager.GerarNovoPedido(dadosClienteAtual, Me.produtosv1, formaPagamento)
    
    If numeroPedido <> "" Then
        MsgBox "✅ PEDIDO GERADO COM SUCESSO!" & vbCrLf & vbCrLf & _
               "📋 Número: " & numeroPedido & vbCrLf & _
               "👤 Cliente: " & nomeCliente & vbCrLf & _
               "💰 Total: " & Format(CalculadoraManager.CalcularTotalGeral(Me.produtosv1), "R$ #,##0.00") & vbCrLf & _
               "💳 Pagamento: " & formaPagamento & vbCrLf & vbCrLf & _
               "O pedido foi salvo e está pronto para impressão!", _
               vbInformation, "Pedido Gerado"
        
        ' Armazenar número do pedido para uso posterior
        Me.Tag = numeroPedido
        
        ' Habilitar botões de ação
        Call HabilitarBotoesAcao(True)
        
        ' Atualizar Dashboard
        Call UtilsManager.AtualizarDashboard
    Else
        MsgBox "❌ Erro ao gerar pedido!", vbCritical
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnGerarPedido_Click", Err)
End Sub

Private Sub btnImprimir_Click()
    On Error GoTo TratarErro
    
    If Me.Tag = "" Then
        MsgBox "⚠️ Gere um pedido primeiro!", vbExclamation
        Exit Sub
    End If
    
    ' Usar ImpressaoManager para imprimir
    Call ImpressaoManager.ImprimirPedido(Me.Tag)
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnImprimir_Click", Err)
End Sub

Private Sub btnExportarPDF_Click()
    On Error GoTo TratarErro
    
    If Me.Tag = "" Then
        MsgBox "⚠️ Gere um pedido primeiro!", vbExclamation
        Exit Sub
    End If
    
    ' Usar ImpressaoManager para exportar PDF
    Call ImpressaoManager.ExportarParaPDF(Me.Tag)
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnExportarPDF_Click", Err)
End Sub

Private Sub btnEnviarEmail_Click()
    On Error GoTo TratarErro
    
    If Me.Tag = "" Then
        MsgBox "⚠️ Gere um pedido primeiro!", vbExclamation
        Exit Sub
    End If
    
    Dim emailCliente As String
    emailCliente = InputBox("Digite o email do cliente:", "Enviar por Email")
    
    If emailCliente <> "" Then
        If UtilsManager.ValidarEmail(emailCliente) Then
            Call ImpressaoManager.EnviarPorEmail(Me.Tag, emailCliente)
        Else
            MsgBox "❌ Email inválido!", vbExclamation
        End If
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnEnviarEmail_Click", Err)
End Sub

'====================================================================
' EVENTOS DE NAVEGAÇÃO E GESTÃO
'====================================================================
Private Sub btnCadastroClientes_Click()
    On Error GoTo TratarErro
    
    ' Abrir formulário de gestão de clientes (se existir)
    On Error Resume Next
    frmGestaoClientes.Show
    On Error GoTo TratarErro
    
    ' Recarregar clientes após possível cadastro
    Call ClienteManager.CarregarClientes(Me.cmbCliente)
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnCadastroClientes_Click", Err)
End Sub

Private Sub btnPesquisaProdutos_Click()
    On Error GoTo TratarErro
    
    ' Abrir formulário de pesquisa de produtos (se existir)
    On Error Resume Next
    frmPesquisaProdutos.Show
    On Error GoTo TratarErro
    
    ' Recarregar produtos após possível alteração
    Call ProdutoManager.CarregarProdutos(Me.lstProdutos)
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnPesquisaProdutos_Click", Err)
End Sub

Private Sub btnNovaVenda_Click()
    On Error GoTo TratarErro
    
    ' Limpar formulário para nova venda
    Call LimparFormularioCompleto
    
    ' Desabilitar botões de ação
    Call HabilitarBotoesAcao(False)
    
    ' Focar no cliente
    Me.cmbCliente.SetFocus
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnNovaVenda_Click", Err)
End Sub

'====================================================================
' VALIDAÇÕES AVANÇADAS - INTEGRAÇÃO COM UtilsManager
'====================================================================
Private Sub txtCPF_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo TratarErro
    
    If Trim(Me.txtCPF.Text) <> "" Then
        Dim cpfValido As Boolean
        
        ' Verificar se é CPF ou CNPJ
        Dim documento As String
        documento = Replace(Replace(Replace(Me.txtCPF.Text, ".", ""), "-", ""), "/", "")
        
        If Len(documento) = 11 Then
            cpfValido = UtilsManager.ValidarCPF(Me.txtCPF.Text)
            If cpfValido Then
                Me.txtCPF.Value = UtilsManager.FormatarCPF(Me.txtCPF.Text)
                Me.txtCPF.BackColor = COR_VERDE_VALIDO
            Else
                Me.txtCPF.BackColor = COR_VERMELHO_INVALIDO
                MsgBox "❌ CPF inválido!", vbExclamation
                Cancel = True
            End If
        ElseIf Len(documento) = 14 Then
            cpfValido = UtilsManager.ValidarCNPJ(Me.txtCPF.Text)
            If cpfValido Then
                Me.txtCPF.Value = UtilsManager.FormatarCNPJ(Me.txtCPF.Text)
                Me.txtCPF.BackColor = COR_VERDE_VALIDO
            Else
                Me.txtCPF.BackColor = COR_VERMELHO_INVALIDO
                MsgBox "❌ CNPJ inválido!", vbExclamation
                Cancel = True
            End If
        Else
            Me.txtCPF.BackColor = COR_VERMELHO_INVALIDO
            MsgBox "❌ Documento deve ter 11 dígitos (CPF) ou 14 dígitos (CNPJ)!", vbExclamation
            Cancel = True
        End If
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("txtCPF_Exit", Err)
End Sub

'====================================================================
' SISTEMA DE NUMERAÇÃO AUTOMÁTICA - VERSÃO INTEGRADA
'====================================================================
Private Function GerarProximoNumeroPedido() As String
    On Error GoTo TratarErro
    
    ' Usar PedidoManager para gerar número
    GerarProximoNumeroPedido = PedidoManager.GerarNovoPedido(dadosClienteAtual, Me.produtosv1, GetFormaPagamento())
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("GerarProximoNumeroPedido", Err)
    GerarProximoNumeroPedido = Format(Now(), "hhmmss")
End Function

Private Function GetFormaPagamento() As String
    On Error Resume Next
    
    If CampoExiste("cPagamento") Then
        GetFormaPagamento = Me.cPagamento.Value
    Else
        GetFormaPagamento = "À Vista"
    End If
    
    On Error GoTo 0
End Function

'====================================================================
' BOTÃO DE IMPRESSÃO PRINCIPAL - VERSÃO INTEGRADA
'====================================================================
Private Sub cImprimir_Click()
    Call btnImprimirIntegrado_Click
End Sub

Private Sub btnImprimirIntegrado_Click()
    On Error GoTo TratarErro
    
    Debug.Print "🖨️ INICIANDO IMPRESSÃO INTEGRADA | " & Format(Now, "hh:mm:ss")
    
    ' ETAPA 1: Validar dados obrigatórios
    Dim nomeCliente As String
    Dim formaPagamento As String
    
    nomeCliente = IIf(CampoExiste("txtNome"), Me.txtNome.Value, Me.cmbCliente.Value)
    formaPagamento = GetFormaPagamento()
    
    If Not UtilsManager.ValidarDadosObrigatorios(nomeCliente, formaPagamento, Me.produtosv1) Then
        Exit Sub
    End If
    
    ' ETAPA 2: Gerar pedido se ainda não foi gerado
    If Me.Tag = "" Then
        Dim numeroPedido As String
        numeroPedido = PedidoManager.GerarNovoPedido(dadosClienteAtual, Me.produtosv1, formaPagamento)
        
        If numeroPedido = "" Then
            MsgBox "❌ Erro ao gerar pedido!", vbCritical
            Exit Sub
        End If
        
        Me.Tag = numeroPedido
    End If
    
    ' ETAPA 3: Imprimir usando ImpressaoManager
    Call ImpressaoManager.ImprimirPedido(Me.Tag)
    
    ' ETAPA 4: Opções pós-impressão
    Call ExibirOpcoesPosImpressao
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnImprimirIntegrado_Click", Err)
    MsgBox "❌ ERRO NA IMPRESSÃO: " & Err.Description, vbCritical
End Sub

Private Sub ExibirOpcoesPosImpressao()
    On Error GoTo TratarErro
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("🎯 O que deseja fazer agora?" & vbCrLf & vbCrLf & _
                     "Sim = Nova Venda" & vbCrLf & _
                     "Não = Manter dados atuais" & vbCrLf & _
                     "Cancelar = Exportar PDF", _
                     vbYesNoCancel + vbQuestion, "Próxima Ação")
    
    Select Case resposta
        Case vbYes
            Call LimparFormularioCompleto
            Call HabilitarBotoesAcao(False)
            
        Case vbNo
            ' Manter dados atuais
            
        Case vbCancel
            Call ImpressaoManager.ExportarParaPDF(Me.Tag)
    End Select
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ExibirOpcoesPosImpressao", Err)
End Sub

'====================================================================
' SISTEMA DE CORES E INTERFACE - VERSÃO INTEGRADA
'====================================================================
Private Sub InicializarCores()
    COR_BRANCO = RGB(255, 255, 255)
    COR_VERDE_VALIDO = RGB(200, 255, 200)
    COR_AMARELO_DIGITANDO = RGB(255, 255, 200)
    COR_LARANJA_PARCIAL = RGB(255, 235, 200)
    COR_VERMELHO_INVALIDO = RGB(255, 200, 200)
    COR_AMARELO_FOCADO = RGB(255, 255, 220)
End Sub

Private Sub CorrigirCoresFormulario()
    On Error Resume Next
    
    ' Cor de fundo do formulário
    Me.BackColor = RGB(245, 245, 245)
    
    ' Corrigir cores de todos os controles
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "TextBox"
                ctrl.BackColor = COR_BRANCO
                ctrl.Forecolor = RGB(0, 0, 0)
                ctrl.BorderStyle = fmBorderStyleSingle
                
            Case "ComboBox"
                ctrl.BackColor = COR_BRANCO
                ctrl.Forecolor = RGB(0, 0, 0)
                ctrl.Style = fmStyleDropDownCombo
                
            Case "ListBox"
                ctrl.BackColor = COR_BRANCO
                ctrl.Forecolor = RGB(0, 0, 0)
                
            Case "CommandButton"
                ctrl.BackColor = RGB(230, 230, 230)
                ctrl.Forecolor = RGB(0, 0, 0)
        End Select
    Next ctrl
    
    On Error GoTo 0
End Sub

Private Sub ConfigurarInterface()
    On Error Resume Next
    
    ' Configurar propriedades do formulário
    With Me
        .Caption = "🏪 PDV Maria Luzia - " & VERSAO_SISTEMA
        .Width = 950
        .Height = 720
        .StartUpPosition = 1 ' CenterOwner
    End With
    
    ' Configurar ListBox de produtos se existir
    If CampoExiste("lstProdutos") Then
        With Me.lstProdutos
            .ColumnCount = 6
            .ColumnWidths = "80;200;80;50;80;80"
        End With
    End If
    
    ' Configurar ListBox de produtos selecionados
    If CampoExiste("produtosv1") Then
        With Me.produtosv1
            .ColumnCount = 7
            .ColumnWidths = "60;150;40;80;50;60;80"
        End With
    End If
    
    ' Configurar tooltips
    Call ConfigurarTooltips
    
    ' Desabilitar botões de ação inicialmente
    Call HabilitarBotoesAcao(False)
    
    On Error GoTo 0
End Sub

Private Sub ConfigurarTooltips()
    On Error Resume Next
    
    If CampoExiste("cmbCliente") Then
        Me.cmbCliente.ControlTipText = "Selecione um cliente cadastrado"
    End If
    
    If CampoExiste("txtPesquisa") Then
        Me.txtPesquisa.ControlTipText = "Digite para pesquisar produtos"
    End If
    
    If CampoExiste("txtQuantidade") Then
        Me.txtQuantidade.ControlTipText = "Quantidade do produto (padrão: 1)"
    End If
    
    If CampoExiste("txtDesconto") Then
        Me.txtDesconto.ControlTipText = "Valor do desconto (% ou R$)"
    End If
    
    On Error GoTo 0
End Sub

Private Sub HabilitarBotoesAcao(habilitar As Boolean)
    On Error Resume Next
    
    If CampoExiste("btnImprimir") Then Me.btnImprimir.Enabled = habilitar
    If CampoExiste("btnExportarPDF") Then Me.btnExportarPDF.Enabled = habilitar
    If CampoExiste("btnEnviarEmail") Then Me.btnEnviarEmail.Enabled = habilitar
    
    On Error GoTo 0
End Sub

'====================================================================
' PREENCHIMENTO DE DADOS ESTÁTICOS - VERSÃO INTEGRADA
'====================================================================
Private Sub PreencherCidades()
    On Error Resume Next
    
    If CampoExiste("cCidade") Then
        With Me.cCidade
            .Clear
            .BackColor = COR_BRANCO
            
            ' Cidades principais de Pernambuco
            .AddItem "Abreu e Lima"
            .AddItem "Recife"
            .AddItem "Olinda"
            .AddItem "Jaboatão dos Guararapes"
            .AddItem "Caruaru"
            .AddItem "Petrolina"
            .AddItem "Paulista"
            .AddItem "Cabo de Santo Agostinho"
            .AddItem "Camaragibe"
            .AddItem "Garanhuns"
            .AddItem "Vitória de Santo Antão"
            .AddItem "Igarassu"
            .AddItem "São Lourenço da Mata"
        End With
    End If
    
    On Error GoTo 0
End Sub

Private Sub PreencherBairros()
    On Error Resume Next
    
    If CampoExiste("cbairro1") Then
        With Me.cbairro1
            .Clear
            .BackColor = COR_BRANCO
            
            ' Bairros principais de Abreu e Lima e região
            .AddItem "Pau Amarelo"
            .AddItem "Centro"
            .AddItem "Caetés I"
            .AddItem "Caetés II"
            .AddItem "Caetés III"
            .AddItem "Planalto"
            .AddItem "Pitanga"
            .AddItem "Desterro"
            .AddItem "Jardim Caetés"
            .AddItem "Jardim Paulista"
            .AddItem "Timbi"
            .AddItem "Vila das Flores"
        End With
    End If
    
    On Error GoTo 0
End Sub

Private Sub PreencherFormaPagamento()
    On Error Resume Next
    
    If CampoExiste("cPagamento") Then
        With Me.cPagamento
            .Clear
            .BackColor = COR_BRANCO
            
            ' Formas de pagamento com descontos
            .AddItem "💰 À Vista (5% desconto)"
            .AddItem "📱 PIX (3% desconto)"
            .AddItem "💳 Cartão de Débito"
            .AddItem "💳 Cartão de Crédito"
            .AddItem "📋 Parcelado (sem juros até 3x)"
            
            .ListIndex = 0 ' Padrão: À Vista
        End With
    End If
    
    On Error GoTo 0
End Sub

'====================================================================
' EVENTOS DE SISTEMA DE PAGAMENTOS - VERSÃO INTEGRADA
'====================================================================
Private Sub cPagamento_Change()
    On Error GoTo TratarErro
    
    If Me.cPagamento.ListIndex >= 0 Then
        Dim valorSelecionado As String
        valorSelecionado = Me.cPagamento.Value
        
        ' Processar forma de pagamento
        Call ProcessarFormaPagamento(valorSelecionado)
        
        ' Atualizar totais se há produtos
        If Me.produtosv1.ListCount > 0 Then
            Call CalculadoraManager.AtualizarTotais(Me, Me.produtosv1)
        End If
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("cPagamento_Change", Err)
End Sub

Private Sub ProcessarFormaPagamento(formaPagamento As String)
    On Error GoTo TratarErro
    
    ' Atualizar estrutura da venda
    VendaCorrente.FormaPagamento = formaPagamento
    
    ' Definir desconto baseado na forma de pagamento
    Select Case True
        Case InStr(formaPagamento, "À Vista") > 0
            VendaCorrente.desconto = 5
            Me.cPagamento.BackColor = COR_VERDE_VALIDO
            
        Case InStr(formaPagamento, "PIX") > 0
            VendaCorrente.desconto = 3
            Me.cPagamento.BackColor = COR_VERDE_VALIDO
            
        Case Else
            VendaCorrente.desconto = 0
            Me.cPagamento.BackColor = COR_BRANCO
    End Select
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ProcessarFormaPagamento", Err)
End Sub

'====================================================================
' SISTEMA DE VALIDAÇÃO DE CEP - VERSÃO INTEGRADA
'====================================================================
Private Sub txtCEP_Change()
    On Error Resume Next
    
    If CampoExiste("txtCEP") Then
        Dim cepTexto As String
        cepTexto = Me.txtCEP.Text
        
        ' Validação progressiva
        If Trim(cepTexto) = "" Then
            Me.txtCEP.BackColor = COR_BRANCO
        ElseIf Len(Replace(Replace(cepTexto, "-", ""), " ", "")) < 8 Then
            Me.txtCEP.BackColor = COR_AMARELO_DIGITANDO
            ' Aplicar formatação automática
            Me.txtCEP.Value = UtilsManager.FormatarCEP(cepTexto)
        Else
            ' Validar CEP completo
            If ErrorHandler.ValidarCEP(cepTexto) Then
                Me.txtCEP.BackColor = COR_VERDE_VALIDO
            Else
                Me.txtCEP.BackColor = COR_VERMELHO_INVALIDO
            End If
        End If
    End If
    
    On Error GoTo 0
End Sub

'====================================================================
' FUNÇÕES AUXILIARES - VERSÃO INTEGRADA
'====================================================================
Private Function CampoExiste(nomeCampo As String) As Boolean
    On Error Resume Next
    CampoExiste = Not (Me.Controls(nomeCampo) Is Nothing)
    On Error GoTo 0
End Function

Private Sub InicializarVenda()
    With VendaCorrente
        .NumeroVenda = 0
        .numeroPedido = ""
        .nomeCliente = ""
        .DataVenda = Date
        .dataEntrega = Date + 1
        .FormaPagamento = ""
        .SubTotal = 0
        .desconto = 0
        .total = 0
    End With
End Sub

Private Sub LimparFormularioCompleto()
    On Error Resume Next
    
    ' Limpar campos de cliente
    If CampoExiste("cmbCliente") Then Me.cmbCliente.ListIndex = -1
    If CampoExiste("txtNome") Then Me.txtNome.Value = ""
    If CampoExiste("txtCPF") Then Me.txtCPF.Value = ""
    If CampoExiste("txtEnder") Then Me.txtEnder.Value = ""
    If CampoExiste("txtnumero") Then Me.txtnumero.Value = ""
    If CampoExiste("txtCEP") Then Me.txtCEP.Value = ""
    If CampoExiste("cCidade") Then Me.cCidade.ListIndex = -1
    If CampoExiste("cbairro1") Then Me.cbairro1.ListIndex = -1
    
    ' Limpar campos de pesquisa
    If CampoExiste("txtPesquisa") Then Me.txtPesquisa.Value = ""
    If CampoExiste("txtQuantidade") Then Me.txtQuantidade.Value = "1"
    If CampoExiste("txtDesconto") Then Me.txtDesconto.Value = ""
    
    ' Limpar listas
    If CampoExiste("produtosv1") Then Me.produtosv1.Clear
    
    ' Resetar forma de pagamento
    If CampoExiste("cPagamento") Then Me.cPagamento.ListIndex = 0
    
    ' Limpar dados globais
    dadosClienteAtual = ""
    Me.Tag = ""
    Call TransferenciaDados.LimparDadosGlobais
    
    ' Reinicializar venda
    Call InicializarVenda
    
    ' Recarregar produtos
    Call ProdutoManager.CarregarProdutos(Me.lstProdutos)
    
    Debug.Print "🧹 Formulário limpo para nova venda | " & Format(Now, "hh:mm:ss")
    
    On Error GoTo 0
End Sub

'====================================================================
' EVENTOS DE COMPATIBILIDADE COM FORMULÁRIO ATUAL
'====================================================================
' Manter compatibilidade com eventos existentes
Private Sub CcadrastroClientes_Click()
    Call btnCadastroClientes_Click
End Sub

Private Sub cProdutos_Click()
    Call btnPesquisaProdutos_Click
End Sub

Private Sub produtosv2_Click()
    ' Configurar produtosv2 se existir
    On Error Resume Next
    If CampoExiste("produtosv2") Then
        With Me.produtosv2
            .ColumnCount = 7
            .ColumnWidths = "60;150;40;80;50;100;80"
        End With
    End If
    On Error GoTo 0
End Sub

' Eventos de formatação automática
Private Sub txtCPF_Change()
    On Error Resume Next
    
    If CampoExiste("txtCPF") Then
        ' Aplicar formatação automática
        Dim documento As String
        documento = Replace(Replace(Replace(Me.txtCPF.Text, ".", ""), "-", ""), "/", "")
        
        If Len(documento) <= 11 Then
            Me.txtCPF.Value = UtilsManager.FormatarCPF(Me.txtCPF.Text)
        Else
            Me.txtCPF.Value = UtilsManager.FormatarCNPJ(Me.txtCPF.Text)
        End If
    End If
    
    On Error GoTo 0
End Sub

'====================================================================
' EVENTOS DE MENU E NAVEGAÇÃO
'====================================================================
Private Sub btnFechar_Click()
    On Error GoTo TratarErro
    
    ' Verificar se há dados não salvos
    If Me.produtosv1.ListCount > 0 And Me.Tag = "" Then
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("⚠️ Há produtos não salvos!" & vbCrLf & _
                         "Deseja gerar o pedido antes de sair?", _
                         vbYesNoCancel + vbQuestion, "Dados Não Salvos")
        
        Select Case resposta
            Case vbYes
                Call btnGerarPedido_Click
                If Me.Tag <> "" Then Unload Me
                
            Case vbNo
                Unload Me
                
            Case vbCancel
                Exit Sub
        End Select
    Else
        Unload Me
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnFechar_Click", Err)
End Sub

'====================================================================
' EVENTOS DE RESUMO E RELATÓRIOS
'====================================================================
Private Sub btnResumoFinanceiro_Click()
    On Error GoTo TratarErro
    
    If Me.produtosv1.ListCount = 0 Then
        MsgBox "⚠️ Adicione produtos para ver o resumo financeiro!", vbExclamation
        Exit Sub
    End If
    
    Dim resumo As String
    resumo = CalculadoraManager.GerarResumoFinanceiro(Me.produtosv1, GetFormaPagamento(), 0)
    
    MsgBox resumo, vbInformation, "Resumo Financeiro"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnResumoFinanceiro_Click", Err)
End Sub

'====================================================================
' EVENTOS DE VALIDAÇÃO EM TEMPO REAL
'====================================================================
Private Sub produtosv1_Click()
    On Error GoTo TratarErro
    
    ' Atualizar totais quando lista muda
    Call CalculadoraManager.AtualizarTotais(Me, Me.produtosv1)
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("produtosv1_Click", Err)
End Sub

Private Sub txtQuantidade_Change()
    On Error Resume Next
    
    If CampoExiste("txtQuantidade") Then
        ' Validar quantidade
        If Not IsNumeric(Me.txtQuantidade.Value) Or Me.txtQuantidade.Value <= 0 Then
            Me.txtQuantidade.BackColor = COR_VERMELHO_INVALIDO
        Else
            Me.txtQuantidade.BackColor = COR_BRANCO
        End If
    End If
    
    On Error GoTo 0
End Sub

'====================================================================
' SISTEMA DE DEBUG E LOGS
'====================================================================
Private Sub btnVerificarSistema_Click()
    On Error GoTo TratarErro
    
    Dim relatorio As String
    relatorio = UtilsManager.VerificarIntegridadeSistema()
    
    MsgBox relatorio, vbInformation, "Verificação do Sistema"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnVerificarSistema_Click", Err)
End Sub

Private Sub btnExportarLog_Click()
    On Error GoTo TratarErro
    
    Call UtilsManager.ExportarLogErros
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("btnExportarLog_Click", Err)
End Sub

'====================================================================
' EVENTOS ESPECIAIS E ATALHOS
'====================================================================
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo TratarErro
    
    ' Atalhos do teclado
    Select Case KeyCode
        Case vbKeyF1 ' F1 = Ajuda
            Call ExibirAjuda
            
        Case vbKeyF2 ' F2 = Nova Venda
            Call btnNovaVenda_Click
            
        Case vbKeyF3 ' F3 = Pesquisar Produtos
            If CampoExiste("txtPesquisa") Then Me.txtPesquisa.SetFocus
            
        Case vbKeyF5 ' F5 = Atualizar
            Call AtualizarFormulario
            
        Case vbKeyF12 ' F12 = Gerar Pedido
            Call btnGerarPedido_Click
    End Select
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("UserForm_KeyDown", Err)
End Sub

Private Sub ExibirAjuda()
    MsgBox "🆘 AJUDA DO SISTEMA PDV" & vbCrLf & vbCrLf & _
           "ATALHOS DO TECLADO:" & vbCrLf & _
           "F1 = Esta ajuda" & vbCrLf & _
           "F2 = Nova venda" & vbCrLf & _
           "F3 = Focar na pesquisa" & vbCrLf & _
           "F5 = Atualizar dados" & vbCrLf & _
           "F12 = Gerar pedido" & vbCrLf & vbCrLf & _
           "FLUXO DE USO:" & vbCrLf & _
           "1. Selecionar cliente" & vbCrLf & _
           "2. Pesquisar e adicionar produtos" & vbCrLf & _
           "3. Aplicar descontos (opcional)" & vbCrLf & _
           "4. Gerar pedido" & vbCrLf & _
           "5. Imprimir ou exportar", _
           vbInformation, "Ajuda do Sistema"
End Sub

Private Sub AtualizarFormulario()
    On Error GoTo TratarErro
    
    ' Recarregar dados
    Call ClienteManager.CarregarClientes(Me.cmbCliente)
    Call ProdutoManager.CarregarProdutos(Me.lstProdutos)
    
    ' Atualizar totais
    If Me.produtosv1.ListCount > 0 Then
        Call CalculadoraManager.AtualizarTotais(Me, Me.produtosv1)
    End If
    
    MsgBox "✅ Dados atualizados!", vbInformation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AtualizarFormulario", Err)
End Sub

'====================================================================
' FIM DO CÓDIGO - VERSÃO INTEGRADA COMPLETA
'
' ✅ INTEGRAÇÃO COMPLETA COM TODOS OS MÓDULOS
' ✅ VALIDAÇÕES AVANÇADAS
' ✅ INTERFACE MODERNA E RESPONSIVA
' ✅ TRATAMENTO DE ERROS ROBUSTO
' ✅ COMPATIBILIDADE COM FORMULÁRIO EXISTENTE
' ✅ ATALHOS DE TECLADO
' ✅ SISTEMA DE CORES INTELIGENTE
' ✅ LOGS DETALHADOS
' ✅ PERFORMANCE OTIMIZADA
'
' Data de Criação: 2025-01-27
' Versão: INTEGRADA v2.0
' Status: PRONTO PARA INTEGRAÇÃO
'====================================================================