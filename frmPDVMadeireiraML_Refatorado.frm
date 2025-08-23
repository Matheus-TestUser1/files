'====================================================================
' SISTEMA PDV MADEIREIRA MARIA LUZIA - VERSÃO REFATORADA PROFISSIONAL
' UserForm principal com arquitetura modular e padrões profissionais
' Data/Hora: 2025-01-27
' Versão: PROFISSIONAL v3.0 - ARQUITETURA MODULAR AVANÇADA
' Desenvolvedor: Sistema PDV Enterprise
'====================================================================

Option Explicit

' ===== CONSTANTES DO SISTEMA =====
Private Const SYSTEM_VERSION As String = "PROFESSIONAL v3.0"
Private Const SYSTEM_NAME As String = "PDV Madeireira Maria Luzia"
Private Const DEVELOPER As String = "Sistema PDV Enterprise"
Private Const COPYRIGHT As String = "© 2025 - Todos os direitos reservados"

' ===== ENUMERAÇÕES =====
Private Enum OrderStatus
    Draft = 0
    Pending = 1
    Confirmed = 2
    Completed = 3
    Cancelled = 4
End Enum

Private Enum PaymentMethod
    Cash = 1
    CreditCard = 2
    DebitCard = 3
    BankTransfer = 4
    Check = 5
End Enum

' ===== ESTRUTURAS DE DADOS =====
Private Type CustomerData
    ID As Long
    Name As String
    Document As String
    Address As String
    City As String
    State As String
    ZipCode As String
    Phone As String
    Email As String
End Type

Private Type OrderData
    OrderNumber As String
    CustomerID As Long
    OrderDate As Date
    DeliveryDate As Date
    PaymentMethod As PaymentMethod
    SubTotal As Currency
    DiscountAmount As Currency
    TaxAmount As Currency
    TotalAmount As Currency
    Status As OrderStatus
    Notes As String
End Type

Private Type ProductItem
    ProductID As String
    Description As String
    UnitPrice As Currency
    Quantity As Long
    DiscountPercent As Double
    TotalPrice As Currency
End Type

' ===== VARIÁVEIS DE INSTÂNCIA =====
Private mCurrentOrder As OrderData
Private mCurrentCustomer As CustomerData
Private mProductList As Collection
Private mIsInitialized As Boolean
Private mIsOrderModified As Boolean

' ===== VARIÁVEIS DE INTERFACE =====
Private mFormColors As Dictionary
Private mValidationRules As Dictionary

'====================================================================
' INICIALIZAÇÃO DO SISTEMA
'====================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Inicialização sequencial e controlada
    Call InitializeSystem
    Call InitializeInterface
    Call LoadInitialData
    Call SetupEventHandlers
    
    mIsInitialized = True
    Application.ScreenUpdating = True
    
    ' Log de inicialização bem-sucedida
    Call LogManager.LogInfo("Sistema inicializado com sucesso", "UserForm_Initialize")
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Call HandleCriticalError("UserForm_Initialize", Err)
End Sub

'====================================================================
' INICIALIZAÇÃO DO SISTEMA
'====================================================================
Private Sub InitializeSystem()
    ' Inicializar estruturas de dados
    Set mProductList = New Collection
    Set mFormColors = New Dictionary
    Set mValidationRules = New Dictionary
    
    ' Configurar cores do sistema
    Call SetupSystemColors
    
    ' Configurar regras de validação
    Call SetupValidationRules
    
    ' Inicializar dados da ordem atual
    Call InitializeCurrentOrder
End Sub

Private Sub InitializeCurrentOrder()
    With mCurrentOrder
        .OrderNumber = ""
        .CustomerID = 0
        .OrderDate = Date
        .DeliveryDate = Date + 7
        .PaymentMethod = PaymentMethod.Cash
        .SubTotal = 0
        .DiscountAmount = 0
        .TaxAmount = 0
        .TotalAmount = 0
        .Status = OrderStatus.Draft
        .Notes = ""
    End With
    
    mIsOrderModified = False
End Sub

Private Sub SetupSystemColors()
    With mFormColors
        .Add "Primary", RGB(0, 123, 255)      ' Azul principal
        .Add "Success", RGB(40, 167, 69)      ' Verde sucesso
        .Add "Warning", RGB(255, 193, 7)      ' Amarelo aviso
        .Add "Danger", RGB(220, 53, 69)       ' Vermelho erro
        .Add "Info", RGB(23, 162, 184)        ' Azul informação
        .Add "Light", RGB(248, 249, 250)      ' Cinza claro
        .Add "Dark", RGB(52, 58, 64)          ' Cinza escuro
        .Add "White", RGB(255, 255, 255)      ' Branco
    End With
End Sub

Private Sub SetupValidationRules()
    With mValidationRules
        .Add "CustomerRequired", True
        .Add "ProductsRequired", True
        .Add "PaymentMethodRequired", True
        .Add "MinOrderValue", 0.01
        .Add "MaxDiscountPercent", 50
        .Add "MaxOrderValue", 999999.99
    End With
End Sub

'====================================================================
' INICIALIZAÇÃO DA INTERFACE
'====================================================================
Private Sub InitializeInterface()
    ' Configurar aparência do formulário
    Call ConfigureFormAppearance
    
    ' Configurar controles
    Call ConfigureControls
    
    ' Aplicar tema visual
    Call ApplyVisualTheme
End Sub

Private Sub ConfigureFormAppearance()
    With Me
        .Caption = SYSTEM_NAME & " - " & SYSTEM_VERSION
        .Width = 1200
        .Height = 800
        .StartUpPosition = 0 ' Manual
        .Left = Application.Left + (Application.Width - .Width) / 2
        .Top = Application.Top + (Application.Height - .Height) / 2
    End With
End Sub

Private Sub ConfigureControls()
    ' Configurar ComboBox de clientes
    With Me.cmbCliente
        .Clear
        .AddItem "Selecione um cliente..."
        .ListIndex = 0
    End With
    
    ' Configurar ListBox de produtos
    With Me.lstProdutos
        .Clear
        .ColumnCount = 4
        .ColumnWidths = "100;300;100;100"
    End With
    
    ' Configurar ListBox de produtos selecionados
    With Me.produtosv1
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "100;300;80;80;100"
    End With
    
    ' Configurar ComboBox de pagamento
    With Me.cPagamento
        .Clear
        .AddItem "Dinheiro"
        .AddItem "Cartão de Crédito"
        .AddItem "Cartão de Débito"
        .AddItem "Transferência Bancária"
        .AddItem "Cheque"
        .ListIndex = 0
    End With
End Sub

Private Sub ApplyVisualTheme()
    ' Aplicar cores aos controles principais
    Me.cmbCliente.BackColor = mFormColors("White")
    Me.lstProdutos.BackColor = mFormColors("White")
    Me.produtosv1.BackColor = mFormColors("White")
    
    ' Configurar botões
    Call StyleButton(Me.btnGerarPedido, "Primary")
    Call StyleButton(Me.btnImprimir, "Success")
    Call StyleButton(Me.btnLimpar, "Warning")
    Call StyleButton(Me.btnCancelar, "Danger")
End Sub

Private Sub StyleButton(btn As MSForms.CommandButton, colorKey As String)
    btn.BackColor = mFormColors(colorKey)
    btn.ForeColor = mFormColors("White")
    btn.Font.Bold = True
End Sub

'====================================================================
' CARREGAMENTO DE DADOS INICIAIS
'====================================================================
Private Sub LoadInitialData()
    ' Carregar clientes
    Call CustomerManager.LoadCustomers(Me.cmbCliente)
    
    ' Carregar produtos
    Call ProductManager.LoadProducts(Me.lstProdutos)
    
    ' Carregar configurações do sistema
    Call LoadSystemSettings
    
    ' Atualizar dashboard
    Call DashboardManager.UpdateDashboard
End Sub

Private Sub LoadSystemSettings()
    ' Carregar configurações do arquivo de configuração
    On Error Resume Next
    
    ' Aqui você pode carregar configurações específicas
    ' como taxas, descontos padrão, etc.
    
    On Error GoTo 0
End Sub

'====================================================================
' CONFIGURAÇÃO DE EVENTOS
'====================================================================
Private Sub SetupEventHandlers()
    ' Configurar eventos de validação
    Call SetupValidationEvents
    
    ' Configurar eventos de interface
    Call SetupInterfaceEvents
End Sub

Private Sub SetupValidationEvents()
    ' Eventos serão configurados nos respectivos controles
End Sub

Private Sub SetupInterfaceEvents()
    ' Eventos serão configurados nos respectivos controles
End Sub

'====================================================================
' EVENTOS DE CLIENTE
'====================================================================
Private Sub cmbCliente_Change()
    On Error GoTo ErrorHandler
    
    If Not mIsInitialized Then Exit Sub
    
    If Me.cmbCliente.ListIndex > 0 Then
        Call LoadCustomerData
        Call UpdateCustomerInterface
        Call ValidateOrder
    Else
        Call ClearCustomerData
    End If
    
    Exit Sub

ErrorHandler:
    Call HandleError("cmbCliente_Change", Err)
End Sub

Private Sub LoadCustomerData()
    Dim customerID As Long
    customerID = CLng(Me.cmbCliente.List(Me.cmbCliente.ListIndex, 1))
    
    ' Carregar dados do cliente usando CustomerManager
    mCurrentCustomer = CustomerManager.GetCustomerByID(customerID)
    
    ' Atualizar ordem atual
    mCurrentOrder.CustomerID = customerID
    
    ' Marcar como modificado
    mIsOrderModified = True
End Sub

Private Sub UpdateCustomerInterface()
    ' Preencher campos do cliente se existirem
    If ControlExists("txtNome") Then Me.txtNome.Value = mCurrentCustomer.Name
    If ControlExists("txtEndereco") Then Me.txtEndereco.Value = mCurrentCustomer.Address
    If ControlExists("txtCidade") Then Me.txtCidade.Value = mCurrentCustomer.City
    If ControlExists("txtTelefone") Then Me.txtTelefone.Value = mCurrentCustomer.Phone
End Sub

Private Sub ClearCustomerData()
    ' Limpar dados do cliente
    mCurrentCustomer.ID = 0
    mCurrentCustomer.Name = ""
    mCurrentCustomer.Document = ""
    mCurrentCustomer.Address = ""
    mCurrentCustomer.City = ""
    mCurrentCustomer.State = ""
    mCurrentCustomer.ZipCode = ""
    mCurrentCustomer.Phone = ""
    mCurrentCustomer.Email = ""
    
    ' Limpar interface
    If ControlExists("txtNome") Then Me.txtNome.Value = ""
    If ControlExists("txtEndereco") Then Me.txtEndereco.Value = ""
    If ControlExists("txtCidade") Then Me.txtCidade.Value = ""
    If ControlExists("txtTelefone") Then Me.txtTelefone.Value = ""
End Sub

'====================================================================
' EVENTOS DE PRODUTOS
'====================================================================
Private Sub lstProdutos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    
    If Me.lstProdutos.ListIndex >= 0 Then
        Call AddProductToOrder
    End If
    
    Exit Sub

ErrorHandler:
    Call HandleError("lstProdutos_DblClick", Err)
End Sub

Private Sub AddProductToOrder()
    Dim productID As String
    Dim quantity As Long
    
    ' Obter ID do produto selecionado
    productID = Me.lstProdutos.List(Me.lstProdutos.ListIndex, 0)
    
    ' Solicitar quantidade
    quantity = GetProductQuantity(productID)
    
    If quantity > 0 Then
        ' Adicionar produto à ordem
        Call ProductManager.AddProductToOrder(Me.produtosv1, productID, quantity)
        
        ' Atualizar totais
        Call UpdateOrderTotals
        
        ' Marcar como modificado
        mIsOrderModified = True
        
        ' Validar ordem
        Call ValidateOrder
    End If
End Sub

Private Function GetProductQuantity(productID As String) As Long
    Dim quantity As String
    Dim productName As String
    
    ' Obter nome do produto para exibição
    productName = ProductManager.GetProductName(productID)
    
    ' Solicitar quantidade
    quantity = InputBox("Digite a quantidade para '" & productName & "':", "Quantidade", "1")
    
    If quantity = "" Then
        GetProductQuantity = 0
    ElseIf IsNumeric(quantity) And CLng(quantity) > 0 Then
        GetProductQuantity = CLng(quantity)
    Else
        MsgBox "Por favor, digite uma quantidade válida.", vbExclamation, "Quantidade Inválida"
        GetProductQuantity = 0
    End If
End Function

Private Sub produtosv1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    
    If Me.produtosv1.ListIndex >= 0 Then
        Call RemoveProductFromOrder
    End If
    
    Exit Sub

ErrorHandler:
    Call HandleError("produtosv1_DblClick", Err)
End Sub

Private Sub RemoveProductFromOrder()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("Deseja remover este produto da ordem?", vbQuestion + vbYesNo, "Remover Produto")
    
    If response = vbYes Then
        ' Remover produto da ordem
        Call ProductManager.RemoveProductFromOrder(Me.produtosv1, Me.produtosv1.ListIndex)
        
        ' Atualizar totais
        Call UpdateOrderTotals
        
        ' Marcar como modificado
        mIsOrderModified = True
        
        ' Validar ordem
        Call ValidateOrder
    End If
End Sub

'====================================================================
' ATUALIZAÇÃO DE TOTAIS
'====================================================================
Private Sub UpdateOrderTotals()
    ' Calcular totais usando CalculadoraManager
    Dim totals As Dictionary
    Set totals = CalculadoraManager.CalculateOrderTotals(Me.produtosv1)
    
    ' Atualizar variáveis da ordem
    mCurrentOrder.SubTotal = totals("SubTotal")
    mCurrentOrder.DiscountAmount = totals("Discount")
    mCurrentOrder.TaxAmount = totals("Tax")
    mCurrentOrder.TotalAmount = totals("Total")
    
    ' Atualizar interface
    Call UpdateTotalsInterface(totals)
End Sub

Private Sub UpdateTotalsInterface(totals As Dictionary)
    ' Atualizar campos de totais se existirem
    If ControlExists("txtSubTotal") Then Me.txtSubTotal.Value = Format(totals("SubTotal"), "R$ #,##0.00")
    If ControlExists("txtDesconto") Then Me.txtDesconto.Value = Format(totals("Discount"), "R$ #,##0.00")
    If ControlExists("txtTotal") Then Me.txtTotal.Value = Format(totals("Total"), "R$ #,##0.00")
End Sub

'====================================================================
' VALIDAÇÃO DE ORDEM
'====================================================================
Private Sub ValidateOrder()
    Dim isValid As Boolean
    Dim errorMessage As String
    
    ' Validar ordem usando ValidationManager
    isValid = ValidationManager.ValidateOrder(mCurrentOrder, mCurrentCustomer, Me.produtosv1, errorMessage)
    
    ' Atualizar interface baseado na validação
    Call UpdateValidationInterface(isValid, errorMessage)
End Sub

Private Sub UpdateValidationInterface(isValid As Boolean, errorMessage As String)
    ' Habilitar/desabilitar botão de gerar pedido
    Me.btnGerarPedido.Enabled = isValid
    
    ' Atualizar cor do botão
    If isValid Then
        Call StyleButton(Me.btnGerarPedido, "Primary")
    Else
        Call StyleButton(Me.btnGerarPedido, "Light")
    End If
    
    ' Exibir mensagem de erro se houver
    If Not isValid And errorMessage <> "" Then
        Call ShowValidationMessage(errorMessage)
    End If
End Sub

Private Sub ShowValidationMessage(message As String)
    ' Exibir mensagem de validação em um label ou status bar
    If ControlExists("lblStatus") Then
        Me.lblStatus.Caption = "⚠️ " & message
        Me.lblStatus.ForeColor = mFormColors("Warning")
    End If
End Sub

'====================================================================
' EVENTOS PRINCIPAIS
'====================================================================
Private Sub btnGerarPedido_Click()
    On Error GoTo ErrorHandler
    
    ' Validar ordem antes de gerar
    If Not ValidationManager.ValidateOrder(mCurrentOrder, mCurrentCustomer, Me.produtosv1) Then
        MsgBox "Por favor, corrija os erros antes de gerar o pedido.", vbExclamation, "Validação Necessária"
        Exit Sub
    End If
    
    ' Gerar pedido usando OrderManager
    Dim orderNumber As String
    orderNumber = OrderManager.GenerateOrder(mCurrentOrder, mCurrentCustomer, Me.produtosv1)
    
    If orderNumber <> "" Then
        ' Atualizar interface
        Call UpdateOrderSuccess(orderNumber)
        
        ' Log de sucesso
        Call LogManager.LogInfo("Pedido gerado com sucesso: " & orderNumber, "btnGerarPedido_Click")
    Else
        MsgBox "Erro ao gerar pedido. Tente novamente.", vbCritical, "Erro"
    End If
    
    Exit Sub

ErrorHandler:
    Call HandleError("btnGerarPedido_Click", Err)
End Sub

Private Sub UpdateOrderSuccess(orderNumber As String)
    ' Atualizar número do pedido
    mCurrentOrder.OrderNumber = orderNumber
    mCurrentOrder.Status = OrderStatus.Confirmed
    
    ' Exibir mensagem de sucesso
    MsgBox "✅ Pedido gerado com sucesso!" & vbCrLf & vbCrLf & _
           "📋 Número: " & orderNumber & vbCrLf & _
           "👤 Cliente: " & mCurrentCustomer.Name & vbCrLf & _
           "💰 Total: " & Format(mCurrentOrder.TotalAmount, "R$ #,##0.00") & vbCrLf & vbCrLf & _
           "O pedido foi salvo e está pronto para impressão!", _
           vbInformation, "Pedido Gerado"
    
    ' Habilitar botões de ação
    Call EnableActionButtons(True)
    
    ' Atualizar dashboard
    Call DashboardManager.UpdateDashboard
    
    ' Marcar como não modificado
    mIsOrderModified = False
End Sub

Private Sub EnableActionButtons(enable As Boolean)
    Me.btnImprimir.Enabled = enable
    Me.btnExportarPDF.Enabled = enable
    Me.btnEnviarEmail.Enabled = enable
End Sub

Private Sub btnImprimir_Click()
    On Error GoTo ErrorHandler
    
    If mCurrentOrder.OrderNumber = "" Then
        MsgBox "Gere um pedido primeiro!", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' Imprimir usando PrintManager
    Call PrintManager.PrintOrder(mCurrentOrder.OrderNumber)
    
    Exit Sub

ErrorHandler:
    Call HandleError("btnImprimir_Click", Err)
End Sub

Private Sub btnLimpar_Click()
    On Error GoTo ErrorHandler
    
    If mIsOrderModified Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Existem alterações não salvas. Deseja realmente limpar?", vbQuestion + vbYesNo, "Confirmar Limpeza")
        
        If response = vbNo Then Exit Sub
    End If
    
    ' Limpar ordem atual
    Call ClearCurrentOrder
    
    ' Limpar interface
    Call ClearInterface
    
    Exit Sub

ErrorHandler:
    Call HandleError("btnLimpar_Click", Err)
End Sub

Private Sub ClearCurrentOrder()
    ' Limpar dados da ordem
    Call InitializeCurrentOrder
    
    ' Limpar lista de produtos
    Me.produtosv1.Clear
    
    ' Desabilitar botões de ação
    Call EnableActionButtons(False)
End Sub

Private Sub ClearInterface()
    ' Limpar campos do cliente
    Me.cmbCliente.ListIndex = 0
    Call ClearCustomerData
    
    ' Limpar totais
    Call UpdateTotalsInterface(CreateEmptyTotals)
    
    ' Limpar mensagens de status
    If ControlExists("lblStatus") Then
        Me.lblStatus.Caption = ""
    End If
End Sub

Private Function CreateEmptyTotals() As Dictionary
    Set CreateEmptyTotals = New Dictionary
    CreateEmptyTotals.Add "SubTotal", 0
    CreateEmptyTotals.Add "Discount", 0
    CreateEmptyTotals.Add "Tax", 0
    CreateEmptyTotals.Add "Total", 0
End Function

'====================================================================
' UTILITÁRIOS
'====================================================================
Private Function ControlExists(controlName As String) As Boolean
    On Error Resume Next
    ControlExists = Not Me.Controls(controlName) Is Nothing
    On Error GoTo 0
End Function

Private Sub HandleError(procedureName As String, err As ErrObject)
    ' Registrar erro
    Call ErrorHandler.LogError(procedureName, err)
    
    ' Exibir mensagem para o usuário
    MsgBox "Ocorreu um erro inesperado. Por favor, tente novamente." & vbCrLf & _
           "Detalhes: " & err.Description, vbCritical, "Erro"
End Sub

Private Sub HandleCriticalError(procedureName As String, err As ErrObject)
    ' Registrar erro crítico
    Call ErrorHandler.LogCriticalError(procedureName, err)
    
    ' Exibir mensagem crítica
    MsgBox "❌ ERRO CRÍTICO NO SISTEMA!" & vbCrLf & vbCrLf & _
           "Erro: " & err.Description & vbCrLf & _
           "Contate o suporte técnico imediatamente.", vbCritical, "Erro Crítico"
End Sub

'====================================================================
' FINALIZAÇÃO DO SISTEMA
'====================================================================
Private Sub UserForm_Terminate()
    ' Salvar configurações se necessário
    Call SaveSystemSettings
    
    ' Log de finalização
    Call LogManager.LogInfo("Sistema finalizado", "UserForm_Terminate")
End Sub

Private Sub SaveSystemSettings()
    ' Salvar configurações do sistema se necessário
    On Error Resume Next
    ' Implementar salvamento de configurações
    On Error GoTo 0
End Sub