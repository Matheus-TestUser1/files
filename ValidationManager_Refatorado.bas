'====================================================================
' MÓDULO VALIDATION MANAGER - VERSÃO REFATORADA PROFISSIONAL
' Responsável por todas as validações do sistema
' Data/Hora: 2025-01-27
' Versão: PROFESSIONAL v3.0
' Desenvolvedor: Sistema PDV Enterprise
'====================================================================

Option Explicit

' ===== CONSTANTES =====
Private Const MIN_ORDER_VALUE As Currency = 0.01
Private Const MAX_ORDER_VALUE As Currency = 999999.99
Private Const MAX_DISCOUNT_PERCENT As Double = 50.0
Private Const MIN_PRODUCT_QUANTITY As Long = 1
Private Const MAX_PRODUCT_QUANTITY As Long = 9999

' ===== ENUMERAÇÕES =====
Private Enum ValidationResult
    Valid = 0
    InvalidCustomer = 1
    InvalidProducts = 2
    InvalidPaymentMethod = 3
    InvalidOrderValue = 4
    InvalidDiscount = 5
    InvalidQuantity = 6
    InvalidDocument = 7
    InvalidEmail = 8
    InvalidPhone = 9
    DuplicateProduct = 10
    InsufficientStock = 11
End Enum

Private Enum ValidationSeverity
    Info = 0
    Warning = 1
    Error = 2
    Critical = 3
End Enum

' ===== ESTRUTURAS DE DADOS =====
Private Type ValidationRule
    FieldName As String
    IsRequired As Boolean
    MinValue As Variant
    MaxValue As Variant
    Pattern As String
    CustomValidation As String
    ErrorMessage As String
    Severity As ValidationSeverity
End Type

Private Type ValidationContext
    OrderData As OrderData
    CustomerData As CustomerData
    ProductList As MSForms.ListBox
    ValidationRules As Collection
End Type

' ===== VARIÁVEIS DE INSTÂNCIA =====
Private mValidationRules As Collection
Private mCustomValidators As Dictionary

'====================================================================
' INICIALIZAÇÃO DO MÓDULO
'====================================================================
Private Sub Class_Initialize()
    Call InitializeValidationRules
    Set mCustomValidators = New Dictionary
    Call RegisterCustomValidators
End Sub

Private Sub InitializeValidationRules()
    Set mValidationRules = New Collection
    
    ' Regras de validação de cliente
    Call AddValidationRule("CustomerName", True, "", "", "", "", "Nome do cliente é obrigatório", ValidationSeverity.Error)
    Call AddValidationRule("CustomerDocument", True, "", "", "", "", "Documento do cliente é obrigatório", ValidationSeverity.Error)
    Call AddValidationRule("CustomerPhone", True, "", "", "", "", "Telefone do cliente é obrigatório", ValidationSeverity.Error)
    
    ' Regras de validação de produtos
    Call AddValidationRule("ProductQuantity", True, MIN_PRODUCT_QUANTITY, MAX_PRODUCT_QUANTITY, "", "", "Quantidade deve estar entre " & MIN_PRODUCT_QUANTITY & " e " & MAX_PRODUCT_QUANTITY, ValidationSeverity.Error)
    Call AddValidationRule("ProductPrice", True, 0.01, 99999.99, "", "", "Preço do produto deve ser maior que zero", ValidationSeverity.Error)
    
    ' Regras de validação de pedido
    Call AddValidationRule("OrderValue", True, MIN_ORDER_VALUE, MAX_ORDER_VALUE, "", "", "Valor do pedido deve estar entre R$ " & Format(MIN_ORDER_VALUE, "0.00") & " e R$ " & Format(MAX_ORDER_VALUE, "0.00"), ValidationSeverity.Error)
    Call AddValidationRule("DiscountPercent", False, 0, MAX_DISCOUNT_PERCENT, "", "", "Desconto não pode exceder " & MAX_DISCOUNT_PERCENT & "%", ValidationSeverity.Warning)
    
    ' Regras de validação de pagamento
    Call AddValidationRule("PaymentMethod", True, "", "", "", "", "Forma de pagamento é obrigatória", ValidationSeverity.Error)
End Sub

Private Sub AddValidationRule(fieldName As String, isRequired As Boolean, minValue As Variant, maxValue As Variant, pattern As String, customValidation As String, errorMessage As String, severity As ValidationSeverity)
    Dim rule As ValidationRule
    
    With rule
        .FieldName = fieldName
        .IsRequired = isRequired
        .MinValue = minValue
        .MaxValue = maxValue
        .Pattern = pattern
        .CustomValidation = customValidation
        .ErrorMessage = errorMessage
        .Severity = severity
    End With
    
    mValidationRules.Add rule, fieldName
End Sub

Private Sub RegisterCustomValidators()
    ' Registrar validadores customizados
    mCustomValidators.Add "Email", "ValidateEmail"
    mCustomValidators.Add "Document", "ValidateDocument"
    mCustomValidators.Add "Phone", "ValidatePhone"
    mCustomValidators.Add "ZipCode", "ValidateZipCode"
End Sub

'====================================================================
' VALIDAÇÕES PRINCIPAIS
'====================================================================

' === VALIDAR PEDIDO COMPLETO ===
Public Function ValidateOrder(orderData As OrderData, customerData As CustomerData, productList As MSForms.ListBox, Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar cliente
    If Not ValidateCustomer(customerData, errorMessage) Then
        ValidateOrder = False
        Exit Function
    End If
    
    ' Validar produtos
    If Not ValidateProducts(productList, errorMessage) Then
        ValidateOrder = False
        Exit Function
    End If
    
    ' Validar valor do pedido
    If Not ValidateOrderValue(orderData, errorMessage) Then
        ValidateOrder = False
        Exit Function
    End If
    
    ' Validar desconto
    If Not ValidateDiscount(orderData, errorMessage) Then
        ValidateOrder = False
        Exit Function
    End If
    
    ' Validar forma de pagamento
    If Not ValidatePaymentMethod(orderData, errorMessage) Then
        ValidateOrder = False
        Exit Function
    End If
    
    ValidateOrder = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("ValidateOrder", Err)
    errorMessage = "Erro durante validação: " & Err.Description
    ValidateOrder = False
End Function

' === VALIDAR CLIENTE ===
Public Function ValidateCustomer(customerData As CustomerData, Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar nome
    If Not ValidateRequiredField(customerData.Name, "Nome do cliente", errorMessage) Then
        ValidateCustomer = False
        Exit Function
    End If
    
    ' Validar documento
    If Not ValidateRequiredField(customerData.Document, "Documento do cliente", errorMessage) Then
        ValidateCustomer = False
        Exit Function
    End If
    
    If Not ValidateDocument(customerData.Document) Then
        errorMessage = "Documento inválido: " & customerData.Document
        ValidateCustomer = False
        Exit Function
    End If
    
    ' Validar telefone
    If Not ValidateRequiredField(customerData.Phone, "Telefone do cliente", errorMessage) Then
        ValidateCustomer = False
        Exit Function
    End If
    
    If Not ValidatePhone(customerData.Phone) Then
        errorMessage = "Telefone inválido: " & customerData.Phone
        ValidateCustomer = False
        Exit Function
    End If
    
    ' Validar email (se fornecido)
    If Trim(customerData.Email) <> "" Then
        If Not ValidateEmail(customerData.Email) Then
            errorMessage = "Email inválido: " & customerData.Email
            ValidateCustomer = False
            Exit Function
        End If
    End If
    
    ValidateCustomer = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("ValidateCustomer", Err)
    errorMessage = "Erro durante validação do cliente: " & Err.Description
    ValidateCustomer = False
End Function

' === VALIDAR PRODUTOS ===
Public Function ValidateProducts(productList As MSForms.ListBox, Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificar se há produtos selecionados
    If productList.ListCount = 0 Then
        errorMessage = "Nenhum produto selecionado para o pedido"
        ValidateProducts = False
        Exit Function
    End If
    
    ' Validar cada produto
    Dim i As Long
    For i = 0 To productList.ListCount - 1
        If Not ValidateProductItem(productList, i, errorMessage) Then
            ValidateProducts = False
            Exit Function
        End If
    Next i
    
    ' Verificar duplicatas
    If HasDuplicateProducts(productList) Then
        errorMessage = "Existem produtos duplicados no pedido"
        ValidateProducts = False
        Exit Function
    End If
    
    ValidateProducts = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("ValidateProducts", Err)
    errorMessage = "Erro durante validação dos produtos: " & Err.Description
    ValidateProducts = False
End Function

Private Function ValidateProductItem(productList As MSForms.ListBox, index As Long, Optional ByRef errorMessage As String = "") As Boolean
    ' Validar quantidade
    Dim quantity As Long
    quantity = CLng(productList.List(index, 2)) ' Assumindo que a quantidade está na coluna 2
    
    If quantity < MIN_PRODUCT_QUANTITY Or quantity > MAX_PRODUCT_QUANTITY Then
        errorMessage = "Quantidade inválida para produto: " & productList.List(index, 1)
        ValidateProductItem = False
        Exit Function
    End If
    
    ' Validar preço
    Dim price As Currency
    price = CCur(productList.List(index, 3)) ' Assumindo que o preço está na coluna 3
    
    If price <= 0 Then
        errorMessage = "Preço inválido para produto: " & productList.List(index, 1)
        ValidateProductItem = False
        Exit Function
    End If
    
    ' Verificar estoque (se implementado)
    If Not CheckProductStock(productList.List(index, 0), quantity) Then
        errorMessage = "Estoque insuficiente para produto: " & productList.List(index, 1)
        ValidateProductItem = False
        Exit Function
    End If
    
    ValidateProductItem = True
End Function

Private Function HasDuplicateProducts(productList As MSForms.ListBox) As Boolean
    Dim i As Long, j As Long
    
    For i = 0 To productList.ListCount - 2
        For j = i + 1 To productList.ListCount - 1
            If productList.List(i, 0) = productList.List(j, 0) Then ' Comparar IDs dos produtos
                HasDuplicateProducts = True
                Exit Function
            End If
        Next j
    Next i
    
    HasDuplicateProducts = False
End Function

Private Function CheckProductStock(productID As String, quantity As Long) As Boolean
    ' Implementar verificação de estoque
    ' Por enquanto, retorna True
    CheckProductStock = True
End Function

' === VALIDAR VALOR DO PEDIDO ===
Public Function ValidateOrderValue(orderData As OrderData, Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar subtotal
    If orderData.SubTotal < MIN_ORDER_VALUE Then
        errorMessage = "Valor mínimo do pedido não atingido"
        ValidateOrderValue = False
        Exit Function
    End If
    
    If orderData.SubTotal > MAX_ORDER_VALUE Then
        errorMessage = "Valor máximo do pedido excedido"
        ValidateOrderValue = False
        Exit Function
    End If
    
    ' Validar total
    If orderData.TotalAmount < MIN_ORDER_VALUE Then
        errorMessage = "Total do pedido deve ser maior que zero"
        ValidateOrderValue = False
        Exit Function
    End If
    
    If orderData.TotalAmount > MAX_ORDER_VALUE Then
        errorMessage = "Total do pedido excede o valor máximo permitido"
        ValidateOrderValue = False
        Exit Function
    End If
    
    ValidateOrderValue = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("ValidateOrderValue", Err)
    errorMessage = "Erro durante validação do valor do pedido: " & Err.Description
    ValidateOrderValue = False
End Function

' === VALIDAR DESCONTO ===
Public Function ValidateDiscount(orderData As OrderData, Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar valor do desconto
    If orderData.DiscountAmount < 0 Then
        errorMessage = "Desconto não pode ser negativo"
        ValidateDiscount = False
        Exit Function
    End If
    
    ' Validar percentual de desconto
    If orderData.SubTotal > 0 Then
        Dim discountPercent As Double
        discountPercent = (orderData.DiscountAmount / orderData.SubTotal) * 100
        
        If discountPercent > MAX_DISCOUNT_PERCENT Then
            errorMessage = "Desconto não pode exceder " & MAX_DISCOUNT_PERCENT & "%"
            ValidateDiscount = False
            Exit Function
        End If
    End If
    
    ValidateDiscount = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("ValidateDiscount", Err)
    errorMessage = "Erro durante validação do desconto: " & Err.Description
    ValidateDiscount = False
End Function

' === VALIDAR FORMA DE PAGAMENTO ===
Public Function ValidatePaymentMethod(orderData As OrderData, Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificar se a forma de pagamento foi selecionada
    If orderData.PaymentMethod < PaymentMethod.Cash Or orderData.PaymentMethod > PaymentMethod.Check Then
        errorMessage = "Forma de pagamento inválida"
        ValidatePaymentMethod = False
        Exit Function
    End If
    
    ValidatePaymentMethod = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("ValidatePaymentMethod", Err)
    errorMessage = "Erro durante validação da forma de pagamento: " & Err.Description
    ValidatePaymentMethod = False
End Function

'====================================================================
' VALIDAÇÕES ESPECÍFICAS
'====================================================================

' === VALIDAR CAMPO OBRIGATÓRIO ===
Public Function ValidateRequiredField(value As String, fieldName As String, Optional ByRef errorMessage As String = "") As Boolean
    If Trim(value) = "" Then
        errorMessage = fieldName & " é obrigatório"
        ValidateRequiredField = False
    Else
        ValidateRequiredField = True
    End If
End Function

' === VALIDAR EMAIL ===
Public Function ValidateEmail(email As String) As Boolean
    ' Validação básica de email
    If InStr(email, "@") = 0 Or InStr(email, ".") = 0 Then
        ValidateEmail = False
        Exit Function
    End If
    
    If Left(email, 1) = "@" Or Right(email, 1) = "@" Then
        ValidateEmail = False
        Exit Function
    End If
    
    If InStr(email, "..") > 0 Or InStr(email, "@@") > 0 Then
        ValidateEmail = False
        Exit Function
    End If
    
    ' Verificar se há pelo menos um caractere antes e depois do @
    Dim atPosition As Long
    atPosition = InStr(email, "@")
    
    If atPosition <= 1 Or atPosition >= Len(email) Then
        ValidateEmail = False
        Exit Function
    End If
    
    ' Verificar se há pelo menos um ponto após o @
    Dim domainPart As String
    domainPart = Mid(email, atPosition + 1)
    
    If InStr(domainPart, ".") = 0 Then
        ValidateEmail = False
        Exit Function
    End If
    
    ValidateEmail = True
End Function

' === VALIDAR DOCUMENTO ===
Public Function ValidateDocument(document As String) As Boolean
    ' Remover caracteres especiais
    document = Replace(document, ".", "")
    document = Replace(document, "-", "")
    document = Replace(document, "/", "")
    
    ' Verificar se é numérico
    If Not IsNumeric(document) Then
        ValidateDocument = False
        Exit Function
    End If
    
    ' Verificar tamanho (CPF = 11, CNPJ = 14)
    If Len(document) <> 11 And Len(document) <> 14 Then
        ValidateDocument = False
        Exit Function
    End If
    
    ' Aqui você pode adicionar validação específica de CPF/CNPJ
    ' Por enquanto, apenas validação básica
    ValidateDocument = True
End Function

' === VALIDAR TELEFONE ===
Public Function ValidatePhone(phone As String) As Boolean
    ' Remover caracteres especiais
    phone = Replace(phone, "(", "")
    phone = Replace(phone, ")", "")
    phone = Replace(phone, "-", "")
    phone = Replace(phone, " ", "")
    
    ' Verificar se é numérico
    If Not IsNumeric(phone) Then
        ValidatePhone = False
        Exit Function
    End If
    
    ' Verificar tamanho (mínimo 10, máximo 11 dígitos)
    If Len(phone) < 10 Or Len(phone) > 11 Then
        ValidatePhone = False
        Exit Function
    End If
    
    ValidatePhone = True
End Function

' === VALIDAR CEP ===
Public Function ValidateZipCode(zipCode As String) As Boolean
    ' Remover caracteres especiais
    zipCode = Replace(zipCode, "-", "")
    zipCode = Replace(zipCode, ".", "")
    
    ' Verificar se é numérico
    If Not IsNumeric(zipCode) Then
        ValidateZipCode = False
        Exit Function
    End If
    
    ' Verificar tamanho (8 dígitos)
    If Len(zipCode) <> 8 Then
        ValidateZipCode = False
        Exit Function
    End If
    
    ValidateZipCode = True
End Function

' === VALIDAR VALOR NUMÉRICO ===
Public Function ValidateNumericValue(value As Variant, minValue As Double, maxValue As Double, fieldName As String, Optional ByRef errorMessage As String = "") As Boolean
    ' Verificar se é numérico
    If Not IsNumeric(value) Then
        errorMessage = fieldName & " deve ser um valor numérico"
        ValidateNumericValue = False
        Exit Function
    End If
    
    ' Verificar limites
    If CDbl(value) < minValue Or CDbl(value) > maxValue Then
        errorMessage = fieldName & " deve estar entre " & minValue & " e " & maxValue
        ValidateNumericValue = False
        Exit Function
    End If
    
    ValidateNumericValue = True
End Function

' === VALIDAR DATA ===
Public Function ValidateDate(dateValue As Variant, fieldName As String, Optional ByRef errorMessage As String = "") As Boolean
    ' Verificar se é uma data válida
    If Not IsDate(dateValue) Then
        errorMessage = fieldName & " deve ser uma data válida"
        ValidateDate = False
        Exit Function
    End If
    
    ' Verificar se não é uma data futura muito distante
    If CDate(dateValue) > Date + 365 Then
        errorMessage = fieldName & " não pode ser uma data muito futura"
        ValidateDate = False
        Exit Function
    End If
    
    ValidateDate = True
End Function

'====================================================================
' VALIDAÇÕES DE NEGÓCIO
'====================================================================

' === VALIDAR REGRAS DE NEGÓCIO ===
Public Function ValidateBusinessRules(orderData As OrderData, customerData As CustomerData, productList As MSForms.ListBox, Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar limite de crédito do cliente (se implementado)
    If Not ValidateCustomerCreditLimit(customerData, orderData.TotalAmount, errorMessage) Then
        ValidateBusinessRules = False
        Exit Function
    End If
    
    ' Validar horário de funcionamento (se implementado)
    If Not ValidateBusinessHours(errorMessage) Then
        ValidateBusinessRules = False
        Exit Function
    End If
    
    ' Validar produtos sazonais (se implementado)
    If Not ValidateSeasonalProducts(productList, errorMessage) Then
        ValidateBusinessRules = False
        Exit Function
    End If
    
    ValidateBusinessRules = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("ValidateBusinessRules", Err)
    errorMessage = "Erro durante validação das regras de negócio: " & Err.Description
    ValidateBusinessRules = False
End Function

Private Function ValidateCustomerCreditLimit(customerData As CustomerData, orderValue As Currency, Optional ByRef errorMessage As String = "") As Boolean
    ' Implementar validação de limite de crédito
    ' Por enquanto, retorna True
    ValidateCustomerCreditLimit = True
End Function

Private Function ValidateBusinessHours(Optional ByRef errorMessage As String = "") As Boolean
    ' Implementar validação de horário de funcionamento
    ' Por enquanto, retorna True
    ValidateBusinessHours = True
End Function

Private Function ValidateSeasonalProducts(productList As MSForms.ListBox, Optional ByRef errorMessage As String = "") As Boolean
    ' Implementar validação de produtos sazonais
    ' Por enquanto, retorna True
    ValidateSeasonalProducts = True
End Function

'====================================================================
' UTILITÁRIOS
'====================================================================

' === OBTER MENSAGEM DE ERRO ===
Public Function GetErrorMessage(validationResult As ValidationResult) As String
    Select Case validationResult
        Case ValidationResult.InvalidCustomer
            GetErrorMessage = "Dados do cliente inválidos"
        Case ValidationResult.InvalidProducts
            GetErrorMessage = "Produtos inválidos ou não selecionados"
        Case ValidationResult.InvalidPaymentMethod
            GetErrorMessage = "Forma de pagamento inválida"
        Case ValidationResult.InvalidOrderValue
            GetErrorMessage = "Valor do pedido inválido"
        Case ValidationResult.InvalidDiscount
            GetErrorMessage = "Desconto inválido"
        Case ValidationResult.InvalidQuantity
            GetErrorMessage = "Quantidade inválida"
        Case ValidationResult.InvalidDocument
            GetErrorMessage = "Documento inválido"
        Case ValidationResult.InvalidEmail
            GetErrorMessage = "Email inválido"
        Case ValidationResult.InvalidPhone
            GetErrorMessage = "Telefone inválido"
        Case ValidationResult.DuplicateProduct
            GetErrorMessage = "Produto duplicado no pedido"
        Case ValidationResult.InsufficientStock
            GetErrorMessage = "Estoque insuficiente"
        Case Else
            GetErrorMessage = "Erro de validação desconhecido"
    End Select
End Function

' === VALIDAR CAMPO COM REGRA ESPECÍFICA ===
Public Function ValidateFieldWithRule(value As Variant, ruleName As String, Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Obter regra de validação
    Dim rule As ValidationRule
    Set rule = mValidationRules(ruleName)
    
    ' Validar campo obrigatório
    If rule.IsRequired Then
        If IsEmpty(value) Or Trim(CStr(value)) = "" Then
            errorMessage = rule.ErrorMessage
            ValidateFieldWithRule = False
            Exit Function
        End If
    End If
    
    ' Validar valor mínimo
    If Not IsEmpty(rule.MinValue) Then
        If value < rule.MinValue Then
            errorMessage = rule.ErrorMessage
            ValidateFieldWithRule = False
            Exit Function
        End If
    End If
    
    ' Validar valor máximo
    If Not IsEmpty(rule.MaxValue) Then
        If value > rule.MaxValue Then
            errorMessage = rule.ErrorMessage
            ValidateFieldWithRule = False
            Exit Function
        End If
    End If
    
    ' Validar padrão (regex)
    If rule.Pattern <> "" Then
        If Not ValidatePattern(CStr(value), rule.Pattern) Then
            errorMessage = rule.ErrorMessage
            ValidateFieldWithRule = False
            Exit Function
        End If
    End If
    
    ' Validação customizada
    If rule.CustomValidation <> "" Then
        If mCustomValidators.Exists(rule.CustomValidation) Then
            Dim validatorName As String
            validatorName = mCustomValidators(rule.CustomValidation)
            
            ' Chamar validador customizado
            If Not CallCustomValidator(validatorName, value) Then
                errorMessage = rule.ErrorMessage
                ValidateFieldWithRule = False
                Exit Function
            End If
        End If
    End If
    
    ValidateFieldWithRule = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("ValidateFieldWithRule", Err)
    errorMessage = "Erro durante validação do campo: " & Err.Description
    ValidateFieldWithRule = False
End Function

Private Function ValidatePattern(value As String, pattern As String) As Boolean
    ' Implementar validação de padrão (regex)
    ' Por enquanto, retorna True
    ValidatePattern = True
End Function

Private Function CallCustomValidator(validatorName As String, value As Variant) As Boolean
    ' Implementar chamada de validador customizado
    ' Por enquanto, retorna True
    CallCustomValidator = True
End Function

' === LIMPAR MENSAGENS DE ERRO ===
Public Sub ClearValidationMessages()
    ' Implementar limpeza de mensagens de validação
    ' Por enquanto, apenas log
    Call LogManager.LogInfo("Mensagens de validação limpas", "ClearValidationMessages")
End Sub