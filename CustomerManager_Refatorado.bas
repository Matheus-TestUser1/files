'====================================================================
' MÓDULO CUSTOMER MANAGER - VERSÃO REFATORADA PROFISSIONAL
' Responsável por todas as operações relacionadas aos clientes
' Data/Hora: 2025-01-27
' Versão: PROFESSIONAL v3.0
' Desenvolvedor: Sistema PDV Enterprise
'====================================================================

Option Explicit

' ===== CONSTANTES =====
Private Const CUSTOMER_SHEET_NAME As String = "Clientes"
Private Const CUSTOMER_ID_COLUMN As String = "A"
Private Const CUSTOMER_NAME_COLUMN As String = "B"
Private Const CUSTOMER_DOCUMENT_COLUMN As String = "C"
Private Const CUSTOMER_ADDRESS_COLUMN As String = "D"
Private Const CUSTOMER_CITY_COLUMN As String = "E"
Private Const CUSTOMER_STATE_COLUMN As String = "F"
Private Const CUSTOMER_ZIPCODE_COLUMN As String = "G"
Private Const CUSTOMER_PHONE_COLUMN As String = "H"
Private Const CUSTOMER_EMAIL_COLUMN As String = "I"

' ===== ENUMERAÇÕES =====
Private Enum CustomerValidationResult
    Valid = 0
    InvalidName = 1
    InvalidDocument = 2
    InvalidEmail = 3
    InvalidPhone = 4
    DuplicateDocument = 5
End Enum

' ===== ESTRUTURAS DE DADOS =====
Private Type CustomerSearchCriteria
    Name As String
    Document As String
    City As String
    Phone As String
    Email As String
End Type

Private Type CustomerValidationRules
    RequireName As Boolean
    RequireDocument As Boolean
    RequireAddress As Boolean
    RequirePhone As Boolean
    ValidateEmail As Boolean
    ValidateDocument As Boolean
End Type

' ===== VARIÁVEIS DE INSTÂNCIA =====
Private mValidationRules As CustomerValidationRules
Private mCustomerCache As Dictionary

'====================================================================
' INICIALIZAÇÃO DO MÓDULO
'====================================================================
Private Sub Class_Initialize()
    Call InitializeValidationRules
    Set mCustomerCache = New Dictionary
End Sub

Private Sub InitializeValidationRules()
    With mValidationRules
        .RequireName = True
        .RequireDocument = True
        .RequireAddress = True
        .RequirePhone = True
        .ValidateEmail = True
        .ValidateDocument = True
    End With
End Sub

'====================================================================
' OPERAÇÕES PRINCIPAIS
'====================================================================

' === CARREGAR CLIENTES ===
Public Sub LoadCustomers(comboBox As MSForms.ComboBox)
    On Error GoTo ErrorHandler
    
    ' Validar parâmetro
    If comboBox Is Nothing Then
        Call LogManager.LogError("LoadCustomers", "ComboBox parameter is null")
        Exit Sub
    End If
    
    ' Limpar cache se necessário
    Call ClearCustomerCache
    
    ' Obter worksheet
    Dim ws As Worksheet
    Set ws = GetCustomerWorksheet
    
    If ws Is Nothing Then
        Call LogManager.LogError("LoadCustomers", "Customer worksheet not found")
        Exit Sub
    End If
    
    ' Limpar ComboBox
    comboBox.Clear
    comboBox.AddItem "Selecione um cliente..."
    
    ' Carregar clientes
    Call LoadCustomersFromWorksheet(ws, comboBox)
    
    ' Definir índice padrão
    comboBox.ListIndex = 0
    
    ' Log de sucesso
    Call LogManager.LogInfo("Clientes carregados com sucesso", "LoadCustomers")
    
    Exit Sub

ErrorHandler:
    Call LogManager.LogError("LoadCustomers", Err)
    Call ShowErrorMessage("Erro ao carregar clientes: " & Err.Description)
End Sub

Private Sub LoadCustomersFromWorksheet(ws As Worksheet, comboBox As MSForms.ComboBox)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, CUSTOMER_ID_COLUMN)
    
    If lastRow < 2 Then
        Call LogManager.LogWarning("LoadCustomersFromWorksheet", "No customers found in worksheet")
        Exit Sub
    End If
    
    Dim i As Long
    For i = 2 To lastRow
        If IsValidCustomerRow(ws, i) Then
            Call AddCustomerToComboBox(ws, i, comboBox)
        End If
    Next i
End Sub

Private Function IsValidCustomerRow(ws As Worksheet, row As Long) As Boolean
    ' Verificar se a linha contém dados válidos
    Dim name As String
    name = Trim(CStr(ws.Cells(row, CUSTOMER_NAME_COLUMN).Value))
    
    IsValidCustomerRow = (name <> "" And name <> "N/A" And name <> "NULL")
End Function

Private Sub AddCustomerToComboBox(ws As Worksheet, row As Long, comboBox As MSForms.ComboBox)
    Dim customerID As Long
    Dim customerName As String
    
    customerID = CLng(ws.Cells(row, CUSTOMER_ID_COLUMN).Value)
    customerName = Trim(CStr(ws.Cells(row, CUSTOMER_NAME_COLUMN).Value))
    
    ' Adicionar ao ComboBox
    comboBox.AddItem customerName
    comboBox.List(comboBox.ListCount - 1, 1) = customerID
    
    ' Adicionar ao cache
    Call AddCustomerToCache(ws, row)
End Sub

Private Sub AddCustomerToCache(ws As Worksheet, row As Long)
    Dim customer As CustomerData
    
    With customer
        .ID = CLng(ws.Cells(row, CUSTOMER_ID_COLUMN).Value)
        .Name = Trim(CStr(ws.Cells(row, CUSTOMER_NAME_COLUMN).Value))
        .Document = Trim(CStr(ws.Cells(row, CUSTOMER_DOCUMENT_COLUMN).Value))
        .Address = Trim(CStr(ws.Cells(row, CUSTOMER_ADDRESS_COLUMN).Value))
        .City = Trim(CStr(ws.Cells(row, CUSTOMER_CITY_COLUMN).Value))
        .State = Trim(CStr(ws.Cells(row, CUSTOMER_STATE_COLUMN).Value))
        .ZipCode = Trim(CStr(ws.Cells(row, CUSTOMER_ZIPCODE_COLUMN).Value))
        .Phone = Trim(CStr(ws.Cells(row, CUSTOMER_PHONE_COLUMN).Value))
        .Email = Trim(CStr(ws.Cells(row, CUSTOMER_EMAIL_COLUMN).Value))
    End With
    
    mCustomerCache.Add customer.ID, customer
End Sub

' === OBTER CLIENTE POR ID ===
Public Function GetCustomerByID(customerID As Long) As CustomerData
    On Error GoTo ErrorHandler
    
    ' Verificar cache primeiro
    If mCustomerCache.Exists(customerID) Then
        GetCustomerByID = mCustomerCache(customerID)
        Exit Function
    End If
    
    ' Buscar no worksheet
    Dim ws As Worksheet
    Set ws = GetCustomerWorksheet
    
    If ws Is Nothing Then
        Call LogManager.LogError("GetCustomerByID", "Customer worksheet not found")
        Exit Function
    End If
    
    GetCustomerByID = FindCustomerInWorksheet(ws, customerID)
    
    Exit Function

ErrorHandler:
    Call LogManager.LogError("GetCustomerByID", Err)
    Call ShowErrorMessage("Erro ao buscar cliente: " & Err.Description)
End Function

Private Function FindCustomerInWorksheet(ws As Worksheet, customerID As Long) As CustomerData
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, CUSTOMER_ID_COLUMN)
    
    Dim i As Long
    For i = 2 To lastRow
        If CLng(ws.Cells(i, CUSTOMER_ID_COLUMN).Value) = customerID Then
            FindCustomerInWorksheet = GetCustomerFromRow(ws, i)
            Exit Function
        End If
    Next i
    
    ' Cliente não encontrado
    Call LogManager.LogWarning("FindCustomerInWorksheet", "Customer not found: " & customerID)
End Function

Private Function GetCustomerFromRow(ws As Worksheet, row As Long) As CustomerData
    With GetCustomerFromRow
        .ID = CLng(ws.Cells(row, CUSTOMER_ID_COLUMN).Value)
        .Name = Trim(CStr(ws.Cells(row, CUSTOMER_NAME_COLUMN).Value))
        .Document = Trim(CStr(ws.Cells(row, CUSTOMER_DOCUMENT_COLUMN).Value))
        .Address = Trim(CStr(ws.Cells(row, CUSTOMER_ADDRESS_COLUMN).Value))
        .City = Trim(CStr(ws.Cells(row, CUSTOMER_CITY_COLUMN).Value))
        .State = Trim(CStr(ws.Cells(row, CUSTOMER_STATE_COLUMN).Value))
        .ZipCode = Trim(CStr(ws.Cells(row, CUSTOMER_ZIPCODE_COLUMN).Value))
        .Phone = Trim(CStr(ws.Cells(row, CUSTOMER_PHONE_COLUMN).Value))
        .Email = Trim(CStr(ws.Cells(row, CUSTOMER_EMAIL_COLUMN).Value))
    End With
End Function

' === CADASTRAR NOVO CLIENTE ===
Public Function RegisterNewCustomer(customer As CustomerData) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar dados do cliente
    Dim validationResult As CustomerValidationResult
    validationResult = ValidateCustomer(customer)
    
    If validationResult <> CustomerValidationResult.Valid Then
        Call ShowValidationErrorMessage(validationResult)
        RegisterNewCustomer = False
        Exit Function
    End If
    
    ' Verificar se documento já existe
    If DocumentExists(customer.Document) Then
        Call ShowErrorMessage("Documento já cadastrado: " & customer.Document)
        RegisterNewCustomer = False
        Exit Function
    End If
    
    ' Obter worksheet
    Dim ws As Worksheet
    Set ws = GetCustomerWorksheet
    
    If ws Is Nothing Then
        Call LogManager.LogError("RegisterNewCustomer", "Customer worksheet not found")
        RegisterNewCustomer = False
        Exit Function
    End If
    
    ' Gerar novo ID
    customer.ID = GetNextCustomerID(ws)
    
    ' Inserir cliente
    Call InsertCustomerIntoWorksheet(ws, customer)
    
    ' Atualizar cache
    mCustomerCache.Add customer.ID, customer
    
    ' Log de sucesso
    Call LogManager.LogInfo("Novo cliente cadastrado: " & customer.Name, "RegisterNewCustomer")
    
    RegisterNewCustomer = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("RegisterNewCustomer", Err)
    Call ShowErrorMessage("Erro ao cadastrar cliente: " & Err.Description)
    RegisterNewCustomer = False
End Function

Private Function ValidateCustomer(customer As CustomerData) As CustomerValidationResult
    ' Validar nome
    If mValidationRules.RequireName Then
        If Trim(customer.Name) = "" Then
            ValidateCustomer = CustomerValidationResult.InvalidName
            Exit Function
        End If
    End If
    
    ' Validar documento
    If mValidationRules.RequireDocument Then
        If Trim(customer.Document) = "" Then
            ValidateCustomer = CustomerValidationResult.InvalidDocument
            Exit Function
        End If
    End If
    
    If mValidationRules.ValidateDocument Then
        If Not IsValidDocument(customer.Document) Then
            ValidateCustomer = CustomerValidationResult.InvalidDocument
            Exit Function
        End If
    End If
    
    ' Validar email
    If mValidationRules.ValidateEmail And Trim(customer.Email) <> "" Then
        If Not IsValidEmail(customer.Email) Then
            ValidateCustomer = CustomerValidationResult.InvalidEmail
            Exit Function
        End If
    End If
    
    ' Validar telefone
    If mValidationRules.RequirePhone Then
        If Trim(customer.Phone) = "" Then
            ValidateCustomer = CustomerValidationResult.InvalidPhone
            Exit Function
        End If
    End If
    
    ValidateCustomer = CustomerValidationResult.Valid
End Function

Private Function IsValidDocument(document As String) As Boolean
    ' Remover caracteres especiais
    document = Replace(document, ".", "")
    document = Replace(document, "-", "")
    document = Replace(document, "/", "")
    
    ' Verificar se é numérico
    If Not IsNumeric(document) Then
        IsValidDocument = False
        Exit Function
    End If
    
    ' Verificar tamanho (CPF = 11, CNPJ = 14)
    If Len(document) <> 11 And Len(document) <> 14 Then
        IsValidDocument = False
        Exit Function
    End If
    
    ' Aqui você pode adicionar validação específica de CPF/CNPJ
    IsValidDocument = True
End Function

Private Function IsValidEmail(email As String) As Boolean
    ' Validação básica de email
    If InStr(email, "@") = 0 Or InStr(email, ".") = 0 Then
        IsValidEmail = False
        Exit Function
    End If
    
    If Left(email, 1) = "@" Or Right(email, 1) = "@" Then
        IsValidEmail = False
        Exit Function
    End If
    
    IsValidEmail = True
End Function

Private Function DocumentExists(document As String) As Boolean
    Dim ws As Worksheet
    Set ws = GetCustomerWorksheet
    
    If ws Is Nothing Then
        DocumentExists = False
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, CUSTOMER_DOCUMENT_COLUMN)
    
    Dim i As Long
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, CUSTOMER_DOCUMENT_COLUMN).Value)) = Trim(document) Then
            DocumentExists = True
            Exit Function
        End If
    Next i
    
    DocumentExists = False
End Function

Private Function GetNextCustomerID(ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, CUSTOMER_ID_COLUMN)
    
    If lastRow < 2 Then
        GetNextCustomerID = 1
    Else
        GetNextCustomerID = CLng(ws.Cells(lastRow, CUSTOMER_ID_COLUMN).Value) + 1
    End If
End Function

Private Sub InsertCustomerIntoWorksheet(ws As Worksheet, customer As CustomerData)
    Dim nextRow As Long
    nextRow = GetLastDataRow(ws, CUSTOMER_ID_COLUMN) + 1
    
    With ws
        .Cells(nextRow, CUSTOMER_ID_COLUMN).Value = customer.ID
        .Cells(nextRow, CUSTOMER_NAME_COLUMN).Value = customer.Name
        .Cells(nextRow, CUSTOMER_DOCUMENT_COLUMN).Value = customer.Document
        .Cells(nextRow, CUSTOMER_ADDRESS_COLUMN).Value = customer.Address
        .Cells(nextRow, CUSTOMER_CITY_COLUMN).Value = customer.City
        .Cells(nextRow, CUSTOMER_STATE_COLUMN).Value = customer.State
        .Cells(nextRow, CUSTOMER_ZIPCODE_COLUMN).Value = customer.ZipCode
        .Cells(nextRow, CUSTOMER_PHONE_COLUMN).Value = customer.Phone
        .Cells(nextRow, CUSTOMER_EMAIL_COLUMN).Value = customer.Email
    End With
End Sub

' === BUSCAR CLIENTES ===
Public Function SearchCustomers(criteria As CustomerSearchCriteria) As Collection
    On Error GoTo ErrorHandler
    
    Set SearchCustomers = New Collection
    
    Dim ws As Worksheet
    Set ws = GetCustomerWorksheet
    
    If ws Is Nothing Then
        Call LogManager.LogError("SearchCustomers", "Customer worksheet not found")
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, CUSTOMER_ID_COLUMN)
    
    Dim i As Long
    For i = 2 To lastRow
        If MatchesSearchCriteria(ws, i, criteria) Then
            Dim customer As CustomerData
            customer = GetCustomerFromRow(ws, i)
            SearchCustomers.Add customer
        End If
    Next i
    
    Exit Function

ErrorHandler:
    Call LogManager.LogError("SearchCustomers", Err)
    Call ShowErrorMessage("Erro ao buscar clientes: " & Err.Description)
End Function

Private Function MatchesSearchCriteria(ws As Worksheet, row As Long, criteria As CustomerSearchCriteria) As Boolean
    ' Verificar nome
    If criteria.Name <> "" Then
        If InStr(1, LCase(CStr(ws.Cells(row, CUSTOMER_NAME_COLUMN).Value)), LCase(criteria.Name)) = 0 Then
            MatchesSearchCriteria = False
            Exit Function
        End If
    End If
    
    ' Verificar documento
    If criteria.Document <> "" Then
        If InStr(1, CStr(ws.Cells(row, CUSTOMER_DOCUMENT_COLUMN).Value), criteria.Document) = 0 Then
            MatchesSearchCriteria = False
            Exit Function
        End If
    End If
    
    ' Verificar cidade
    If criteria.City <> "" Then
        If InStr(1, LCase(CStr(ws.Cells(row, CUSTOMER_CITY_COLUMN).Value)), LCase(criteria.City)) = 0 Then
            MatchesSearchCriteria = False
            Exit Function
        End If
    End If
    
    ' Verificar telefone
    If criteria.Phone <> "" Then
        If InStr(1, CStr(ws.Cells(row, CUSTOMER_PHONE_COLUMN).Value), criteria.Phone) = 0 Then
            MatchesSearchCriteria = False
            Exit Function
        End If
    End If
    
    ' Verificar email
    If criteria.Email <> "" Then
        If InStr(1, LCase(CStr(ws.Cells(row, CUSTOMER_EMAIL_COLUMN).Value)), LCase(criteria.Email)) = 0 Then
            MatchesSearchCriteria = False
            Exit Function
        End If
    End If
    
    MatchesSearchCriteria = True
End Function

' === ATUALIZAR CLIENTE ===
Public Function UpdateCustomer(customer As CustomerData) As Boolean
    On Error GoTo ErrorHandler
    
    ' Validar dados do cliente
    Dim validationResult As CustomerValidationResult
    validationResult = ValidateCustomer(customer)
    
    If validationResult <> CustomerValidationResult.Valid Then
        Call ShowValidationErrorMessage(validationResult)
        UpdateCustomer = False
        Exit Function
    End If
    
    ' Verificar se cliente existe
    If Not CustomerExists(customer.ID) Then
        Call ShowErrorMessage("Cliente não encontrado: " & customer.ID)
        UpdateCustomer = False
        Exit Function
    End If
    
    ' Obter worksheet
    Dim ws As Worksheet
    Set ws = GetCustomerWorksheet
    
    If ws Is Nothing Then
        Call LogManager.LogError("UpdateCustomer", "Customer worksheet not found")
        UpdateCustomer = False
        Exit Function
    End If
    
    ' Atualizar cliente
    Call UpdateCustomerInWorksheet(ws, customer)
    
    ' Atualizar cache
    mCustomerCache(customer.ID) = customer
    
    ' Log de sucesso
    Call LogManager.LogInfo("Cliente atualizado: " & customer.Name, "UpdateCustomer")
    
    UpdateCustomer = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("UpdateCustomer", Err)
    Call ShowErrorMessage("Erro ao atualizar cliente: " & Err.Description)
    UpdateCustomer = False
End Function

Private Function CustomerExists(customerID As Long) As Boolean
    Dim ws As Worksheet
    Set ws = GetCustomerWorksheet
    
    If ws Is Nothing Then
        CustomerExists = False
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, CUSTOMER_ID_COLUMN)
    
    Dim i As Long
    For i = 2 To lastRow
        If CLng(ws.Cells(i, CUSTOMER_ID_COLUMN).Value) = customerID Then
            CustomerExists = True
            Exit Function
        End If
    Next i
    
    CustomerExists = False
End Function

Private Sub UpdateCustomerInWorksheet(ws As Worksheet, customer As CustomerData)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, CUSTOMER_ID_COLUMN)
    
    Dim i As Long
    For i = 2 To lastRow
        If CLng(ws.Cells(i, CUSTOMER_ID_COLUMN).Value) = customer.ID Then
            With ws
                .Cells(i, CUSTOMER_NAME_COLUMN).Value = customer.Name
                .Cells(i, CUSTOMER_DOCUMENT_COLUMN).Value = customer.Document
                .Cells(i, CUSTOMER_ADDRESS_COLUMN).Value = customer.Address
                .Cells(i, CUSTOMER_CITY_COLUMN).Value = customer.City
                .Cells(i, CUSTOMER_STATE_COLUMN).Value = customer.State
                .Cells(i, CUSTOMER_ZIPCODE_COLUMN).Value = customer.ZipCode
                .Cells(i, CUSTOMER_PHONE_COLUMN).Value = customer.Phone
                .Cells(i, CUSTOMER_EMAIL_COLUMN).Value = customer.Email
            End With
            Exit For
        End If
    Next i
End Sub

' === EXCLUIR CLIENTE ===
Public Function DeleteCustomer(customerID As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificar se cliente existe
    If Not CustomerExists(customerID) Then
        Call ShowErrorMessage("Cliente não encontrado: " & customerID)
        DeleteCustomer = False
        Exit Function
    End If
    
    ' Verificar se cliente tem pedidos (implementar se necessário)
    If HasActiveOrders(customerID) Then
        Call ShowErrorMessage("Não é possível excluir cliente com pedidos ativos")
        DeleteCustomer = False
        Exit Function
    End If
    
    ' Obter worksheet
    Dim ws As Worksheet
    Set ws = GetCustomerWorksheet
    
    If ws Is Nothing Then
        Call LogManager.LogError("DeleteCustomer", "Customer worksheet not found")
        DeleteCustomer = False
        Exit Function
    End If
    
    ' Excluir cliente
    Call DeleteCustomerFromWorksheet(ws, customerID)
    
    ' Remover do cache
    If mCustomerCache.Exists(customerID) Then
        mCustomerCache.Remove customerID
    End If
    
    ' Log de sucesso
    Call LogManager.LogInfo("Cliente excluído: " & customerID, "DeleteCustomer")
    
    DeleteCustomer = True
    Exit Function

ErrorHandler:
    Call LogManager.LogError("DeleteCustomer", Err)
    Call ShowErrorMessage("Erro ao excluir cliente: " & Err.Description)
    DeleteCustomer = False
End Function

Private Function HasActiveOrders(customerID As Long) As Boolean
    ' Implementar verificação de pedidos ativos
    ' Por enquanto, retorna False
    HasActiveOrders = False
End Function

Private Sub DeleteCustomerFromWorksheet(ws As Worksheet, customerID As Long)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, CUSTOMER_ID_COLUMN)
    
    Dim i As Long
    For i = 2 To lastRow
        If CLng(ws.Cells(i, CUSTOMER_ID_COLUMN).Value) = customerID Then
            ws.Rows(i).Delete
            Exit For
        End If
    Next i
End Sub

'====================================================================
' UTILITÁRIOS
'====================================================================

Private Function GetCustomerWorksheet() As Worksheet
    On Error Resume Next
    Set GetCustomerWorksheet = ThisWorkbook.Worksheets(CUSTOMER_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function GetLastDataRow(ws As Worksheet, column As String) As Long
    GetLastDataRow = ws.Cells(ws.Rows.Count, column).End(xlUp).Row
End Function

Private Sub ClearCustomerCache()
    If Not mCustomerCache Is Nothing Then
        mCustomerCache.RemoveAll
    End If
End Sub

Private Sub ShowValidationErrorMessage(validationResult As CustomerValidationResult)
    Dim message As String
    
    Select Case validationResult
        Case CustomerValidationResult.InvalidName
            message = "Nome do cliente é obrigatório."
        Case CustomerValidationResult.InvalidDocument
            message = "Documento inválido ou obrigatório."
        Case CustomerValidationResult.InvalidEmail
            message = "Email inválido."
        Case CustomerValidationResult.InvalidPhone
            message = "Telefone é obrigatório."
        Case CustomerValidationResult.DuplicateDocument
            message = "Documento já cadastrado."
        Case Else
            message = "Dados do cliente inválidos."
    End Select
    
    Call ShowErrorMessage(message)
End Sub

Private Sub ShowErrorMessage(message As String)
    MsgBox message, vbExclamation, "Aviso - Gestão de Clientes"
End Sub

'====================================================================
' RELATÓRIOS E ESTATÍSTICAS
'====================================================================

Public Function GetCustomerStatistics() As Dictionary
    On Error GoTo ErrorHandler
    
    Set GetCustomerStatistics = New Dictionary
    
    Dim ws As Worksheet
    Set ws = GetCustomerWorksheet
    
    If ws Is Nothing Then
        Call LogManager.LogError("GetCustomerStatistics", "Customer worksheet not found")
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws, CUSTOMER_ID_COLUMN)
    
    If lastRow < 2 Then
        GetCustomerStatistics.Add "TotalCustomers", 0
        GetCustomerStatistics.Add "ActiveCustomers", 0
        GetCustomerStatistics.Add "NewCustomersThisMonth", 0
        Exit Function
    End If
    
    ' Calcular estatísticas
    Dim totalCustomers As Long
    Dim activeCustomers As Long
    Dim newCustomersThisMonth As Long
    
    totalCustomers = lastRow - 1 ' Excluir cabeçalho
    
    ' Contar clientes ativos (com telefone ou email)
    Dim i As Long
    For i = 2 To lastRow
        Dim phone As String, email As String
        phone = Trim(CStr(ws.Cells(i, CUSTOMER_PHONE_COLUMN).Value))
        email = Trim(CStr(ws.Cells(i, CUSTOMER_EMAIL_COLUMN).Value))
        
        If phone <> "" Or email <> "" Then
            activeCustomers = activeCustomers + 1
        End If
    Next i
    
    ' Contar novos clientes este mês (implementar se necessário)
    newCustomersThisMonth = 0
    
    ' Adicionar estatísticas ao dicionário
    GetCustomerStatistics.Add "TotalCustomers", totalCustomers
    GetCustomerStatistics.Add "ActiveCustomers", activeCustomers
    GetCustomerStatistics.Add "NewCustomersThisMonth", newCustomersThisMonth
    
    Exit Function

ErrorHandler:
    Call LogManager.LogError("GetCustomerStatistics", Err)
    Call ShowErrorMessage("Erro ao gerar estatísticas: " & Err.Description)
End Function

Public Sub ExportCustomerReport(filePath As String)
    On Error GoTo ErrorHandler
    
    ' Implementar exportação de relatório
    ' Por enquanto, apenas log
    Call LogManager.LogInfo("Exportação de relatório solicitada: " & filePath, "ExportCustomerReport")
    
    Exit Sub

ErrorHandler:
    Call LogManager.LogError("ExportCustomerReport", Err)
    Call ShowErrorMessage("Erro ao exportar relatório: " & Err.Description)
End Sub