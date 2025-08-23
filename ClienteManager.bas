' ====================================================================
' MÓDULO CLIENTE MANAGER - SISTEMA PDV MADEIREIRA MARIA LUZIA
' Responsável por todas as operações relacionadas aos clientes
' ====================================================================

Option Explicit

' === CARREGAR CLIENTES ===
Public Sub CarregarClientes(cmbCliente As MSForms.ComboBox)
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Clientes")
    
    cmbCliente.Clear
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If ultimaLinha < 2 Then
        MsgBox "⚠️ Nenhum cliente cadastrado!", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If Trim(CStr(ws.Cells(i, 2).Value)) <> "" Then
            cmbCliente.AddItem ws.Cells(i, 2).Value ' Nome/Razão Social
            cmbCliente.List(cmbCliente.ListCount - 1, 1) = ws.Cells(i, 1).Value ' ID
        End If
    Next i
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("CarregarClientes", Err)
    MsgBox "❌ Erro ao carregar clientes: " & Err.Description, vbCritical
End Sub

' === PREENCHER DADOS DO CLIENTE ===
Public Sub PreencherDadosCliente(cmbCliente As MSForms.ComboBox, frm As Object)
    On Error GoTo TratarErro
    
    If cmbCliente.ListIndex < 0 Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Clientes")
    
    Dim idCliente As Long
    idCliente = CLng(cmbCliente.List(cmbCliente.ListIndex, 1))
    
    ' Buscar dados do cliente
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If CLng(ws.Cells(i, 1).Value) = idCliente Then
            ' Armazenar dados do cliente no Tag do formulário
            frm.Tag = idCliente & "|" & _
                     ws.Cells(i, 2).Value & "|" & _
                     ws.Cells(i, 3).Value & "|" & _
                     ws.Cells(i, 4).Value & "|" & _
                     ws.Cells(i, 5).Value & "|" & _
                     ws.Cells(i, 6).Value & "|" & _
                     ws.Cells(i, 7).Value & "|" & _
                     ws.Cells(i, 8).Value
            
            ' Armazenar também na variável global
            Call TransferenciaDados.ArmazenarDadosCliente(frm.Tag)
            Exit For
        End If
    Next i
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("PreencherDadosCliente", Err)
    MsgBox "❌ Erro ao carregar dados do cliente: " & Err.Description, vbCritical
End Sub

' === CADASTRAR NOVO CLIENTE ===
Public Sub CadastrarNovoCliente()
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Clientes")
    
    ' Obter próximo ID
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim novoID As Long
    If ultimaLinha < 2 Then
        novoID = 1
    Else
        novoID = ws.Cells(ultimaLinha, 1).Value + 1
    End If
    
    ' Solicitar dados do cliente
    Dim nome As String, cpfCnpj As String, endereco As String
    Dim cidade As String, uf As String, cep As String, telefone As String
    
    nome = InputBox("Digite o nome/razão social:", "Cadastro de Cliente")
    If nome = "" Then Exit Sub
    
    cpfCnpj = InputBox("Digite o CPF/CNPJ:", "Cadastro de Cliente")
    If cpfCnpj = "" Then Exit Sub
    
    endereco = InputBox("Digite o endereço:", "Cadastro de Cliente")
    cidade = InputBox("Digite a cidade:", "Cadastro de Cliente")
    uf = InputBox("Digite o UF (ex: PE):", "Cadastro de Cliente")
    cep = InputBox("Digite o CEP:", "Cadastro de Cliente")
    telefone = InputBox("Digite o telefone:", "Cadastro de Cliente")
    
    ' Inserir dados
    Dim novaLinha As Long
    novaLinha = ultimaLinha + 1
    
    With ws
        .Cells(novaLinha, 1).Value = novoID
        .Cells(novaLinha, 2).Value = nome
        .Cells(novaLinha, 3).Value = cpfCnpj
        .Cells(novaLinha, 4).Value = endereco
        .Cells(novaLinha, 5).Value = cidade
        .Cells(novaLinha, 6).Value = uf
        .Cells(novaLinha, 7).Value = cep
        .Cells(novaLinha, 8).Value = telefone
    End With
    
    ' Atualizar estatísticas
    Call AtualizarEstatisticasClientes
    
    MsgBox "✅ Cliente cadastrado com sucesso!" & vbCrLf & _
           "ID: " & novoID & vbCrLf & _
           "Nome: " & nome, vbInformation, "Cadastro Concluído"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("CadastrarNovoCliente", Err)
    MsgBox "❌ Erro ao cadastrar cliente: " & Err.Description, vbCritical
End Sub

' === BUSCAR CLIENTE ===
Public Function BuscarCliente(criterio As String) As String
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Clientes")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        ' Buscar por nome ou CPF/CNPJ
        If InStr(1, UCase(ws.Cells(i, 2).Value), UCase(criterio)) > 0 Or _
           InStr(1, ws.Cells(i, 3).Value, criterio) > 0 Then
            
            BuscarCliente = ws.Cells(i, 1).Value & "|" & _
                           ws.Cells(i, 2).Value & "|" & _
                           ws.Cells(i, 3).Value & "|" & _
                           ws.Cells(i, 4).Value & "|" & _
                           ws.Cells(i, 5).Value & "|" & _
                           ws.Cells(i, 6).Value & "|" & _
                           ws.Cells(i, 7).Value & "|" & _
                           ws.Cells(i, 8).Value
            Exit Function
        End If
    Next i
    
    BuscarCliente = ""
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("BuscarCliente", Err)
    BuscarCliente = ""
End Function

' === VALIDAR CPF/CNPJ ===
Public Function ValidarCPFCNPJ(documento As String) As Boolean
    On Error GoTo TratarErro
    
    ' Remove caracteres especiais
    documento = Replace(Replace(Replace(documento, ".", ""), "-", ""), "/", "")
    
    If Len(documento) = 11 Then
        ' Validação básica de CPF
        ValidarCPFCNPJ = ValidarCPF(documento)
    ElseIf Len(documento) = 14 Then
        ' Validação básica de CNPJ
        ValidarCPFCNPJ = ValidarCNPJ(documento)
    Else
        ValidarCPFCNPJ = False
    End If
    
    Exit Function
TratarErro:
    ValidarCPFCNPJ = False
End Function

' === VALIDAR CPF ===
Private Function ValidarCPF(cpf As String) As Boolean
    On Error GoTo TratarErro
    
    ' Validação básica de CPF (sem dígitos verificadores completos)
    If Len(cpf) <> 11 Then
        ValidarCPF = False
        Exit Function
    End If
    
    ' Verificar se todos os dígitos são iguais
    Dim i As Integer
    Dim digitoAnterior As String
    digitoAnterior = Mid(cpf, 1, 1)
    
    For i = 2 To 11
        If Mid(cpf, i, 1) <> digitoAnterior Then
            ValidarCPF = True
            Exit Function
        End If
    Next i
    
    ValidarCPF = False ' Todos os dígitos são iguais
    
    Exit Function
TratarErro:
    ValidarCPF = False
End Function

' === VALIDAR CNPJ ===
Private Function ValidarCNPJ(cnpj As String) As Boolean
    On Error GoTo TratarErro
    
    ' Validação básica de CNPJ
    If Len(cnpj) <> 14 Then
        ValidarCNPJ = False
        Exit Function
    End If
    
    ' Verificar se todos os dígitos são iguais
    Dim i As Integer
    Dim digitoAnterior As String
    digitoAnterior = Mid(cnpj, 1, 1)
    
    For i = 2 To 14
        If Mid(cnpj, i, 1) <> digitoAnterior Then
            ValidarCNPJ = True
            Exit Function
        End If
    Next i
    
    ValidarCNPJ = False ' Todos os dígitos são iguais
    
    Exit Function
TratarErro:
    ValidarCNPJ = False
End Function

' === ATUALIZAR ESTATÍSTICAS ===
Public Sub AtualizarEstatisticasClientes()
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Clientes")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim totalClientes As Long
    totalClientes = IIf(ultimaLinha > 1, ultimaLinha - 1, 0)
    
    ' Atualizar planilha de controle
    Dim wsControle As Worksheet
    Set wsControle = ThisWorkbook.Worksheets("Controle")
    wsControle.Cells(2, 3).Value = totalClientes
    
    ' Atualizar Dashboard se existir
    On Error Resume Next
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    If Not wsDashboard Is Nothing Then
        wsDashboard.Range("H8:J10").Value = "👥 TOTAL DE CLIENTES" & vbCrLf & vbCrLf & totalClientes
    End If
    On Error GoTo TratarErro
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AtualizarEstatisticasClientes", Err)
End Sub

' === LISTAR TODOS OS CLIENTES ===
Public Function ListarClientes() As String
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Clientes")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim resultado As String
    resultado = "ID|Nome|CPF/CNPJ|Cidade" & vbCrLf
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If Trim(CStr(ws.Cells(i, 2).Value)) <> "" Then
            resultado = resultado & ws.Cells(i, 1).Value & "|" & _
                       ws.Cells(i, 2).Value & "|" & _
                       ws.Cells(i, 3).Value & "|" & _
                       ws.Cells(i, 5).Value & vbCrLf
        End If
    Next i
    
    ListarClientes = resultado
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ListarClientes", Err)
    ListarClientes = "Erro ao listar clientes"
End Function

' === EDITAR CLIENTE ===
Public Sub EditarCliente(idCliente As Long)
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Clientes")
    
    ' Buscar cliente
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim linhaCliente As Long
    linhaCliente = 0
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If CLng(ws.Cells(i, 1).Value) = idCliente Then
            linhaCliente = i
            Exit For
        End If
    Next i
    
    If linhaCliente = 0 Then
        MsgBox "❌ Cliente não encontrado!", vbExclamation
        Exit Sub
    End If
    
    ' Solicitar novos dados
    Dim nome As String, cpfCnpj As String, endereco As String
    Dim cidade As String, uf As String, cep As String, telefone As String
    
    nome = InputBox("Nome/Razão Social:", "Editar Cliente", ws.Cells(linhaCliente, 2).Value)
    If nome = "" Then Exit Sub
    
    cpfCnpj = InputBox("CPF/CNPJ:", "Editar Cliente", ws.Cells(linhaCliente, 3).Value)
    endereco = InputBox("Endereço:", "Editar Cliente", ws.Cells(linhaCliente, 4).Value)
    cidade = InputBox("Cidade:", "Editar Cliente", ws.Cells(linhaCliente, 5).Value)
    uf = InputBox("UF:", "Editar Cliente", ws.Cells(linhaCliente, 6).Value)
    cep = InputBox("CEP:", "Editar Cliente", ws.Cells(linhaCliente, 7).Value)
    telefone = InputBox("Telefone:", "Editar Cliente", ws.Cells(linhaCliente, 8).Value)
    
    ' Atualizar dados
    With ws
        .Cells(linhaCliente, 2).Value = nome
        .Cells(linhaCliente, 3).Value = cpfCnpj
        .Cells(linhaCliente, 4).Value = endereco
        .Cells(linhaCliente, 5).Value = cidade
        .Cells(linhaCliente, 6).Value = uf
        .Cells(linhaCliente, 7).Value = cep
        .Cells(linhaCliente, 8).Value = telefone
    End With
    
    MsgBox "✅ Cliente atualizado com sucesso!", vbInformation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("EditarCliente", Err)
    MsgBox "❌ Erro ao editar cliente: " & Err.Description, vbCritical
End Sub

' === EXCLUIR CLIENTE ===
Public Sub ExcluirCliente(idCliente As Long)
    On Error GoTo TratarErro
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("⚠️ Tem certeza que deseja excluir este cliente?" & vbCrLf & _
                     "Esta ação não pode ser desfeita!", vbYesNo + vbQuestion, "Confirmar Exclusão")
    
    If resposta = vbNo Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Clientes")
    
    ' Buscar e excluir cliente
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If CLng(ws.Cells(i, 1).Value) = idCliente Then
            ws.Rows(i).Delete
            MsgBox "✅ Cliente excluído com sucesso!", vbInformation
            Call AtualizarEstatisticasClientes
            Exit Sub
        End If
    Next i
    
    MsgBox "❌ Cliente não encontrado!", vbExclamation
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ExcluirCliente", Err)
    MsgBox "❌ Erro ao excluir cliente: " & Err.Description, vbCritical
End Sub

' === OBTER DADOS DO CLIENTE POR ID ===
Public Function ObterClientePorID(idCliente As Long) As String
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Clientes")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If CLng(ws.Cells(i, 1).Value) = idCliente Then
            ObterClientePorID = ws.Cells(i, 1).Value & "|" & _
                               ws.Cells(i, 2).Value & "|" & _
                               ws.Cells(i, 3).Value & "|" & _
                               ws.Cells(i, 4).Value & "|" & _
                               ws.Cells(i, 5).Value & "|" & _
                               ws.Cells(i, 6).Value & "|" & _
                               ws.Cells(i, 7).Value & "|" & _
                               ws.Cells(i, 8).Value
            Exit Function
        End If
    Next i
    
    ObterClientePorID = ""
    
    Exit Function
TratarErro:
    Call ErrorHandler.RegistrarErro("ObterClientePorID", Err)
    ObterClientePorID = ""
End Function