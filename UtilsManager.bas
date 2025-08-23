' ====================================================================
' MÓDULO UTILS MANAGER - SISTEMA PDV MADEIREIRA MARIA LUZIA
' Responsável por funções utilitárias e validações
' ====================================================================

Option Explicit

' === VALIDAR CPF ===
Public Function ValidarCPF(cpf As String) As Boolean
    On Error GoTo TratarErro
    
    ' Remover formatação
    Dim cpfNumerico As String
    cpfNumerico = ""
    
    Dim i As Integer
    For i = 1 To Len(cpf)
        If IsNumeric(Mid(cpf, i, 1)) Then
            cpfNumerico = cpfNumerico & Mid(cpf, i, 1)
        End If
    Next i
    
    ' Verificar se tem 11 dígitos
    If Len(cpfNumerico) <> 11 Then
        ValidarCPF = False
        Exit Function
    End If
    
    ' Verificar se não são todos iguais
    If cpfNumerico = String(11, Left(cpfNumerico, 1)) Then
        ValidarCPF = False
        Exit Function
    End If
    
    ' Calcular primeiro dígito verificador
    Dim soma As Long
    soma = 0
    
    For i = 1 To 9
        soma = soma + (CInt(Mid(cpfNumerico, i, 1)) * (11 - i))
    Next i
    
    Dim resto As Integer
    resto = soma Mod 11
    
    Dim digito1 As Integer
    If resto < 2 Then
        digito1 = 0
    Else
        digito1 = 11 - resto
    End If
    
    ' Verificar primeiro dígito
    If digito1 <> CInt(Mid(cpfNumerico, 10, 1)) Then
        ValidarCPF = False
        Exit Function
    End If
    
    ' Calcular segundo dígito verificador
    soma = 0
    For i = 1 To 10
        soma = soma + (CInt(Mid(cpfNumerico, i, 1)) * (12 - i))
    Next i
    
    resto = soma Mod 11
    
    Dim digito2 As Integer
    If resto < 2 Then
        digito2 = 0
    Else
        digito2 = 11 - resto
    End If
    
    ' Verificar segundo dígito
    ValidarCPF = (digito2 = CInt(Mid(cpfNumerico, 11, 1)))
    
    Exit Function
TratarErro:
    ValidarCPF = False
End Function

' === VALIDAR CNPJ ===
Public Function ValidarCNPJ(cnpj As String) As Boolean
    On Error GoTo TratarErro
    
    ' Remover formatação
    Dim cnpjNumerico As String
    cnpjNumerico = ""
    
    Dim i As Integer
    For i = 1 To Len(cnpj)
        If IsNumeric(Mid(cnpj, i, 1)) Then
            cnpjNumerico = cnpjNumerico & Mid(cnpj, i, 1)
        End If
    Next i
    
    ' Verificar se tem 14 dígitos
    If Len(cnpjNumerico) <> 14 Then
        ValidarCNPJ = False
        Exit Function
    End If
    
    ' Verificar se não são todos iguais
    If cnpjNumerico = String(14, Left(cnpjNumerico, 1)) Then
        ValidarCNPJ = False
        Exit Function
    End If
    
    ' Calcular primeiro dígito verificador
    Dim multiplicadores1 As Variant
    multiplicadores1 = Array(5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    
    Dim soma As Long
    soma = 0
    
    For i = 1 To 12
        soma = soma + (CInt(Mid(cnpjNumerico, i, 1)) * multiplicadores1(i - 1))
    Next i
    
    Dim resto As Integer
    resto = soma Mod 11
    
    Dim digito1 As Integer
    If resto < 2 Then
        digito1 = 0
    Else
        digito1 = 11 - resto
    End If
    
    ' Verificar primeiro dígito
    If digito1 <> CInt(Mid(cnpjNumerico, 13, 1)) Then
        ValidarCNPJ = False
        Exit Function
    End If
    
    ' Calcular segundo dígito verificador
    Dim multiplicadores2 As Variant
    multiplicadores2 = Array(6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    
    soma = 0
    For i = 1 To 13
        soma = soma + (CInt(Mid(cnpjNumerico, i, 1)) * multiplicadores2(i - 1))
    Next i
    
    resto = soma Mod 11
    
    Dim digito2 As Integer
    If resto < 2 Then
        digito2 = 0
    Else
        digito2 = 11 - resto
    End If
    
    ' Verificar segundo dígito
    ValidarCNPJ = (digito2 = CInt(Mid(cnpjNumerico, 14, 1)))
    
    Exit Function
TratarErro:
    ValidarCNPJ = False
End Function

' === VALIDAR EMAIL ===
Public Function ValidarEmail(email As String) As Boolean
    On Error GoTo TratarErro
    
    ValidarEmail = False
    
    ' Verificações básicas
    If Len(Trim(email)) = 0 Then Exit Function
    If InStr(email, "@") = 0 Then Exit Function
    If InStr(email, ".") = 0 Then Exit Function
    
    ' Verificar formato básico
    Dim partes() As String
    partes = Split(email, "@")
    
    If UBound(partes) <> 1 Then Exit Function
    If Len(partes(0)) = 0 Or Len(partes(1)) = 0 Then Exit Function
    If InStr(partes(1), ".") = 0 Then Exit Function
    
    ' Verificar caracteres inválidos
    Dim caracteresInvalidos As String
    caracteresInvalidos = " !#$%&'*+/=?^`{|}~"
    
    Dim i As Integer
    For i = 1 To Len(caracteresInvalidos)
        If InStr(email, Mid(caracteresInvalidos, i, 1)) > 0 Then Exit Function
    Next i
    
    ValidarEmail = True
    
    Exit Function
TratarErro:
    ValidarEmail = False
End Function

' === VALIDAR TELEFONE ===
Public Function ValidarTelefone(telefone As String) As Boolean
    On Error GoTo TratarErro
    
    ' Extrair apenas números
    Dim telefoneNumerico As String
    telefoneNumerico = ""
    
    Dim i As Integer
    For i = 1 To Len(telefone)
        If IsNumeric(Mid(telefone, i, 1)) Then
            telefoneNumerico = telefoneNumerico & Mid(telefone, i, 1)
        End If
    Next i
    
    ' Verificar quantidade de dígitos (10 ou 11 dígitos)
    ValidarTelefone = (Len(telefoneNumerico) = 10 Or Len(telefoneNumerico) = 11)
    
    Exit Function
TratarErro:
    ValidarTelefone = False
End Function

' === FORMATAR CPF ===
Public Function FormatarCPF(cpf As String) As String
    On Error GoTo TratarErro
    
    ' Extrair apenas números
    Dim cpfNumerico As String
    cpfNumerico = ""
    
    Dim i As Integer
    For i = 1 To Len(cpf)
        If IsNumeric(Mid(cpf, i, 1)) Then
            cpfNumerico = cpfNumerico & Mid(cpf, i, 1)
        End If
    Next i
    
    ' Formatar se tiver 11 dígitos
    If Len(cpfNumerico) = 11 Then
        FormatarCPF = Left(cpfNumerico, 3) & "." & _
                      Mid(cpfNumerico, 4, 3) & "." & _
                      Mid(cpfNumerico, 7, 3) & "-" & _
                      Right(cpfNumerico, 2)
    Else
        FormatarCPF = cpf
    End If
    
    Exit Function
TratarErro:
    FormatarCPF = cpf
End Function

' === FORMATAR CNPJ ===
Public Function FormatarCNPJ(cnpj As String) As String
    On Error GoTo TratarErro
    
    ' Extrair apenas números
    Dim cnpjNumerico As String
    cnpjNumerico = ""
    
    Dim i As Integer
    For i = 1 To Len(cnpj)
        If IsNumeric(Mid(cnpj, i, 1)) Then
            cnpjNumerico = cnpjNumerico & Mid(cnpj, i, 1)
        End If
    Next i
    
    ' Formatar se tiver 14 dígitos
    If Len(cnpjNumerico) = 14 Then
        FormatarCNPJ = Left(cnpjNumerico, 2) & "." & _
                       Mid(cnpjNumerico, 3, 3) & "." & _
                       Mid(cnpjNumerico, 6, 3) & "/" & _
                       Mid(cnpjNumerico, 9, 4) & "-" & _
                       Right(cnpjNumerico, 2)
    Else
        FormatarCNPJ = cnpj
    End If
    
    Exit Function
TratarErro:
    FormatarCNPJ = cnpj
End Function

' === FORMATAR TELEFONE ===
Public Function FormatarTelefone(telefone As String) As String
    On Error GoTo TratarErro
    
    ' Extrair apenas números
    Dim telefoneNumerico As String
    telefoneNumerico = ""
    
    Dim i As Integer
    For i = 1 To Len(telefone)
        If IsNumeric(Mid(telefone, i, 1)) Then
            telefoneNumerico = telefoneNumerico & Mid(telefone, i, 1)
        End If
    Next i
    
    ' Formatar baseado na quantidade de dígitos
    Select Case Len(telefoneNumerico)
        Case 10 ' Telefone fixo
            FormatarTelefone = "(" & Left(telefoneNumerico, 2) & ") " & _
                              Mid(telefoneNumerico, 3, 4) & "-" & _
                              Right(telefoneNumerico, 4)
        Case 11 ' Celular
            FormatarTelefone = "(" & Left(telefoneNumerico, 2) & ") " & _
                              Mid(telefoneNumerico, 3, 5) & "-" & _
                              Right(telefoneNumerico, 4)
        Case Else
            FormatarTelefone = telefone
    End Select
    
    Exit Function
TratarErro:
    FormatarTelefone = telefone
End Function

' === FORMATAR CEP ===
Public Function FormatarCEP(cep As String) As String
    On Error GoTo TratarErro
    
    ' Extrair apenas números
    Dim cepNumerico As String
    cepNumerico = ""
    
    Dim i As Integer
    For i = 1 To Len(cep)
        If IsNumeric(Mid(cep, i, 1)) Then
            cepNumerico = cepNumerico & Mid(cep, i, 1)
        End If
    Next i
    
    ' Formatar se tiver 8 dígitos
    If Len(cepNumerico) = 8 Then
        FormatarCEP = Left(cepNumerico, 5) & "-" & Right(cepNumerico, 3)
    Else
        FormatarCEP = cep
    End If
    
    Exit Function
TratarErro:
    FormatarCEP = cep
End Function

' === VALIDAR VALOR NUMÉRICO ===
Public Function ValidarValorNumerico(valor As String) As Boolean
    On Error GoTo TratarErro
    
    If Trim(valor) = "" Then
        ValidarValorNumerico = False
        Exit Function
    End If
    
    ' Remover formatação monetária
    valor = Replace(Replace(valor, "R$", ""), " ", "")
    valor = Replace(valor, ".", "")
    valor = Replace(valor, ",", ".")
    
    ValidarValorNumerico = IsNumeric(valor) And CDbl(valor) >= 0
    
    Exit Function
TratarErro:
    ValidarValorNumerico = False
End Function

' === VALIDAR TEXTO ===
Public Function ValidarTexto(texto As String, tamanhoMinimo As Integer, tamanhoMaximo As Integer) As Boolean
    On Error GoTo TratarErro
    
    Dim tamanho As Integer
    tamanho = Len(Trim(texto))
    
    ValidarTexto = (tamanho >= tamanhoMinimo And tamanho <= tamanhoMaximo)
    
    Exit Function
TratarErro:
    ValidarTexto = False
End Function

' === LIMPAR TEXTO ===
Public Function LimparTexto(texto As String) As String
    On Error GoTo TratarErro
    
    ' Remover caracteres especiais e múltiplos espaços
    Dim resultado As String
    resultado = Trim(texto)
    
    ' Substituir múltiplos espaços por um único espaço
    Do While InStr(resultado, "  ") > 0
        resultado = Replace(resultado, "  ", " ")
    Loop
    
    LimparTexto = resultado
    
    Exit Function
TratarErro:
    LimparTexto = texto
End Function

' === CAPITALIZAR TEXTO ===
Public Function CapitalizarTexto(texto As String) As String
    On Error GoTo TratarErro
    
    If Trim(texto) = "" Then
        CapitalizarTexto = ""
        Exit Function
    End If
    
    Dim palavras() As String
    palavras = Split(LCase(Trim(texto)), " ")
    
    Dim i As Integer
    For i = 0 To UBound(palavras)
        If Len(palavras(i)) > 0 Then
            ' Não capitalizar preposições pequenas (exceto se for a primeira palavra)
            If i = 0 Or Not (palavras(i) = "de" Or palavras(i) = "da" Or palavras(i) = "do" Or _
                            palavras(i) = "das" Or palavras(i) = "dos" Or palavras(i) = "e") Then
                palavras(i) = UCase(Left(palavras(i), 1)) & Mid(palavras(i), 2)
            End If
        End If
    Next i
    
    CapitalizarTexto = Join(palavras, " ")
    
    Exit Function
TratarErro:
    CapitalizarTexto = texto
End Function

' === GERAR ID ÚNICO ===
Public Function GerarIDUnico() As String
    On Error GoTo TratarErro
    
    ' Gerar ID baseado em timestamp + números aleatórios
    Randomize
    
    Dim timestamp As String
    timestamp = Format(Now(), "yyyymmddhhmmss")
    
    Dim numeroAleatorio As String
    numeroAleatorio = Format(Int(Rnd() * 1000), "000")
    
    GerarIDUnico = timestamp & numeroAleatorio
    
    Exit Function
TratarErro:
    GerarIDUnico = Format(Now(), "yyyymmddhhmmss")
End Function

' === CRIAR BACKUP ===
Public Sub CriarBackup()
    On Error GoTo TratarErro
    
    Dim nomeBackup As String
    Dim caminhoBackup As String
    
    nomeBackup = "Backup_PDV_" & Format(Now(), "yyyymmdd_hhmmss") & ".xlsm"
    caminhoBackup = ThisWorkbook.Path & "\" & nomeBackup
    
    ' Salvar cópia
    ThisWorkbook.SaveCopyAs caminhoBackup
    
    MsgBox "✅ Backup criado com sucesso!" & vbCrLf & _
           "Arquivo: " & nomeBackup & vbCrLf & _
           "Local: " & ThisWorkbook.Path, vbInformation, "Backup Concluído"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("CriarBackup", Err)
    MsgBox "❌ Erro ao criar backup: " & Err.Description, vbCritical
End Sub

' === EXPORTAR LOG DE ERROS ===
Public Sub ExportarLogErros()
    On Error GoTo TratarErro
    
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets("Log_Erros")
    
    ' Verificar se há dados
    If wsLog.Cells(2, 1).Value = "" Then
        MsgBox "⚠️ Não há erros registrados para exportar.", vbInformation
        Exit Sub
    End If
    
    ' Criar arquivo de texto
    Dim nomeArquivo As String
    Dim caminhoArquivo As String
    
    nomeArquivo = "Log_Erros_" & Format(Now(), "yyyymmdd_hhmmss") & ".txt"
    caminhoArquivo = ThisWorkbook.Path & "\" & nomeArquivo
    
    Dim arquivo As Integer
    arquivo = FreeFile
    
    Open caminhoArquivo For Output As arquivo
    
    ' Cabeçalho
    Print #arquivo, "LOG DE ERROS - SISTEMA PDV MADEIREIRA MARIA LUZIA"
    Print #arquivo, "Exportado em: " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
    Print #arquivo, String(60, "=")
    Print #arquivo, ""
    
    ' Dados
    Dim ultimaLinha As Long
    ultimaLinha = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        Print #arquivo, "Data/Hora: " & wsLog.Cells(i, 1).Value
        Print #arquivo, "Procedimento: " & wsLog.Cells(i, 2).Value
        Print #arquivo, "Número: " & wsLog.Cells(i, 3).Value
        Print #arquivo, "Descrição: " & wsLog.Cells(i, 4).Value
        Print #arquivo, "Usuário: " & wsLog.Cells(i, 5).Value
        Print #arquivo, String(40, "-")
    Next i
    
    Close arquivo
    
    MsgBox "✅ Log de erros exportado!" & vbCrLf & _
           "Arquivo: " & nomeArquivo, vbInformation, "Exportação Concluída"
    
    Exit Sub
TratarErro:
    If arquivo > 0 Then Close arquivo
    Call ErrorHandler.RegistrarErro("ExportarLogErros", Err)
    MsgBox "❌ Erro ao exportar log: " & Err.Description, vbCritical
End Sub

' === VERIFICAR INTEGRIDADE DO SISTEMA ===
Public Function VerificarIntegridadeSistema() As String
    On Error GoTo TratarErro
    
    Dim relatorio As String
    relatorio = "VERIFICAÇÃO DE INTEGRIDADE DO SISTEMA" & vbCrLf & vbCrLf
    
    Dim problemas As Long
    problemas = 0
    
    ' Verificar planilhas obrigatórias
    Dim planilhasObrigatorias As Variant
    planilhasObrigatorias = Array("Dashboard", "Clientes", "Produtos", "Template_Pedido", "Controle", "Log_Erros")
    
    Dim i As Integer
    For i = 0 To UBound(planilhasObrigatorias)
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(planilhasObrigatorias(i))
        On Error GoTo TratarErro
        
        If ws Is Nothing Then
            relatorio = relatorio & "❌ Planilha '" & planilhasObrigatorias(i) & "' não encontrada" & vbCrLf
            problemas = problemas + 1
        Else
            relatorio = relatorio & "✅ Planilha '" & planilhasObrigatorias(i) & "' OK" & vbCrLf
        End If
        Set ws = Nothing
    Next i
    
    ' Verificar módulos VBA
    Dim modulosObrigatorios As Variant
    modulosObrigatorios = Array("ClienteManager", "ProdutoManager", "DescontoManager", _
                               "CalculadoraManager", "PedidoManager", "UtilsManager", _
                               "ErrorHandler", "ImpressaoManager", "TransferenciaDados")
    
    relatorio = relatorio & vbCrLf
    
    For i = 0 To UBound(modulosObrigatorios)
        On Error Resume Next
        Dim modulo As Object
        Set modulo = ThisWorkbook.VBProject.VBComponents(modulosObrigatorios(i))
        On Error GoTo TratarErro
        
        If modulo Is Nothing Then
            relatorio = relatorio & "❌ Módulo '" & modulosObrigatorios(i) & "' não encontrado" & vbCrLf
            problemas = problemas + 1
        Else
            relatorio = relatorio & "✅ Módulo '" & modulosObrigatorios(i) & "' OK" & vbCrLf
        End If
        Set modulo = Nothing
    Next i
    
    ' Verificar estrutura das planilhas
    relatorio = relatorio & vbCrLf & "Verificando estruturas:" & vbCrLf
    
    ' Verificar planilha Produtos
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Produtos")
    If Not ws Is Nothing Then
        If ws.Cells(1, 1).Value = "Referencia" And ws.Cells(1, 2).Value = "Descricao" Then
            relatorio = relatorio & "✅ Estrutura da planilha Produtos OK" & vbCrLf
        Else
            relatorio = relatorio & "❌ Estrutura da planilha Produtos incorreta" & vbCrLf
            problemas = problemas + 1
        End If
    End If
    
    ' Verificar planilha Clientes
    Set ws = ThisWorkbook.Worksheets("Clientes")
    If Not ws Is Nothing Then
        If ws.Cells(1, 1).Value = "ID_Cliente" And ws.Cells(1, 2).Value = "Nome_RazaoSocial" Then
            relatorio = relatorio & "✅ Estrutura da planilha Clientes OK" & vbCrLf
        Else
            relatorio = relatorio & "❌ Estrutura da planilha Clientes incorreta" & vbCrLf
            problemas = problemas + 1
        End If
    End If
    On Error GoTo TratarErro
    
    ' Resumo final
    relatorio = relatorio & vbCrLf & "RESUMO:" & vbCrLf
    If problemas = 0 Then
        relatorio = relatorio & "✅ Sistema íntegro - Nenhum problema encontrado"
    Else
        relatorio = relatorio & "⚠️ " & problemas & " problema(s) encontrado(s)"
    End If
    
    VerificarIntegridadeSistema = relatorio
    
    Exit Function
TratarErro:
    VerificarIntegridadeSistema = "Erro ao verificar integridade: " & Err.Description
End Function

' === LIMPAR DADOS TEMPORÁRIOS ===
Public Sub LimparDadosTemporarios()
    On Error GoTo TratarErro
    
    Dim planilhasRemovidas As Long
    planilhasRemovidas = 0
    
    Application.DisplayAlerts = False
    
    ' Remover planilhas temporárias
    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(i)
        
        If InStr(ws.Name, "Temp_") > 0 Or _
           InStr(ws.Name, "Backup_") > 0 Or _
           Left(ws.Name, 8) = "Relatorio" Then
            ws.Delete
            planilhasRemovidas = planilhasRemovidas + 1
        End If
    Next i
    
    Application.DisplayAlerts = True
    
    ' Limpar dados globais
    Call TransferenciaDados.LimparDadosGlobais
    
    MsgBox "✅ Limpeza concluída!" & vbCrLf & _
           "Planilhas temporárias removidas: " & planilhasRemovidas & vbCrLf & _
           "Dados globais limpos", vbInformation, "Limpeza de Dados"
    
    Exit Sub
TratarErro:
    Application.DisplayAlerts = True
    Call ErrorHandler.RegistrarErro("LimparDadosTemporarios", Err)
    MsgBox "❌ Erro na limpeza: " & Err.Description, vbCritical
End Sub

' === OBTER INFORMAÇÕES DO SISTEMA ===
Public Function ObterInformacoesSistema() As String
    On Error GoTo TratarErro
    
    Dim info As String
    info = "INFORMAÇÕES DO SISTEMA PDV" & vbCrLf & vbCrLf
    
    ' Informações básicas
    info = info & "Nome: Sistema PDV Madeireira Maria Luzia" & vbCrLf
    info = info & "Versão: 1.0" & vbCrLf
    info = info & "Data de instalação: " & Format(Now(), "dd/mm/yyyy") & vbCrLf
    info = info & "Usuário: " & Environ("USERNAME") & vbCrLf
    info = info & "Computador: " & Environ("COMPUTERNAME") & vbCrLf & vbCrLf
    
    ' Estatísticas
    Dim totalPlanilhas As Long
    Dim totalPedidos As Long
    Dim totalClientes As Long
    Dim totalProdutos As Long
    
    totalPlanilhas = ThisWorkbook.Worksheets.Count
    
    ' Contar pedidos
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Left(ThisWorkbook.Worksheets(i).Name, 7) = "Pedido_" Then
            totalPedidos = totalPedidos + 1
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
    
    ' Contar produtos
    Dim wsProdutos As Worksheet
    Set wsProdutos = ThisWorkbook.Worksheets("Produtos")
    If Not wsProdutos Is Nothing Then
        totalProdutos = wsProdutos.Cells(wsProdutos.Rows.Count, "A").End(xlUp).Row - 1
        If totalProdutos < 0 Then totalProdutos = 0
    End If
    On Error GoTo TratarErro
    
    info = info & "ESTATÍSTICAS:" & vbCrLf
    info = info & "Total de planilhas: " & totalPlanilhas & vbCrLf
    info = info & "Total de pedidos: " & totalPedidos & vbCrLf
    info = info & "Total de clientes: " & totalClientes & vbCrLf
    info = info & "Total de produtos: " & totalProdutos & vbCrLf & vbCrLf
    
    ' Informações técnicas
    info = info & "INFORMAÇÕES TÉCNICAS:" & vbCrLf
    info = info & "Excel: " & Application.Version & vbCrLf
    info = info & "Sistema Operacional: " & Application.OperatingSystem & vbCrLf
    info = info & "Arquivo: " & ThisWorkbook.Name & vbCrLf
    info = info & "Caminho: " & ThisWorkbook.Path
    
    ObterInformacoesSistema = info
    
    Exit Function
TratarErro:
    ObterInformacoesSistema = "Erro ao obter informações: " & Err.Description
End Function

' === CONFIGURAR SISTEMA ===
Public Sub ConfigurarSistema()
    On Error GoTo TratarErro
    
    ' Configurações gerais do Excel para o sistema
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    ' Configurar planilhas ocultas
    On Error Resume Next
    ThisWorkbook.Worksheets("Controle").Visible = xlSheetHidden
    ThisWorkbook.Worksheets("Log_Erros").Visible = xlSheetHidden
    On Error GoTo TratarErro
    
    MsgBox "✅ Sistema configurado com sucesso!" & vbCrLf & _
           "Configurações aplicadas:" & vbCrLf & _
           "• Planilhas de controle ocultas" & vbCrLf & _
           "• Cálculo automático ativado" & vbCrLf & _
           "• Eventos habilitados", vbInformation, "Configuração Concluída"
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("ConfigurarSistema", Err)
    MsgBox "❌ Erro na configuração: " & Err.Description, vbCritical
End Sub

' === RESETAR SISTEMA ===
Public Sub ResetarSistema()
    On Error GoTo TratarErro
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("⚠️ ATENÇÃO: Esta operação irá:" & vbCrLf & vbCrLf & _
                     "• Remover todos os pedidos" & vbCrLf & _
                     "• Zerar estatísticas" & vbCrLf & _
                     "• Limpar log de erros" & vbCrLf & _
                     "• Manter apenas clientes e produtos" & vbCrLf & vbCrLf & _
                     "Tem certeza que deseja continuar?", _
                     vbYesNo + vbCritical, "Confirmar Reset")
    
    If resposta = vbNo Then Exit Sub
    
    ' Criar backup antes do reset
    Call CriarBackup
    
    Application.DisplayAlerts = False
    
    ' Remover todas as planilhas de pedidos
    Dim pedidosRemovidos As Long
    pedidosRemovidos = 0
    
    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If Left(ThisWorkbook.Worksheets(i).Name, 7) = "Pedido_" Then
            ThisWorkbook.Worksheets(i).Delete
            pedidosRemovidos = pedidosRemovidos + 1
        End If
    Next i
    
    ' Zerar estatísticas
    Dim wsControle As Worksheet
    Set wsControle = ThisWorkbook.Worksheets("Controle")
    wsControle.Cells(2, 1).Value = 0 ' Último pedido
    wsControle.Cells(2, 2).Value = 0 ' Total vendas
    
    ' Limpar log de erros
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets("Log_Erros")
    wsLog.Range("A2:E1000").ClearContents
    
    Application.DisplayAlerts = True
    
    MsgBox "✅ Sistema resetado com sucesso!" & vbCrLf & vbCrLf & _
           "Pedidos removidos: " & pedidosRemovidos & vbCrLf & _
           "Estatísticas zeradas" & vbCrLf & _
           "Log de erros limpo" & vbCrLf & _
           "Backup criado antes do reset", vbInformation, "Reset Concluído"
    
    Exit Sub
TratarErro:
    Application.DisplayAlerts = True
    Call ErrorHandler.RegistrarErro("ResetarSistema", Err)
    MsgBox "❌ Erro no reset: " & Err.Description, vbCritical
End Sub

' === ATUALIZAR DASHBOARD ===
Public Sub AtualizarDashboard()
    On Error GoTo TratarErro
    
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Worksheets("Dashboard")
    
    ' Calcular estatísticas
    Dim totalPedidos As Long
    Dim totalVendas As Double
    Dim totalClientes As Long
    
    ' Contar pedidos
    Dim i As Long
    For i = 1 To ThisWorkbook.Worksheets.Count
        If Left(ThisWorkbook.Worksheets(i).Name, 7) = "Pedido_" Then
            totalPedidos = totalPedidos + 1
            
            ' Somar vendas
            On Error Resume Next
            totalVendas = totalVendas + CDbl(ThisWorkbook.Worksheets(i).Range("K25").Value)
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
    wsDashboard.Range("B6:D8").Value = "Total de Pedidos" & vbCrLf & totalPedidos
    wsDashboard.Range("E6:G8").Value = "Total de Vendas" & vbCrLf & Format(totalVendas, "R$ #,##0.00")
    wsDashboard.Range("H6:J8").Value = "Total de Clientes" & vbCrLf & totalClientes
    
    ' Atualizar últimos pedidos
    Dim ultimosPedidos As String
    ultimosPedidos = PedidoManager.ObterUltimosPedidos(5)
    
    If ultimosPedidos <> "" Then
        Dim pedidos() As String
        pedidos = Split(ultimosPedidos, vbCrLf)
        
        Dim linha As Long
        linha = 23
        
        For i = 0 To UBound(pedidos) - 1
            If i < 5 And Trim(pedidos(i)) <> "" Then
                Dim dadosPedido() As String
                dadosPedido = Split(pedidos(i), "|")
                
                If UBound(dadosPedido) >= 4 Then
                    wsDashboard.Cells(linha, 2).Value = dadosPedido(0) ' Número
                    wsDashboard.Cells(linha, 3).Value = dadosPedido(1) ' Cliente
                    wsDashboard.Cells(linha, 5).Value = dadosPedido(2) ' Data
                    wsDashboard.Cells(linha, 6).Value = dadosPedido(3) ' Valor
                    wsDashboard.Cells(linha, 7).Value = dadosPedido(4) ' Status
                    
                    linha = linha + 1
                End If
            End If
        Next i
    End If
    
    Exit Sub
TratarErro:
    Call ErrorHandler.RegistrarErro("AtualizarDashboard", Err)
End Sub

' === VALIDAR DADOS OBRIGATÓRIOS ===
Public Function ValidarDadosObrigatorios(nomeCliente As String, formaPagamento As String, listaProdutos As MSForms.ListBox) As Boolean
    On Error GoTo TratarErro
    
    ValidarDadosObrigatorios = True
    
    ' Validar nome do cliente
    If Trim(nomeCliente) = "" Then
        MsgBox "⚠️ Nome do cliente é obrigatório!", vbExclamation, "Campo Obrigatório"
        ValidarDadosObrigatorios = False
        Exit Function
    End If
    
    ' Validar forma de pagamento
    If Trim(formaPagamento) = "" Then
        MsgBox "⚠️ Forma de pagamento é obrigatória!", vbExclamation, "Campo Obrigatório"
        ValidarDadosObrigatorios = False
        Exit Function
    End If
    
    ' Validar produtos
    If listaProdutos.ListCount = 0 Then
        MsgBox "⚠️ Adicione pelo menos um produto!", vbExclamation, "Produtos Obrigatórios"
        ValidarDadosObrigatorios = False
        Exit Function
    End If
    
    Exit Function
TratarErro:
    ValidarDadosObrigatorios = False
End Function