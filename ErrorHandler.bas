' ====================================================================
' MÓDULO ERROR HANDLER - SISTEMA PDV MADEIREIRA MARIA LUZIA
' Responsável pelo tratamento centralizado de erros
' ====================================================================

Option Explicit

' === REGISTRAR ERRO ===
Public Sub RegistrarErro(procedimento As String, erro As ErrObject)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Log_Erros")
    
    ' Obter próxima linha disponível
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Registrar erro
    With ws
        .Cells(ultimaLinha, 1).Value = Format(Now(), "dd/mm/yyyy hh:mm:ss")
        .Cells(ultimaLinha, 2).Value = procedimento
        .Cells(ultimaLinha, 3).Value = erro.Number
        .Cells(ultimaLinha, 4).Value = erro.Description
        .Cells(ultimaLinha, 5).Value = Environ("USERNAME")
    End With
    
    ' Limitar log a 1000 entradas
    If ultimaLinha > 1001 Then
        ws.Rows(2).Delete ' Remove a entrada mais antiga
    End If
    
    On Error GoTo 0
End Sub

' === TRATAR ERRO GENÉRICO ===
Public Sub TratarErro(procedimento As String, erro As ErrObject)
    Call RegistrarErro(procedimento, erro)
    
    Dim mensagem As String
    mensagem = "❌ Erro no procedimento: " & procedimento & vbCrLf & vbCrLf & _
               "Número: " & erro.Number & vbCrLf & _
               "Descrição: " & erro.Description & vbCrLf & vbCrLf & _
               "O erro foi registrado no log do sistema."
    
    MsgBox mensagem, vbCritical, "Erro do Sistema"
End Sub

' === LIMPAR LOG DE ERROS ===
Public Sub LimparLogErros()
    On Error GoTo TratarErro
    
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("⚠️ Tem certeza que deseja limpar todo o log de erros?", _
                     vbYesNo + vbQuestion, "Confirmar Limpeza")
    
    If resposta = vbYes Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("Log_Erros")
        
        ' Manter apenas o cabeçalho
        ws.Range("A2:E" & ws.Rows.Count).Clear
        
        MsgBox "✅ Log de erros limpo com sucesso!", vbInformation
    End If
    
    Exit Sub
TratarErro:
    MsgBox "❌ Erro ao limpar log: " & Err.Description, vbCritical
End Sub

' === EXPORTAR LOG DE ERROS ===
Public Sub ExportarLogErros()
    On Error GoTo TratarErro
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Log_Erros")
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If ultimaLinha < 2 Then
        MsgBox "ℹ️ Não há erros registrados para exportar.", vbInformation
        Exit Sub
    End If
    
    ' Criar arquivo de texto com log
    Dim nomeArquivo As String
    nomeArquivo = ThisWorkbook.Path & "\Log_Erros_" & Format(Now(), "yyyymmdd_hhmmss") & ".txt"
    
    Dim arquivo As Integer
    arquivo = FreeFile
    
    Open nomeArquivo For Output As #arquivo
    
    ' Cabeçalho
    Print #arquivo, "LOG DE ERROS - SISTEMA PDV MADEIREIRA MARIA LUZIA"
    Print #arquivo, "Exportado em: " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
    Print #arquivo, String(60, "=")
    Print #arquivo, ""
    
    ' Dados
    Dim i As Long
    For i = 2 To ultimaLinha
        Print #arquivo, "Data/Hora: " & ws.Cells(i, 1).Value
        Print #arquivo, "Procedimento: " & ws.Cells(i, 2).Value
        Print #arquivo, "Número do Erro: " & ws.Cells(i, 3).Value
        Print #arquivo, "Descrição: " & ws.Cells(i, 4).Value
        Print #arquivo, "Usuário: " & ws.Cells(i, 5).Value
        Print #arquivo, String(40, "-")
    Next i
    
    Close #arquivo
    
    MsgBox "✅ Log exportado com sucesso!" & vbCrLf & _
           "Arquivo: " & nomeArquivo, vbInformation, "Exportação Concluída"
    
    Exit Sub
TratarErro:
    If arquivo > 0 Then Close #arquivo
    MsgBox "❌ Erro ao exportar log: " & Err.Description, vbCritical
End Sub

' === VERIFICAR INTEGRIDADE DO SISTEMA ===
Public Function VerificarIntegridadeSistema() As String
    On Error GoTo TratarErro
    
    Dim resultado As String
    resultado = "VERIFICAÇÃO DE INTEGRIDADE DO SISTEMA" & vbCrLf & vbCrLf
    
    ' Verificar planilhas obrigatórias
    Dim planilhasObrigatorias As Variant
    planilhasObrigatorias = Array("Dashboard", "Clientes", "Produtos", "Template_Pedido", "Controle", "Log_Erros")
    
    Dim i As Integer
    For i = 0 To UBound(planilhasObrigatorias)
        On Error Resume Next
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(planilhasObrigatorias(i))
        
        If ws Is Nothing Then
            resultado = resultado & "❌ Planilha '" & planilhasObrigatorias(i) & "' não encontrada!" & vbCrLf
        Else
            resultado = resultado & "✅ Planilha '" & planilhasObrigatorias(i) & "' OK" & vbCrLf
        End If
        
        Set ws = Nothing
        On Error GoTo TratarErro
    Next i
    
    resultado = resultado & vbCrLf
    
    ' Verificar dados básicos
    On Error Resume Next
    Dim wsClientes As Worksheet, wsProdutos As Worksheet
    Set wsClientes = ThisWorkbook.Worksheets("Clientes")
    Set wsProdutos = ThisWorkbook.Worksheets("Produtos")
    
    If Not wsClientes Is Nothing Then
        Dim totalClientes As Long
        totalClientes = wsClientes.Cells(wsClientes.Rows.Count, "A").End(xlUp).Row - 1
        resultado = resultado & "📊 Total de Clientes: " & IIf(totalClientes > 0, totalClientes, 0) & vbCrLf
    End If
    
    If Not wsProdutos Is Nothing Then
        Dim totalProdutos As Long
        totalProdutos = wsProdutos.Cells(wsProdutos.Rows.Count, "A").End(xlUp).Row - 1
        resultado = resultado & "📦 Total de Produtos: " & IIf(totalProdutos > 0, totalProdutos, 0) & vbCrLf
    End If
    
    On Error GoTo TratarErro
    
    resultado = resultado & vbCrLf & "Verificação concluída em: " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
    VerificarIntegridadeSistema = resultado
    
    Exit Function
TratarErro:
    VerificarIntegridadeSistema = "Erro durante verificação: " & Err.Description
End Function

' === BACKUP AUTOMÁTICO ===
Public Sub CriarBackupAutomatico()
    On Error GoTo TratarErro
    
    Dim nomeBackup As String
    nomeBackup = ThisWorkbook.Path & "\Backup_PDV_" & Format(Now(), "yyyymmdd_hhmmss") & ".xlsm"
    
    ' Salvar cópia
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs nomeBackup
    Application.DisplayAlerts = True
    
    MsgBox "✅ Backup criado com sucesso!" & vbCrLf & _
           "Arquivo: " & nomeBackup, vbInformation, "Backup Automático"
    
    Exit Sub
TratarErro:
    Application.DisplayAlerts = True
    Call RegistrarErro("CriarBackupAutomatico", Err)
    MsgBox "❌ Erro ao criar backup: " & Err.Description, vbCritical
End Sub

' === VALIDAR DADOS DE ENTRADA ===
Public Function ValidarDados(valor As String, tipo As String) As Boolean
    On Error GoTo TratarErro
    
    Select Case UCase(tipo)
        Case "NUMERICO"
            ValidarDados = IsNumeric(valor)
            
        Case "TEXTO"
            ValidarDados = (Trim(valor) <> "")
            
        Case "EMAIL"
            ValidarDados = (InStr(valor, "@") > 0 And InStr(valor, ".") > 0)
            
        Case "TELEFONE"
            ' Validação básica de telefone
            Dim telefone As String
            telefone = Replace(Replace(Replace(Replace(valor, "(", ""), ")", ""), "-", ""), " ", "")
            ValidarDados = (Len(telefone) >= 10 And IsNumeric(telefone))
            
        Case "CEP"
            ' Validação básica de CEP
            Dim cep As String
            cep = Replace(valor, "-", "")
            ValidarDados = (Len(cep) = 8 And IsNumeric(cep))
            
        Case Else
            ValidarDados = True
    End Select
    
    Exit Function
TratarErro:
    ValidarDados = False
End Function

' === FORMATAR VALOR MONETÁRIO ===
Public Function FormatarMoeda(valor As Double) As String
    On Error GoTo TratarErro
    
    FormatarMoeda = Format(valor, "R$ #,##0.00")
    
    Exit Function
TratarErro:
    FormatarMoeda = "R$ 0,00"
End Function

' === CONVERTER TEXTO PARA VALOR ===
Public Function ConverterTextoParaValor(texto As String) As Double
    On Error GoTo TratarErro
    
    ' Remover formatação monetária
    texto = Replace(Replace(texto, "R$", ""), " ", "")
    texto = Replace(texto, ".", "")
    texto = Replace(texto, ",", ".")
    
    If IsNumeric(texto) Then
        ConverterTextoParaValor = CDbl(texto)
    Else
        ConverterTextoParaValor = 0
    End If
    
    Exit Function
TratarErro:
    ConverterTextoParaValor = 0
End Function