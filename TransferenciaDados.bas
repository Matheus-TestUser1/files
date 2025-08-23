' ====================================================================
' MÓDULO PARA TRANSFERÊNCIA DE DADOS ENTRE FORMULÁRIOS
' Data/Hora: 2025-01-27
' Usuário: Sistema PDV Madeireira Maria Luzia
' ====================================================================

Option Explicit

' Variáveis globais para armazenar dados dos produtos
Public ProdutosSelecionadosGlobal As String
Public QuantidadeProdutosGlobal As Long
Public ValorTotalGlobal As Double
Public UsuarioProcessamento As String
Public DataHoraProcessamento As String
Public ClienteSelecionadoGlobal As String

' Função para armazenar dados dos produtos selecionados
Public Sub ArmazenarDadosProdutos(dados As String)
    ProdutosSelecionadosGlobal = dados
    UsuarioProcessamento = "Sistema PDV"
    DataHoraProcessamento = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    ' Calcular quantidade e valor total
    If dados <> "" Then
        Dim produtos() As String
        produtos = Split(dados, ";")
        QuantidadeProdutosGlobal = 0
        ValorTotalGlobal = 0
        
        Dim i As Long
        For i = 0 To UBound(produtos)
            If Trim(produtos(i)) <> "" Then
                Dim dadosProduto() As String
                dadosProduto = Split(produtos(i), "|")
                If UBound(dadosProduto) >= 6 Then
                    QuantidadeProdutosGlobal = QuantidadeProdutosGlobal + CLng(dadosProduto(4))
                    ValorTotalGlobal = ValorTotalGlobal + CDbl(dadosProduto(6))
                End If
            End If
        Next i
    End If
End Sub

' Função para armazenar dados do cliente selecionado
Public Sub ArmazenarDadosCliente(dadosCliente As String)
    ClienteSelecionadoGlobal = dadosCliente
End Sub

' Função para o frmPDVPrincipal receber os dados
Public Sub TransferirParaPDVPrincipal(produtosv1 As Object)
    On Error GoTo TratarErro
    
    If ProdutosSelecionadosGlobal = "" Then
        MsgBox "⚠️ Nenhum produto para transferir!", vbExclamation, "Transferência de Dados"
        Exit Sub
    End If
    
    ' Processar dados
    Dim produtos() As String
    Dim dadosProduto() As String
    Dim i As Long, adicionados As Long
    
    produtos = Split(ProdutosSelecionadosGlobal, ";")
    adicionados = 0
    
    For i = 0 To UBound(produtos) - 1 ' -1 porque o último elemento pode estar vazio
        If Trim(produtos(i)) <> "" Then
            dadosProduto = Split(produtos(i), "|")
            
            If UBound(dadosProduto) >= 6 Then ' Verificar se tem todos os dados
                With produtosv1
                    .AddItem
                    .List(.ListCount - 1, 0) = dadosProduto(0) ' Referencia
                    .List(.ListCount - 1, 1) = dadosProduto(1) ' Descrição do Item
                    .List(.ListCount - 1, 2) = dadosProduto(2) ' uni
                    .List(.ListCount - 1, 3) = dadosProduto(3) ' Valor
                    .List(.ListCount - 1, 4) = dadosProduto(4) ' Quant.
                    .List(.ListCount - 1, 5) = dadosProduto(5) ' Desc.
                    .List(.ListCount - 1, 6) = dadosProduto(6) ' Valor Total
                End With
                adicionados = adicionados + 1
            End If
        End If
    Next i
    
    ' Limpar dados globais
    ProdutosSelecionadosGlobal = ""
    
    ' Mostrar resultado
    MsgBox "✅ PRODUTOS RECEBIDOS COM SUCESSO!" & vbCrLf & vbCrLf & _
           "📦 Produtos adicionados: " & adicionados & vbCrLf & _
           "👤 Enviado por: " & UsuarioProcessamento & vbCrLf & _
           "🕐 Data/Hora: " & DataHoraProcessamento, vbInformation, "Transferência Concluída"
    
    Exit Sub
    
TratarErro:
    MsgBox "❌ Erro na transferência: " & Err.Description, vbExclamation, "Erro de Transferência"
End Sub

' Função para verificar se há dados pendentes
Public Function TemDadosPendentes() As Boolean
    TemDadosPendentes = (ProdutosSelecionadosGlobal <> "")
End Function

' Função para obter dados do cliente
Public Function ObterDadosCliente() As String
    ObterDadosCliente = ClienteSelecionadoGlobal
End Function

' Função para limpar todos os dados globais
Public Sub LimparDadosGlobais()
    ProdutosSelecionadosGlobal = ""
    QuantidadeProdutosGlobal = 0
    ValorTotalGlobal = 0
    ClienteSelecionadoGlobal = ""
    UsuarioProcessamento = ""
    DataHoraProcessamento = ""
End Sub

' Função para obter resumo dos dados armazenados
Public Function ObterResumo() As String
    If ProdutosSelecionadosGlobal = "" Then
        ObterResumo = "Nenhum dado armazenado"
    Else
        ObterResumo = "Produtos: " & QuantidadeProdutosGlobal & " | " & _
                     "Valor Total: R$ " & Format(ValorTotalGlobal, "#,##0.00") & " | " & _
                     "Processado em: " & DataHoraProcessamento
    End If
End Function