'====================================================================
' SISTEMA PESQUISA PRODUTOS MADEIREIRA - VERSÃO CORRIGIDA
' Data: 2025-08-16 - ENTREGA URGENTE
' Correções críticas aplicadas para estabilidade e performance
'====================================================================

Option Explicit

'====================================================================
' ESTRUTURAS DE DADOS OTIMIZADAS
'====================================================================


'====================================================================
' VARIÁVEIS GLOBAIS ORGANIZADAS
'====================================================================
' Sistema de quantidade
Private QuantidadeAtual As Double
Private Const QUANTIDADE_MIN As Double = 1
Private Const QUANTIDADE_MAX As Double = 999

' Cache de performance
Private dictProdutos As Object           ' Dictionary para lookup O(1)
Private dictSelecionados As Object      ' Evita duplicatas
Private arrProdutosFiltrados As Variant
' Controles dinâmicos
Dim WithEvents imgMais As MSForms.Image
Dim WithEvents imgMenos As MSForms.Image

' Estados do sistema
Private placeholderAtivo As Boolean
Private atualizandoInterface As Boolean
Private ultimoTermoPesquisa As String

'====================================================================
' INICIALIZAÇÃO SEGURA
'====================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    ' Inicializar estruturas de dados
    Call InicializarSistema
    
    ' Configurar interface
    Call ConfigurarInterface
    
    ' Carregar dados
    Call CarregarProdutosComCache
    
    ' Finalizar
    Call DefinirEstadoInicial
    
    Exit Sub
    
ErroInicializacao:
    MsgBox "ERRO CRÍTICO: " & Err.Description & vbCrLf & _
           "Sistema será encerrado.", vbCritical, "Falha na Inicialização"
    Unload Me
End Sub

Private Sub InicializarSistema()
    ' Criar dicionários para performance
    Set dictProdutos = CreateObject("Scripting.Dictionary")
    Set dictSelecionados = CreateObject("Scripting.Dictionary")
    
    ' Configurar variáveis
    QuantidadeAtual = QUANTIDADE_MIN
    placeholderAtivo = False
    atualizandoInterface = False
    ultimoTermoPesquisa = ""
End Sub

'====================================================================
' CONFIGURAÇÃO DE INTERFACE OTIMIZADA
'====================================================================
Private Sub ConfigurarInterface()
    ' Configurar formulário
    With Me
        .caption = "Madeireira Maria Luiza - Sistema de Produtos"
        .BackColor = RGB(248, 249, 250)
    End With
    
    ' Configurar listas com melhor performance
    Call ConfigurarListas
    
    ' Criar controles dinâmicos
    Call CriarControlesQuantidade
    
    ' Configurar placeholder
    Call ConfigurarPlaceholder
End Sub

Private Sub ConfigurarListas()
    If Not ControleExiste("lstProdutos") Then Exit Sub
    
    With Me.lstProdutos
        .BackColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ColumnCount = 5
        .ColumnWidths = "80;200;120;100;80"
    End With
    
    If ControleExiste("lstSelecionados") Then
        With Me.lstSelecionados
            .BackColor = RGB(255, 255, 255)
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .ColumnCount = 7
            .ColumnWidths = "60;180;80;80;60;100;80"
        End With
    End If
End Sub

'====================================================================
' SISTEMA DE SETINHAS OTIMIZADO
'====================================================================
Private Sub CriarControlesQuantidade()
    On Error Resume Next
    
    Dim txtQuant As Control
    Set txtQuant = Me.Controls("txtQuantidade")
    If txtQuant Is Nothing Then Exit Sub
    
    ' Configurar campo quantidade
    With txtQuant
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .TextAlign = 2  ' Centro
        .Value = QuantidadeAtual
    End With
    
    ' Remover controles antigos
    Call RemoverControlesAntigos
    
    ' Criar setas otimizadas
    Const LARGURA_SETA As Single = 18
    Const ALTURA_SETA As Single = 11
    
    ' Seta MAIS (+)
    Set imgMais = Me.Controls.Add("Forms.Image.1", "imgMais")
    With imgMais
        .Left = txtQuant.Left + txtQuant.Width + 2
        .Top = txtQuant.Top
        .Width = LARGURA_SETA
        .height = ALTURA_SETA
        .BackColor = RGB(40, 167, 69)
        .BorderStyle = 1
        .Visible = True
    End With
    
    ' Seta MENOS (-)
    Set imgMenos = Me.Controls.Add("Forms.Image.1", "imgMenos")
    With imgMenos
        .Left = txtQuant.Left + txtQuant.Width + 2
        .Top = txtQuant.Top + ALTURA_SETA + 1
        .Width = LARGURA_SETA
        .height = ALTURA_SETA
        .BackColor = RGB(220, 53, 69)
        .BorderStyle = 1
        .Visible = True
    End With
    
    ' Adicionar labels dos símbolos
    Call AdicionarSimbolosSetas(LARGURA_SETA, ALTURA_SETA)
    
    On Error GoTo 0
End Sub

Private Sub AdicionarSimbolosSetas(largura As Single, altura As Single)
    On Error Resume Next
    
    ' Label MAIS
    Dim lblMais As MSForms.Label
    Set lblMais = Me.Controls.Add("Forms.Label.1", "lblMais")
    With lblMais
        .Left = imgMais.Left + 1
        .Top = imgMais.Top
        .Width = largura - 2
        .height = altura
        .caption = "+"
        .Font.Bold = True
        .Forecolor = RGB(255, 255, 255)
        .BackStyle = 0
        .TextAlign = 2
        .Visible = True
    End With
    
    ' Label MENOS
    Dim lblMenos As MSForms.Label
    Set lblMenos = Me.Controls.Add("Forms.Label.1", "lblMenos")
    With lblMenos
        .Left = imgMenos.Left + 1
        .Top = imgMenos.Top
        .Width = largura - 2
        .height = altura
        .caption = "-"
        .Font.Bold = True
        .Forecolor = RGB(255, 255, 255)
        .BackStyle = 0
        .TextAlign = 2
        .Visible = True
    End With
    
    On Error GoTo 0
End Sub

Private Sub RemoverControlesAntigos()
    On Error Resume Next
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If InStr(ctrl.Name, "imgMais") > 0 Or InStr(ctrl.Name, "imgMenos") > 0 Or _
           InStr(ctrl.Name, "lblMais") > 0 Or InStr(ctrl.Name, "lblMenos") > 0 Then
            Me.Controls.Remove ctrl.Name
        End If
    Next ctrl
    On Error GoTo 0
End Sub

'====================================================================
' EVENTOS DAS SETINHAS
'====================================================================
Private Sub imgMais_Click()
    Call AlterarQuantidade(1)
End Sub

Private Sub imgMenos_Click()
    Call AlterarQuantidade(-1)
End Sub

Private Sub AlterarQuantidade(incremento As Integer)
    Dim novaQuantidade As Double
    novaQuantidade = QuantidadeAtual + incremento
    
    ' Validar limites
    If novaQuantidade < QUANTIDADE_MIN Then
        novaQuantidade = QUANTIDADE_MIN
    ElseIf novaQuantidade > QUANTIDADE_MAX Then
        novaQuantidade = QUANTIDADE_MAX
    End If
    
    ' Atualizar se mudou
    If novaQuantidade <> QuantidadeAtual Then
        QuantidadeAtual = novaQuantidade
        Call AtualizarCampoQuantidade
    End If
End Sub

Private Sub AtualizarCampoQuantidade()
    If Not ControleExiste("txtQuantidade") Then Exit Sub
    
    atualizandoInterface = True
    Me.txtQuantidade.Value = FormatarQuantidade(QuantidadeAtual)
    atualizandoInterface = False
End Sub

'====================================================================
' CARREGAMENTO OTIMIZADO DE PRODUTOS
'====================================================================
Private Sub CarregarProdutosComCache()
    On Error GoTo ErroCarregamento
    
    Dim ws As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim produto As Variant
    
    ' Verificar se planilha existe
    Set ws = Nothing
    Set ws = Worksheets("Produtos")
    If ws Is Nothing Then
        MsgBox "ERRO: Planilha 'Produtos' não encontrada!", vbCritical
        Exit Sub
    End If
    
    ' Limpar cache anterior
    dictProdutos.RemoveAll
    
    ' Obter range de dados
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ultimaLinha < 2 Then
        MsgBox "Nenhum produto encontrado na planilha.", vbInformation
        Exit Sub
    End If
    
    ' Carregar produtos no cache
    For i = 2 To ultimaLinha
        If Trim(ws.Cells(i, 1).Value) <> "" Then
            ' Preencher estrutura
            produto.codigo = Trim(ws.Cells(i, 1).Value)
            produto.nome = Trim(ws.Cells(i, 2).Value)
            produto.Material = Trim(ws.Cells(i, 3).Value)
            produto.Dimensoes = Trim(ws.Cells(i, 4).Value)
            produto.Preco = Val(ws.Cells(i, 6).Value)
            produto.Disponivel = True
            
            ' Adicionar ao cache (chave = código)
            If Not dictProdutos.Exists(produto.codigo) Then
                dictProdutos.Add produto.codigo, produto
            End If
        End If
    Next i
    
    ' Carregar lista inicial
    Call ExibirTodosProdutos
    
    Exit Sub
    
ErroCarregamento:
    MsgBox "Erro ao carregar produtos: " & Err.Description, vbCritical
End Sub

Private Sub ExibirTodosProdutos()
    If Not ControleExiste("lstProdutos") Then Exit Sub
    
    Me.lstProdutos.Clear
    
    Dim chave As Variant
    Dim produto As TipoProduto
    
    For Each chave In dictProdutos.Keys
        produto = dictProdutos(chave)
        With Me.lstProdutos
            .AddItem
            .List(.ListCount - 1, 0) = produto.codigo
            .List(.ListCount - 1, 1) = produto.nome
            .List(.ListCount - 1, 2) = produto.Material
            .List(.ListCount - 1, 3) = produto.Dimensoes
            .List(.ListCount - 1, 4) = FormatarMoeda(produto.Preco)
        End With
    Next chave
End Sub

'====================================================================
' SISTEMA DE PESQUISA OTIMIZADO
'====================================================================
Private Sub txtPesquisa_Change()
    If atualizandoInterface Then Exit Sub
    If placeholderAtivo Then Exit Sub
    
    Dim termo As String
    termo = Trim(Me.txtPesquisa.Value)
    
    ' Evitar pesquisas desnecessárias
    If termo = ultimoTermoPesquisa Then Exit Sub
    ultimoTermoPesquisa = termo
    
    If Len(termo) >= 2 Then
        Call PesquisarProdutosOtimizada(termo)
    ElseIf Len(termo) = 0 Then
        Call ExibirTodosProdutos
    End If
End Sub

Private Sub PesquisarProdutosOtimizada(termo As String)
    If Not ControleExiste("lstProdutos") Then Exit Sub
    
    Me.lstProdutos.Clear
    
    Dim chave As Variant
    Dim produto As Variant
    Dim termoUpper As String
    
    termoUpper = UCase(termo)
    
    ' Busca otimizada no cache
    For Each chave In dictProdutos.Keys
        produto = dictProdutos(chave)
        
        If InStr(1, UCase(produto.codigo), termoUpper) > 0 Or _
           InStr(1, UCase(produto.nome), termoUpper) > 0 Or _
           InStr(1, UCase(produto.Material), termoUpper) > 0 Then
            
            With Me.lstProdutos
                .AddItem
                .List(.ListCount - 1, 0) = produto.codigo
                .List(.ListCount - 1, 1) = produto.nome
                .List(.ListCount - 1, 2) = produto.Material
                .List(.ListCount - 1, 3) = produto.Dimensoes
                .List(.ListCount - 1, 4) = FormatarMoeda(produto.Preco)
            End With
        End If
    Next chave
End Sub

'====================================================================
' ADIÇÃO DE PRODUTOS COM VALIDAÇÃO
'====================================================================
Private Sub btnAdicionar_Click()
    On Error GoTo ErroAdicionar
    
    ' Validações
    If Me.lstProdutos.ListIndex = -1 Then
        MsgBox "Selecione um produto da lista.", vbExclamation
        Exit Sub
    End If
    
    If QuantidadeAtual < QUANTIDADE_MIN Then
        MsgBox "Quantidade deve ser pelo menos " & QUANTIDADE_MIN, vbExclamation
        Exit Sub
    End If
    
    ' Obter dados do produto selecionado
    Dim codigoProduto As String
    Dim nomeProduto As String
    Dim precoProduto As Double
    
    codigoProduto = Me.lstProdutos.List(Me.lstProdutos.ListIndex, 0)
    nomeProduto = Me.lstProdutos.List(Me.lstProdutos.ListIndex, 1)
    precoProduto = ExtrairValorMoeda(Me.lstProdutos.List(Me.lstProdutos.ListIndex, 4))
    
    ' Verificar duplicata
    If dictSelecionados.Exists(codigoProduto) Then
        MsgBox "Produto já foi adicionado. Use Editar para alterar quantidade.", vbInformation
        Exit Sub
    End If
    
    ' Adicionar à lista de selecionados
    If ControleExiste("lstSelecionados") Then
        With Me.lstSelecionados
            .AddItem
            .List(.ListCount - 1, 0) = codigoProduto
            .List(.ListCount - 1, 1) = nomeProduto
            .List(.ListCount - 1, 2) = Me.lstProdutos.List(Me.lstProdutos.ListIndex, 2) ' Material
            .List(.ListCount - 1, 3) = FormatarMoeda(precoProduto)
            .List(.ListCount - 1, 4) = FormatarQuantidade(QuantidadeAtual)
            .List(.ListCount - 1, 5) = IIf(ControleExiste("txtDescricao"), Me.txtDescricao.Value, "")
            .List(.ListCount - 1, 6) = FormatarMoeda(precoProduto * QuantidadeAtual)
        End With
    End If
    
    ' Registrar no dicionário
    dictSelecionados.Add codigoProduto, QuantidadeAtual
    
    ' Reset
    QuantidadeAtual = QUANTIDADE_MIN
    Call AtualizarCampoQuantidade
    
    If ControleExiste("txtDescricao") Then Me.txtDescricao.Value = ""
    
    ' Atualizar totais
    Call AtualizarTotais
    
    Exit Sub
    
ErroAdicionar:
    MsgBox "Erro ao adicionar produto: " & Err.Description, vbCritical
End Sub

'====================================================================
' CONTROLES DE LISTA
'====================================================================
Private Sub btnRemover_Click()
    If Not ControleExiste("lstSelecionados") Then Exit Sub
    If Me.lstSelecionados.ListIndex = -1 Then
        MsgBox "Selecione um produto para remover.", vbExclamation
        Exit Sub
    End If
    
    Dim codigo As String
    codigo = Me.lstSelecionados.List(Me.lstSelecionados.ListIndex, 0)
    
    ' Remover do dicionário
    If dictSelecionados.Exists(codigo) Then
        dictSelecionados.Remove codigo
    End If
    
    ' Remover da lista visual
    Me.lstSelecionados.RemoveItem Me.lstSelecionados.ListIndex
    
    Call AtualizarTotais
End Sub

Private Sub btnLimparTudo_Click()
    If MsgBox("Limpar todos os produtos selecionados?", vbYesNo + vbQuestion) = vbYes Then
        If ControleExiste("lstSelecionados") Then Me.lstSelecionados.Clear
        dictSelecionados.RemoveAll
        Call AtualizarTotais
    End If
End Sub

'====================================================================
' SISTEMA DE PLACEHOLDER
'====================================================================
Private Sub ConfigurarPlaceholder()
    If Not ControleExiste("txtPesquisa") Then Exit Sub
    
    With Me.txtPesquisa
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Value = "Digite para pesquisar produtos..."
        .Forecolor = RGB(128, 128, 128)
        .Font.Italic = True
    End With
    placeholderAtivo = True
End Sub

Private Sub txtPesquisa_Enter()
    If placeholderAtivo Then
        atualizandoInterface = True
        Me.txtPesquisa.Value = ""
        Me.txtPesquisa.Forecolor = RGB(0, 0, 0)
        Me.txtPesquisa.Font.Italic = False
        placeholderAtivo = False
        atualizandoInterface = False
    End If
End Sub

Private Sub txtPesquisa_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(Me.txtPesquisa.Value) = "" Then
        Call ConfigurarPlaceholder
    End If
End Sub

'====================================================================
' EVENTOS DE NAVEGAÇÃO
'====================================================================
Private Sub txtPesquisa_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 40 ' Seta baixo
            If Me.lstProdutos.ListCount > 0 Then
                Me.lstProdutos.SetFocus
                Me.lstProdutos.ListIndex = 0
            End If
        Case 13 ' Enter
            If Me.lstProdutos.ListCount > 0 Then
                Me.lstProdutos.SetFocus
                If Me.lstProdutos.ListIndex = -1 Then Me.lstProdutos.ListIndex = 0
            End If
        Case 27 ' ESC
            Me.txtPesquisa.Value = ""
            Call ConfigurarPlaceholder
            Call ExibirTodosProdutos
    End Select
End Sub

Private Sub lstProdutos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.lstProdutos.ListIndex >= 0 Then
        Call btnAdicionar_Click
    End If
End Sub

Private Sub txtQuantidade_Change()
    If atualizandoInterface Then Exit Sub
    
    Dim novaQtd As Double
    novaQtd = InterpretarQuantidade(Me.txtQuantidade.Value)
    
    If novaQtd <> QuantidadeAtual Then
        QuantidadeAtual = novaQtd
    End If
End Sub

Private Sub txtQuantidade_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 13: Call btnAdicionar_Click
        Case 38: Call AlterarQuantidade(1)
        Case 40: Call AlterarQuantidade(-1)
    End Select
End Sub

'====================================================================
' FUNÇÕES UTILITÁRIAS
'====================================================================
Private Sub AtualizarTotais()
    Dim totalItens As Long
    Dim valorTotal As Double
    Dim i As Long
    
    If ControleExiste("lstSelecionados") Then
        For i = 0 To Me.lstSelecionados.ListCount - 1
            totalItens = totalItens + InterpretarQuantidade(Me.lstSelecionados.List(i, 4))
            valorTotal = valorTotal + ExtrairValorMoeda(Me.lstSelecionados.List(i, 6))
        Next i
    End If
    
    If ControleExiste("lblTotalItens") Then
        Me.lblTotalItens.caption = "Total: " & totalItens & " itens"
    End If
    
    If ControleExiste("lblValorTotal") Then
        Me.lblValorTotal.caption = "Valor: " & FormatarMoeda(valorTotal)
    End If
End Sub

Private Function InterpretarQuantidade(texto As String) As Double
    Dim valor As Double
    On Error Resume Next
    
    texto = Trim(Replace(Replace(texto, ",", "."), " ", ""))
    valor = CDbl(texto)
    
    If Err.Number <> 0 Or valor < QUANTIDADE_MIN Then
        valor = QUANTIDADE_MIN
    ElseIf valor > QUANTIDADE_MAX Then
        valor = QUANTIDADE_MAX
    End If
    
    InterpretarQuantidade = valor
    On Error GoTo 0
End Function

Private Function FormatarQuantidade(valor As Double) As String
    If valor = Int(valor) Then
        FormatarQuantidade = Format(valor, "0")
    Else
        FormatarQuantidade = Format(valor, "0.000")
    End If
End Function

Private Function FormatarMoeda(valor As Double) As String
    FormatarMoeda = Format(valor, "R$ #,##0.00")
End Function

Private Function ExtrairValorMoeda(texto As String) As Double
    On Error Resume Next
    Dim limpo As String
    limpo = Replace(Replace(Replace(texto, "R$", ""), ".", ""), ",", ".")
    limpo = Trim(Replace(limpo, " ", ""))
    ExtrairValorMoeda = CDbl(limpo)
    If Err.Number <> 0 Then ExtrairValorMoeda = 0
    On Error GoTo 0
End Function

Private Function ControleExiste(nome As String) As Boolean
    On Error Resume Next
    Dim ctrl As Control
    Set ctrl = Me.Controls(nome)
    ControleExiste = Not (ctrl Is Nothing)
    On Error GoTo 0
End Function

Private Sub DefinirEstadoInicial()
    If ControleExiste("txtPesquisa") Then Me.txtPesquisa.SetFocus
    Call AtualizarTotais
End Sub

'====================================================================
' FINALIZAÇÃO
'====================================================================
Private Sub btnFechar_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Limpar objetos da memória
    Set dictProdutos = Nothing
    Set dictSelecionados = Nothing
End Sub

