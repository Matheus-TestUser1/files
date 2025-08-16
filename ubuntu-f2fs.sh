'====================================================================
' SISTEMA PESQUISA PRODUTOS MADEIREIRA - VERSÃO COM PREÇO
' Data: 2025-08-16 - Busca pelos cabeçalhos: ID do produto, produto, material, dimensões, seção, preço
' Funciona independente da posição das colunas na planilha
'====================================================================

Option Explicit

'====================================================================
' VARIÁVEIS GLOBAIS
'====================================================================
Private QuantidadeAtual As Double
Private Const QUANTIDADE_MIN As Double = 1
Private Const QUANTIDADE_MAX As Double = 999

Private dictProdutos As Object
Private dictSelecionados As Object

' Posições das colunas (encontradas pelos cabeçalhos)
Private colID As Long
Private colProduto As Long
Private colMaterial As Long
Private colDimensoes As Long
Private colSecao As Long
Private colPreco As Long  ' NOVA COLUNA PREÇO

' Controles dinâmicos das setinhas
Dim WithEvents imgMais As MSForms.Image
Dim WithEvents imgMenos As MSForms.Image
Dim WithEvents lblMais As MSForms.Label
Dim WithEvents lblMenos As MSForms.Label

' Estados do sistema
Private placeholderAtivo As Boolean
Private atualizandoInterface As Boolean
Private ultimoTermoPesquisa As String

'====================================================================
' INICIALIZAÇÃO
'====================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErroInicializacao
    
    ' Configurar sistema
    Set dictProdutos = CreateObject("Scripting.Dictionary")
    Set dictSelecionados = CreateObject("Scripting.Dictionary")
    QuantidadeAtual = QUANTIDADE_MIN
    placeholderAtivo = False
    atualizandoInterface = False
    ultimoTermoPesquisa = ""
    
    ' Encontrar posições das colunas pelos cabeçalhos
    If Not EncontrarColunas() Then
        MsgBox "Erro: Cabeçalhos não encontrados na planilha 'Produtos'!" & vbCrLf & _
               "Certifique-se que existem: 'ID do produto', 'produto', 'material', 'dimensões', 'seção', 'preço'", vbCritical
        Unload Me
        Exit Sub
    End If
    
    ' Configurar interface
    Call ConfigurarFormulario
    Call ConfigurarListas
    Call ConfigurarPlaceholder
    Call CarregarProdutos
    Call CriarSetinhas
    
    ' Focar no campo de pesquisa
    If ControleExiste("txtPesquisa") Then Me.txtPesquisa.SetFocus
    
    Debug.Print "Sistema inicializado com sucesso!"
    Exit Sub
    
ErroInicializacao:
    MsgBox "ERRO na inicialização: " & Err.Description, vbCritical
    Unload Me
End Sub

'====================================================================
' ENCONTRAR COLUNAS PELOS CABEÇALHOS - ATUALIZADO COM PREÇO
'====================================================================
Private Function EncontrarColunas() As Boolean
    On Error GoTo ErroEncontrar
    
    Dim ws As Worksheet
    Set ws = Nothing
    
    On Error Resume Next
    Set ws = Worksheets("Produtos")
    On Error GoTo ErroEncontrar
    
    If ws Is Nothing Then
        EncontrarColunas = False
        Exit Function
    End If
    
    ' Inicializar posições
    colID = 0
    colProduto = 0
    colMaterial = 0
    colDimensoes = 0
    colSecao = 0
    colPreco = 0  ' NOVA COLUNA
    
    ' Procurar pelos cabeçalhos na primeira linha
    Dim ultimaColuna As Long
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim i As Long
    For i = 1 To ultimaColuna
        Dim cabecalho As String
        cabecalho = Trim(LCase(ws.Cells(1, i).Value))
        
        Select Case cabecalho
            Case "id do produto"
                colID = i
            Case "produto"
                colProduto = i
            Case "material"
                colMaterial = i
            Case "dimensões", "dimensoes"
                colDimensoes = i
            Case "seção", "secao"
                colSecao = i
            Case "preço", "preco", "valor"  ' NOVA COLUNA - aceita variações
                colPreco = i
        End Select
    Next i
    
    ' Verificar se encontrou todas as colunas (AGORA COM PREÇO)
    If colID > 0 And colProduto > 0 And colMaterial > 0 And colDimensoes > 0 And colSecao > 0 And colPreco > 0 Then
        Debug.Print "Colunas encontradas - ID:" & colID & " Produto:" & colProduto & " Material:" & colMaterial & " Dimensões:" & colDimensoes & " Seção:" & colSecao & " Preço:" & colPreco
        EncontrarColunas = True
    Else
        Debug.Print "Colunas faltando - ID:" & colID & " Produto:" & colProduto & " Material:" & colMaterial & " Dimensões:" & colDimensoes & " Seção:" & colSecao & " Preço:" & colPreco
        EncontrarColunas = False
    End If
    
    Exit Function
    
ErroEncontrar:
    EncontrarColunas = False
End Function

Private Sub ConfigurarFormulario()
    With Me
        .caption = "Sistema de Produtos - Madeireira (Com Preços)"
        .BackColor = RGB(248, 249, 250)
        .Width = 900  ' Aumentado para acomodar coluna preço
        .height = 600
    End With
End Sub

Private Sub ConfigurarListas()
    ' Configurar lista de produtos - AGORA COM 6 COLUNAS: ID, Produto, Material, Dimensões, Seção, Preço
    If ControleExiste("lstProdutos") Then
        With Me.lstProdutos
            .BackColor = RGB(255, 255, 255)
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .ColumnCount = 6  ' ALTERADO DE 5 PARA 6
            .ColumnHeads = False
            .ColumnWidths = "70;120;100;100;80;90"  ' ID, Produto, Material, Dimensões, Seção, Preço
        End With
    End If
    
    ' Configurar lista de selecionados - AGORA COM 7 COLUNAS: ID, Produto, Material, Dimensões, Seção, Preço, Quantidade
    If ControleExiste("lstSelecionados") Then
        With Me.lstSelecionados
            .BackColor = RGB(255, 255, 255)
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .ColumnCount = 7  ' ALTERADO DE 6 PARA 7
            .ColumnHeads = False
            .ColumnWidths = "60;110;90;90;70;80;60"  ' ID, Produto, Material, Dimensões, Seção, Preço, Qtd
        End With
    End If
End Sub

'====================================================================
' SISTEMA DE SETINHAS - MANTIDO IGUAL
'====================================================================
Private Sub CriarSetinhas()
    On Error Resume Next
    
    ' Verificar se existe o campo de quantidade
    If Not ControleExiste("txtQuantidade") Then
        Debug.Print "Campo txtQuantidade não encontrado!"
        Exit Sub
    End If
    
    ' Configurar campo de quantidade
    With Me.txtQuantidade
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .Value = QuantidadeAtual
    End With
    
    ' Remover setinhas antigas se existirem
    Call RemoverSetinhasAntigas
    
    ' Criar setinha MAIS (+)
    Set imgMais = Me.Controls.Add("Forms.Image.1", "imgSetaMais")
    If Not imgMais Is Nothing Then
        With imgMais
            .Left = Me.txtQuantidade.Left + Me.txtQuantidade.Width + 3
            .Top = Me.txtQuantidade.Top
            .Width = 20
            .height = 12
            .BackColor = RGB(40, 167, 69)  ' Verde
            .BorderStyle = fmBorderStyleSingle
            .Visible = True
        End With
        
        Set lblMais = Me.Controls.Add("Forms.Label.1", "lblSetaMais")
        With lblMais
            .Left = imgMais.Left + 2
            .Top = imgMais.Top - 1
            .Width = imgMais.Width - 4
            .height = imgMais.height
            .caption = "+"
            .Font.Bold = True
            .Font.Size = 10
            .Forecolor = RGB(255, 255, 255)
            .BackStyle = fmBackStyleTransparent
            .TextAlign = fmTextAlignCenter
            .Visible = True
        End With
    End If
    
    ' Criar setinha MENOS (-)
    Set imgMenos = Me.Controls.Add("Forms.Image.1", "imgSetaMenos")
    If Not imgMenos Is Nothing Then
        With imgMenos
            .Left = Me.txtQuantidade.Left + Me.txtQuantidade.Width + 3
            .Top = Me.txtQuantidade.Top + 13
            .Width = 20
            .height = 12
            .BackColor = RGB(220, 53, 69)  ' Vermelho
            .BorderStyle = fmBorderStyleSingle
            .Visible = True
        End With
        
        Set lblMenos = Me.Controls.Add("Forms.Label.1", "lblSetaMenos")
        With lblMenos
            .Left = imgMenos.Left + 2
            .Top = imgMenos.Top - 1
            .Width = imgMenos.Width - 4
            .height = imgMenos.height
            .caption = "-"
            .Font.Bold = True
            .Font.Size = 10
            .Forecolor = RGB(255, 255, 255)
            .BackStyle = fmBackStyleTransparent
            .TextAlign = fmTextAlignCenter
            .Visible = True
        End With
    End If
    
    Debug.Print "Setinhas criadas com sucesso!"
    On Error GoTo 0
End Sub

Private Sub RemoverSetinhasAntigas()
    On Error Resume Next
    Dim i As Long
    For i = Me.Controls.Count - 1 To 0 Step -1
        Dim ctrlName As String
        ctrlName = Me.Controls(i).Name
        If InStr(ctrlName, "imgSeta") > 0 Or InStr(ctrlName, "lblSeta") > 0 Then
            Me.Controls.Remove ctrlName
        End If
    Next i
    On Error GoTo 0
End Sub

' EVENTOS DAS SETINHAS (IMAGENS)
Private Sub imgMais_Click()
    Debug.Print "Setinha Imagem + clicada!"
    Call AlterarQuantidade(1)
End Sub

Private Sub imgMenos_Click()
    Debug.Print "Setinha Imagem - clicada!"
    Call AlterarQuantidade(-1)
End Sub

' EVENTOS DAS SETINHAS (LABELS)
Private Sub lblMais_Click()
    Debug.Print "Setinha Label + clicada!"
    Call AlterarQuantidade(1)
End Sub

Private Sub lblMenos_Click()
    Debug.Print "Setinha Label - clicada!"
    Call AlterarQuantidade(-1)
End Sub

Private Sub AlterarQuantidade(incremento As Integer)
    Dim novaQuantidade As Double
    novaQuantidade = QuantidadeAtual + incremento
    
    ' Aplicar limites
    If novaQuantidade < QUANTIDADE_MIN Then
        novaQuantidade = QUANTIDADE_MIN
    ElseIf novaQuantidade > QUANTIDADE_MAX Then
        novaQuantidade = QUANTIDADE_MAX
    End If
    
    ' Atualizar se mudou
    If novaQuantidade <> QuantidadeAtual Then
        QuantidadeAtual = novaQuantidade
        Call AtualizarCampoQuantidade
        Debug.Print "Quantidade alterada para: " & QuantidadeAtual
    End If
End Sub

Private Sub AtualizarCampoQuantidade()
    If ControleExiste("txtQuantidade") Then
        atualizandoInterface = True
        Me.txtQuantidade.Value = QuantidadeAtual
        atualizandoInterface = False
    End If
End Sub

'====================================================================
' CARREGAR PRODUTOS - ATUALIZADO COM PREÇO
'====================================================================
Private Sub CarregarProdutos()
    On Error GoTo ErroCarregamento
    
    Dim ws As Worksheet
    Set ws = Nothing
    
    On Error Resume Next
    Set ws = Worksheets("Produtos")
    On Error GoTo ErroCarregamento
    
    If ws Is Nothing Then
        MsgBox "Planilha 'Produtos' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    dictProdutos.RemoveAll
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, colID).End(xlUp).Row
    
    If ultimaLinha < 2 Then
        MsgBox "Nenhum produto encontrado na planilha.", vbInformation
        Exit Sub
    End If
    
    Dim i As Long
    For i = 2 To ultimaLinha
        If Trim(ws.Cells(i, colID).Value) <> "" Then
            Dim codigo As String
            codigo = Trim(ws.Cells(i, colID).Value)
            If Not dictProdutos.Exists(codigo) Then
                dictProdutos.Add codigo, i
            End If
        End If
    Next i
    
    Call ExibirTodosProdutos
    
    Debug.Print "Produtos carregados: " & dictProdutos.Count
    Exit Sub
    
ErroCarregamento:
    MsgBox "Erro ao carregar produtos: " & Err.Description, vbCritical
End Sub

Private Sub ExibirTodosProdutos()
    If Not ControleExiste("lstProdutos") Then Exit Sub
    
    Me.lstProdutos.Clear
    
    Dim ws As Worksheet
    Set ws = Worksheets("Produtos")
    
    Dim chave As Variant
    For Each chave In dictProdutos.Keys
        Dim linha As Long
        linha = dictProdutos(chave)
        
        With Me.lstProdutos
            .AddItem
            .List(.ListCount - 1, 0) = ws.Cells(linha, colID).Value        ' ID do produto
            .List(.ListCount - 1, 1) = ws.Cells(linha, colProduto).Value   ' Produto
            .List(.ListCount - 1, 2) = ws.Cells(linha, colMaterial).Value  ' Material
            .List(.ListCount - 1, 3) = ws.Cells(linha, colDimensoes).Value ' Dimensões
            .List(.ListCount - 1, 4) = ws.Cells(linha, colSecao).Value     ' Seção
            .List(.ListCount - 1, 5) = FormatarPreco(ws.Cells(linha, colPreco).Value) ' PREÇO FORMATADO
        End With
    Next chave
End Sub

'====================================================================
' FUNÇÃO PARA FORMATAR PREÇO
'====================================================================
Private Function FormatarPreco(valor As Variant) As String
    On Error Resume Next
    
    If IsNumeric(valor) And valor <> 0 Then
        FormatarPreco = "R$ " & Format(CDbl(valor), "0.00")
    Else
        FormatarPreco = "R$ 0,00"
    End If
    
    On Error GoTo 0
End Function

'====================================================================
' SISTEMA DE PESQUISA - ATUALIZADO COM PREÇO
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

Private Sub txtPesquisa_Change()
    If atualizandoInterface Or placeholderAtivo Then Exit Sub
    
    Dim termo As String
    termo = Trim(Me.txtPesquisa.Value)
    
    If termo = ultimoTermoPesquisa Then Exit Sub
    ultimoTermoPesquisa = termo
    
    If Len(termo) >= 2 Then
        Call PesquisarProdutos(termo)
    ElseIf Len(termo) = 0 Then
        Call ExibirTodosProdutos
    End If
End Sub

Private Sub PesquisarProdutos(termo As String)
    If Not ControleExiste("lstProdutos") Then Exit Sub
    
    Me.lstProdutos.Clear
    
    Dim ws As Worksheet
    Set ws = Worksheets("Produtos")
    
    Dim termoUpper As String
    termoUpper = UCase(termo)
    
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, colID).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To ultimaLinha
        ' Pesquisar nas 6 colunas usando posições encontradas pelos cabeçalhos (INCLUINDO PREÇO)
        If InStr(1, UCase(ws.Cells(i, colID).Value), termoUpper) > 0 Or _
           InStr(1, UCase(ws.Cells(i, colProduto).Value), termoUpper) > 0 Or _
           InStr(1, UCase(ws.Cells(i, colMaterial).Value), termoUpper) > 0 Or _
           InStr(1, UCase(ws.Cells(i, colDimensoes).Value), termoUpper) > 0 Or _
           InStr(1, UCase(ws.Cells(i, colSecao).Value), termoUpper) > 0 Or _
           InStr(1, UCase(ws.Cells(i, colPreco).Value), termoUpper) > 0 Then
            
            With Me.lstProdutos
                .AddItem
                .List(.ListCount - 1, 0) = ws.Cells(i, colID).Value        ' ID
                .List(.ListCount - 1, 1) = ws.Cells(i, colProduto).Value   ' Produto
                .List(.ListCount - 1, 2) = ws.Cells(i, colMaterial).Value  ' Material
                .List(.ListCount - 1, 3) = ws.Cells(i, colDimensoes).Value ' Dimensões
                .List(.ListCount - 1, 4) = ws.Cells(i, colSecao).Value     ' Seção
                .List(.ListCount - 1, 5) = FormatarPreco(ws.Cells(i, colPreco).Value) ' PREÇO
            End With
        End If
    Next i
End Sub

'====================================================================
' BOTÃO ADICIONAR - ADAPTADO PARA 6 COLUNAS COM PREÇO
'====================================================================
Private Sub btnAdicionar_Click()
    On Error GoTo ErroAdicionar
    
    Debug.Print "Botão Adicionar clicado!"
    
    ' Verificar se há produto selecionado
    If Not ControleExiste("lstProdutos") Then
        MsgBox "Lista de produtos não encontrada!", vbExclamation
        Exit Sub
    End If
    
    If Me.lstProdutos.ListIndex = -1 Then
        MsgBox "Selecione um produto da lista primeiro.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar quantidade
    If QuantidadeAtual < QUANTIDADE_MIN Then
        MsgBox "Quantidade deve ser pelo menos " & QUANTIDADE_MIN, vbExclamation
        Exit Sub
    End If
    
    ' Obter dados do produto selecionado das 6 colunas
    Dim idProduto As String
    Dim nomeProduto As String
    Dim materialProduto As String
    Dim dimensoesProduto As String
    Dim secaoProduto As String
    Dim precoProduto As String
    
    With Me.lstProdutos
        idProduto = .List(.ListIndex, 0)        ' ID do produto
        nomeProduto = .List(.ListIndex, 1)      ' Produto
        materialProduto = .List(.ListIndex, 2)  ' Material
        dimensoesProduto = .List(.ListIndex, 3) ' Dimensões
        secaoProduto = .List(.ListIndex, 4)     ' Seção
        precoProduto = .List(.ListIndex, 5)     ' PREÇO FORMATADO
    End With
    
    ' Verificar se produto já foi adicionado
    If dictSelecionados.Exists(idProduto) Then
        If MsgBox("Produto já foi adicionado. Deseja alterar a quantidade?", vbYesNo + vbQuestion) = vbYes Then
            dictSelecionados(idProduto) = QuantidadeAtual
            Call AtualizarListaSelecionados
        End If
        Exit Sub
    End If
    
    ' Verificar se existe lista de selecionados
    If Not ControleExiste("lstSelecionados") Then
        MsgBox "Lista de selecionados não encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Adicionar à lista visual com as 6 colunas + quantidade
    With Me.lstSelecionados
        .AddItem
        Dim novoIndex As Long
        novoIndex = .ListCount - 1
        
        .List(novoIndex, 0) = idProduto
        .List(novoIndex, 1) = nomeProduto
        .List(novoIndex, 2) = materialProduto
        .List(novoIndex, 3) = dimensoesProduto
        .List(novoIndex, 4) = secaoProduto
        .List(novoIndex, 5) = precoProduto       ' PREÇO
        .List(novoIndex, 6) = QuantidadeAtual    ' QUANTIDADE NA ÚLTIMA COLUNA
    End With
    
    ' Adicionar ao dicionário
    dictSelecionados.Add idProduto, QuantidadeAtual
    
    ' Resetar quantidade para próximo produto
    QuantidadeAtual = QUANTIDADE_MIN
    Call AtualizarCampoQuantidade
    
    ' Atualizar totais
    Call AtualizarTotais
    
    Debug.Print "Produto adicionado: " & idProduto & " - Qtd: " & dictSelecionados(idProduto)
    Exit Sub
    
ErroAdicionar:
    MsgBox "Erro ao adicionar produto: " & Err.Description, vbCritical
    Debug.Print "Erro em btnAdicionar_Click: " & Err.Description
End Sub

'====================================================================
' CONTROLES DA LISTA DE SELECIONADOS - ATUALIZADO COM PREÇO
'====================================================================
Private Sub btnRemover_Click()
    If Not ControleExiste("lstSelecionados") Then Exit Sub
    
    If Me.lstSelecionados.ListIndex = -1 Then
        MsgBox "Selecione um produto para remover.", vbExclamation
        Exit Sub
    End If
    
    ' Obter ID do produto
    Dim codigo As String
    codigo = Me.lstSelecionados.List(Me.lstSelecionados.ListIndex, 0)
    
    ' Remover do dicionário
    If dictSelecionados.Exists(codigo) Then
        dictSelecionados.Remove codigo
    End If
    
    ' Remover da lista visual
    Me.lstSelecionados.RemoveItem Me.lstSelecionados.ListIndex
    
    ' Atualizar totais
    Call AtualizarTotais
End Sub

Private Sub btnLimparTudo_Click()
    If MsgBox("Limpar todos os produtos selecionados?", vbYesNo + vbQuestion) = vbYes Then
        If ControleExiste("lstSelecionados") Then Me.lstSelecionados.Clear
        dictSelecionados.RemoveAll
        Call AtualizarTotais
    End If
End Sub

Private Sub AtualizarListaSelecionados()
    ' Atualizar lista visual baseada no dicionário
    If Not ControleExiste("lstSelecionados") Then Exit Sub
    
    Me.lstSelecionados.Clear
    
    Dim ws As Worksheet
    Set ws = Worksheets("Produtos")
    
    Dim codigo As Variant
    For Each codigo In dictSelecionados.Keys
        If dictProdutos.Exists(codigo) Then
            Dim linha As Long
            linha = dictProdutos(codigo)
            
            Dim quantidade As Double
            quantidade = dictSelecionados(codigo)
            
            With Me.lstSelecionados
                .AddItem
                .List(.ListCount - 1, 0) = codigo                               ' ID
                .List(.ListCount - 1, 1) = ws.Cells(linha, colProduto).Value   ' Produto
                .List(.ListCount - 1, 2) = ws.Cells(linha, colMaterial).Value  ' Material
                .List(.ListCount - 1, 3) = ws.Cells(linha, colDimensoes).Value ' Dimensões
                .List(.ListCount - 1, 4) = ws.Cells(linha, colSecao).Value     ' Seção
                .List(.ListCount - 1, 5) = FormatarPreco(ws.Cells(linha, colPreco).Value) ' PREÇO
                .List(.ListCount - 1, 6) = quantidade                           ' Quantidade
            End With
        End If
    Next codigo
    
    Call AtualizarTotais
End Sub

'====================================================================
' EVENTOS DE NAVEGAÇÃO - MANTIDOS IGUAIS
'====================================================================
Private Sub txtPesquisa_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 40 ' Seta para baixo
            If Me.lstProdutos.ListCount > 0 Then
                Me.lstProdutos.SetFocus
                If Me.lstProdutos.ListIndex = -1 Then
                    Me.lstProdutos.ListIndex = 0
                End If
            End If
        Case 13 ' Enter
            If Me.lstProdutos.ListCount > 0 And Me.lstProdutos.ListIndex >= 0 Then
                Call btnAdicionar_Click
            End If
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
        Case 13: Call btnAdicionar_Click  ' Enter
        Case 38: Call AlterarQuantidade(1)  ' Seta cima
        Case 40: Call AlterarQuantidade(-1) ' Seta baixo
    End Select
End Sub

'====================================================================
' FUNÇÕES UTILITÁRIAS - ATUALIZADO PARA CALCULAR VALORES
'====================================================================
Private Sub AtualizarTotais()
    If Not ControleExiste("lstSelecionados") Then Exit Sub
    
    Dim totalItens As Double
    Dim valorTotal As Double
    Dim ws As Worksheet
    Set ws = Worksheets("Produtos")
    
    ' Calcular totais percorrendo a lista de selecionados
    Dim i As Long
    For i = 0 To Me.lstSelecionados.ListCount - 1
        Dim quantidade As Double
        quantidade = CDbl(Me.lstSelecionados.List(i, 6))  ' Coluna 6 = Quantidade
        totalItens = totalItens + quantidade
        
        ' Calcular valor total (quantidade x preço unitário)
        Dim idProduto As String
        idProduto = Me.lstSelecionados.List(i, 0)  ' ID do produto
        
        If dictProdutos.Exists(idProduto) Then
            Dim linha As Long
            linha = dictProdutos(idProduto)
            
            Dim precoUnitario As Double
            precoUnitario = 0
            If IsNumeric(ws.Cells(linha, colPreco).Value) Then
                precoUnitario = CDbl(ws.Cells(linha, colPreco).Value)
            End If
            
            valorTotal = valorTotal + (quantidade * precoUnitario)
        End If
    Next i
    
    ' Atualizar labels se existirem
    If ControleExiste("lblTotalItens") Then
        Me.lblTotalItens.caption = "Total: " & totalItens & " itens"
    End If
    
    If ControleExiste("lblValorTotal") Then
        Me.lblValorTotal.caption = "Valor Total: " & Format(valorTotal, "R$ #,##0.00") & " | Tipos: " & Me.lstSelecionados.ListCount & " diferentes"
    End If
End Sub

Private Function InterpretarQuantidade(texto As String) As Double
    On Error Resume Next
    
    Dim valor As Double
    texto = Trim(Replace(texto, ",", "."))
    
    If Not IsNumeric(texto) Then
        valor = QUANTIDADE_MIN
    Else
        valor = CDbl(texto)
    End If
    
    If Err.Number <> 0 Or valor < QUANTIDADE_MIN Then
        valor = QUANTIDADE_MIN
    ElseIf valor > QUANTIDADE_MAX Then
        valor = QUANTIDADE_MAX
    End If
    
    InterpretarQuantidade = valor
    On Error GoTo 0
End Function

Private Function ControleExiste(nome As String) As Boolean
    On Error Resume Next
    Dim ctrl As Control
    Set ctrl = Me.Controls(nome)
    ControleExiste = Not (ctrl Is Nothing)
    On Error GoTo 0
End Function

'====================================================================
' FINALIZAÇÃO
'====================================================================
Private Sub btnFechar_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set dictProdutos = Nothing
    Set dictSelecionados = Nothing
End Sub
