VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenuStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dev: Matheus Nunes Menezes | Brazilian
'Email: math.nunesmenezes@gmail.com
'Chanel: https://www.youtube.com/@programandoemvba
'Creation Date: 10/12/2024 - 16/12/2024
'Tutorial: https://youtu.be/mGxdJuHyD40
'Version: 2.0

'====================================================================================================
'=====================================| Example code (ing) |=========================================
'====================================================================================================

'Dim Menu As New clsMenuStrip
'
'Private Sub UserForm_Initialize()
'    Call CreateMenu
'End Sub
'
'Sub CreateMenu()
'
'    With Menu
'
'        Set .UserForm = Me
'
'        '============== Main Menu 1: People Management ================
'        .AddMenu 1, "People Management", 120, 170
'                .AddSubmenu xlMenu, 1, 1, "Recruitment and Selection", , 150
'                        .AddItem xlSubmenu, 1, "Register Job Position", "Test", True
'                        .AddSubmenu xlSubmenu, 1, 2, "Manage Job Positions", , 120
'                                .AddItem xlSubmenu, 2, "Open Positions", "Test"
'                                .AddItem xlSubmenu, 2, "Closed Positions", "Test"
'                                .AddItem xlSubmenu, 2, "Selection History", "Test"
'                        .AddSubmenu xlSubmenu, 1, 3, "Resume Bank", True, 120
'                                .AddItem xlSubmenu, 3, "View Resumes", "Test"
'                                .AddItem xlSubmenu, 3, "Add New Resume", "Test"
'                                .AddItem xlSubmenu, 3, "Export Database", "Test"
'                        .AddItem xlSubmenu, 1, "Schedule Interviews", "Test"
'                .AddSubmenu xlMenu, 1, 4, "Training and Development", , 160
'                        .AddItem xlSubmenu, 4, "Create Training Program", "Test"
'                        .AddSubmenu xlSubmenu, 4, 5, "Manage Trainings", , 140
'                                        .AddItem xlSubmenu, 5, "Active Trainings", "Test"
'                                        .AddItem xlSubmenu, 5, "Completed Trainings", "Test"
'                                        .AddItem xlSubmenu, 5, "Participation History", "Test"
'                        .AddSubmenu xlSubmenu, 4, 6, "Performance Evaluation", , 140
'                                        .AddItem xlSubmenu, 6, "Pending Evaluations", "Test"
'                                        .AddItem xlSubmenu, 6, "Completed Evaluations", "Test"
'                                        .AddItem xlSubmenu, 6, "Performance Reports", "Test"
'                        .AddItem xlSubmenu, 4, "Feedbacks", "Test"
'                .AddSubmenu xlMenu, 1, 7, "Competence Management", , 200
'                        .AddItem xlSubmenu, 7, "Map Competencies", "Test"
'                        .AddItem xlSubmenu, 7, "View Competencies", "Test"
'                        .AddItem xlSubmenu, 7, "Create Individual Development Plan", "Test"
'
'        '============== Main Menu 2: Personnel Administration ================
'        .AddMenu 2, "Personnel Administration", 140, 160
'                .AddSubmenu xlMenu, 2, 8, "Payroll", , 150
'                        .AddItem xlSubmenu, 8, "Generate Payroll", "Test"
'                        .AddItem xlSubmenu, 8, "View Payments", "Test"
'                        .AddItem xlSubmenu, 8, "Issue Pay Slips", "Test"
'                        .AddItem xlSubmenu, 8, "Process 13th Salary", "Test"
'                        .AddItem xlSubmenu, 8, "Send Accounting Reports", "Test"
'                .AddSubmenu xlMenu, 2, 9, "Benefits", , 175
'                        .AddItem xlSubmenu, 9, "Manage Transportation Voucher", "Test"
'                        .AddItem xlSubmenu, 9, "Manage Meal Voucher", "Test"
'                        .AddItem xlSubmenu, 9, "Manage Health Plan", "Test"
'                        .AddItem xlSubmenu, 9, "View Benefits Reports", "Test"
'                .AddSubmenu xlMenu, 2, 10, "Vacations and Leaves", , 160
'                        .AddItem xlSubmenu, 10, "Vacation Planning", "Test"
'                        .AddItem xlSubmenu, 10, "View Vacation Balance", "Test"
'                        .AddItem xlSubmenu, 10, "Schedule Leave", "Test"
'                        .AddItem xlSubmenu, 10, "Vacation and Leave Reports", "Test"
'                .AddSubmenu xlMenu, 2, 11, "Electronic Timekeeping", , 170
'                        .AddItem xlSubmenu, 11, "Register Manual Adjustments", "Test"
'                        .AddItem xlSubmenu, 11, "View Timekeeping Reports", "Test"
'                        .AddItem xlSubmenu, 11, "Configure Work Shifts", "Test"
'
'        '============== Main Menu 3: Compliance and Legislation ================
'        .AddMenu 3, "Compliance and Legislation", 150, 165
'                .AddSubmenu xlMenu, 3, 12, "Norms and Regulations", , 150
'                        .AddItem xlSubmenu, 12, "View Internal Norms", "Test"
'                        .AddItem xlSubmenu, 12, "Update Conduct Manual", "Test"
'                        .AddItem xlSubmenu, 12, "Internal Communication", "Test"
'                .AddSubmenu xlMenu, 3, 13, "Internal Audits", , 150
'                        .AddItem xlSubmenu, 13, "Schedule Audit", "Test"
'                        .AddItem xlSubmenu, 13, "View Audit Reports", "Test"
'                        .AddItem xlSubmenu, 13, "Non-Conformance History", "Test"
'                .AddSubmenu xlMenu, 3, 14, "LGPD", , 155
'                        .AddItem xlSubmenu, 14, "View Privacy Policies", "Test"
'                        .AddItem xlSubmenu, 14, "Manage Consent Terms", "Test"
'                        .AddItem xlSubmenu, 14, "Adequacy Reports", "Test"
'                .AddSubmenu xlMenu, 3, 15, "CIPA", , 125
'                        .AddItem xlSubmenu, 15, "Elections", "Test"
'                        .AddItem xlSubmenu, 15, "Trainings", "Test"
'                        .AddItem xlSubmenu, 15, "Meeting Minutes", "Test"
'                        .AddItem xlSubmenu, 15, "Work Safety Reports", "Test"
'
'        '============== Main Menu 4: Reports ================
'        .AddMenu 4, "Reports", 50, 130
'                .AddSubmenu xlMenu, 4, 16, "HR Indicators", , 130
'                        .AddItem xlSubmenu, 16, "Turnover", "Test"
'                        .AddItem xlSubmenu, 16, "Absenteeism", "Test"
'                        .AddItem xlSubmenu, 16, "Average Hiring Time", "Test"
'                        .AddItem xlSubmenu, 16, "Training Costs", "Test"
'                .AddSubmenu xlMenu, 4, 17, "Financial Reports", , 120
'                        .AddItem xlSubmenu, 17, "Benefit Expenses", "Test"
'                        .AddItem xlSubmenu, 17, "Payroll Expenses", "Test"
'                        .AddItem xlSubmenu, 17, "Training Budget", "Test"
'
'        '============== Main Menu 5: Settings ================
'        .AddMenu 5, "Settings", 80, 150
'                .AddSubmenu xlMenu, 5, 18, "Users and Permissions", , 140
'                        .AddItem xlSubmenu, 18, "Create New User", "Test"
'                        .AddItem xlSubmenu, 18, "Manage Access Profiles", "Test"
'                        .AddItem xlSubmenu, 18, "Reset Passwords", "Test"
'                .AddSubmenu xlMenu, 5, 19, "Customization", , 155
'                        .AddItem xlSubmenu, 19, "Configure System Layout", "Test"
'                        .AddItem xlSubmenu, 19, "Adjust Form Fields", "Test"
'                        .AddItem xlSubmenu, 19, "Notifications", "Test"
'                .AddSubmenu xlMenu, 5, 20, "System", , 125
'                        .AddItem xlSubmenu, 20, "Data Backup", "Test"
'                        .AddItem xlSubmenu, 20, "Restore Backup", "Test"
'                        .AddItem xlSubmenu, 20, "View System Logs", "Test"
'                        .AddItem xlSubmenu, 20, "Updates", "Test"
'
'    End With
'
'End Sub
'
'Private Sub UserForm_Terminate()
'    Set Menu = Nothing
'End Sub


'======================================================================================================
'=====================================| Example code (pt-br) |=========================================
'======================================================================================================

'Dim Menu As New clsMenuStrip
'
'
'Private Sub UserForm_Initialize()
'    Call CriarMenu
'End Sub
'
'Sub CriarMenu()
'
'    With Menu
'
'        Set .UserForm = Me
'
'        '============== Menu Principal 1: Gestão de Pessoas ================
'        .AddMenu 1, "Gestão de Pessoas", 100, 200
'                .AddSubmenu xlMenu, 1, 1, "Recrutamento e Seleção", , 150
'                        .AddItem xlSubmenu, 1, "Cadastrar Vaga", "Teste", True
'                        .AddSubmenu xlSubmenu, 1, 2, "Gerenciar Vagas", , 135
'                                .AddItem xlSubmenu, 2, "Vagas Abertas", "Teste"
'                                .AddItem xlSubmenu, 2, "Vagas Encerradas", "Teste"
'                                .AddItem xlSubmenu, 2, "Histórico de Seleções", "Teste"
'                        .AddSubmenu xlSubmenu, 1, 3, "Banco de Currículos", True, 150
'                                .AddItem xlSubmenu, 3, "Consultar Currículos", "Teste"
'                                .AddItem xlSubmenu, 3, "Adicionar Novo Currículo", "Teste"
'                                .AddItem xlSubmenu, 3, "Exportar Banco de Dados", "Teste"
'                        .AddItem xlSubmenu, 1, "Agendar Entrevistas", "Teste"
'                .AddSubmenu xlMenu, 1, 4, "Treinamento e Desenvolvimento", , 190
'                        .AddItem xlSubmenu, 4, "Criar Programa de Treinamento", "Teste"
'                        .AddSubmenu xlSubmenu, 4, 5, "Gerenciar Treinamentos", , 160
'                                        .AddItem xlSubmenu, 5, "Treinamentos Ativos", "Teste"
'                                        .AddItem xlSubmenu, 5, "Treinamentos Concluídos", "Teste"
'                                        .AddItem xlSubmenu, 5, "Histórico de Participações", "Teste"
'                        .AddSubmenu xlSubmenu, 4, 6, "Avaliação de Desempenho", , 165
'                                        .AddItem xlSubmenu, 6, "Avaliações Pendentes", "Teste"
'                                        .AddItem xlSubmenu, 6, "Avaliações Realizadas", "Teste"
'                                        .AddItem xlSubmenu, 6, "Relatórios de Desempenho", "Teste"
'                        .AddItem xlSubmenu, 4, "Feedbacks", "Teste"
'                .AddSubmenu xlMenu, 1, 7, "Gestão de Competências", , 235
'                        .AddItem xlSubmenu, 7, "Mapear Competências", "Teste"
'                        .AddItem xlSubmenu, 7, "Consultar Competências", "Teste"
'                        .AddItem xlSubmenu, 7, "Criar Plano de Desenvolvimento Individual", "Teste"
'
'        '============== Menu Principal 2: Administração de Pessoal ================
'        .AddMenu 2, "Administração de Pessoal", 140, 140
'                .AddSubmenu xlMenu, 2, 8, "Folha de Pagamento", , 160
'                        .AddItem xlSubmenu, 8, "Gerar Folha", "Teste"
'                        .AddItem xlSubmenu, 8, "Consultar Pagamentos", "Teste"
'                        .AddItem xlSubmenu, 8, "Emitir Contra-Cheques", "Teste"
'                        .AddItem xlSubmenu, 8, "Processar 13º Salário", "Teste"
'                        .AddItem xlSubmenu, 8, "Enviar Relatórios Contábeis", "Teste"
'                .AddSubmenu xlMenu, 2, 9, "Benefícios", , 190
'                        .AddItem xlSubmenu, 9, "Gerenciar Vale-Transporte", "Teste"
'                        .AddItem xlSubmenu, 9, "Gerenciar Vale-Refeição", "Teste"
'                        .AddItem xlSubmenu, 9, "Gerenciar Plano de Saúde", "Teste"
'                        .AddItem xlSubmenu, 9, "Consultar Relatórios de Benefícios", "Teste"
'                .AddSubmenu xlMenu, 2, 10, "Férias e Licenças", , 170
'                        .AddItem xlSubmenu, 10, "Planejamento de Férias", "Teste"
'                        .AddItem xlSubmenu, 10, "Consultar Saldo de Férias", "Teste"
'                        .AddItem xlSubmenu, 10, "Agendar Licença", "Teste"
'                        .AddItem xlSubmenu, 10, "Relatórios de Férias e Licenças", "Teste"
'                .AddSubmenu xlMenu, 2, 11, "Ponto Eletrônico", , 180
'                        .AddItem xlSubmenu, 11, "Lançar Ajustes Manuais", "Teste"
'                        .AddItem xlSubmenu, 11, "Consultar Relatórios de Ponto", "Teste"
'                        .AddItem xlSubmenu, 11, "Configurar Escalas de Trabalho", "Teste"
'
'        '============== Menu Principal 3: Compliance e Legislação ================
'        .AddMenu 3, "Compliance e Legislação", 130, 165
'                .AddSubmenu xlMenu, 3, 12, "Normas e Regulamentos", , 170
'                        .AddItem xlSubmenu, 12, "Consultar Normas Internas", "Teste"
'                        .AddItem xlSubmenu, 12, "Atualizar Manual de Conduta", "Teste"
'                        .AddItem xlSubmenu, 12, "Comunicação Interna", "Teste"
'                .AddSubmenu xlMenu, 3, 13, "Auditorias Internas", , 190
'                        .AddItem xlSubmenu, 13, "Agendar Auditoria", "Teste"
'                        .AddItem xlSubmenu, 13, "Consultar Relatórios de Auditoria", "Teste"
'                        .AddItem xlSubmenu, 13, "Histórico de Não Conformidades", "Teste"
'                .AddSubmenu xlMenu, 3, 14, "LGPD", , 200
'                        .AddItem xlSubmenu, 14, "Consultar Políticas de Privacidade", "Teste"
'                        .AddItem xlSubmenu, 14, "Gerenciar Termos de Consentimento", "Teste"
'                        .AddItem xlSubmenu, 14, "Relatórios de Adequação", "Teste"
'                .AddSubmenu xlMenu, 3, 15, "CIPA", , 200
'                        .AddItem xlSubmenu, 15, "Eleições", "Teste"
'                        .AddItem xlSubmenu, 15, "Treinamentos", "Teste"
'                        .AddItem xlSubmenu, 15, "Atas de Reuniões", "Teste"
'                        .AddItem xlSubmenu, 15, "Relatórios de Segurança do Trabalho", "Teste"
'
'        '============== Menu Principal 4: Relatórios ================
'        .AddMenu 4, "Relatórios", 60, 150
'                .AddSubmenu xlMenu, 4, 16, "Indicadores de RH", , 170
'                        .AddItem xlSubmenu, 16, "Turnover", "Teste"
'                        .AddItem xlSubmenu, 16, "Absenteísmo", "Teste"
'                        .AddItem xlSubmenu, 16, "Tempo Médio de Contratação", "Teste"
'                        .AddItem xlSubmenu, 16, "Custos de Treinamento", "Teste"
'                .AddSubmenu xlMenu, 4, 17, "Relatórios Financeiros", , 250
'                        .AddItem xlSubmenu, 17, "Gastos com Benefícios", "Teste"
'                        .AddItem xlSubmenu, 17, "Gastos com Folha de Pagamento", "Teste"
'                        .AddItem xlSubmenu, 17, "Orçamento de Treinamento e Desenvolvimento", "Teste"
'
'        '============== Menu Principal 5: Configurações ================
'        .AddMenu 5, "Configurações", 80, 150
'                .AddSubmenu xlMenu, 5, 18, "Usuários e Permissões", , 155
'                        .AddItem xlSubmenu, 18, "Criar Novo Usuário", "Teste"
'                        .AddItem xlSubmenu, 18, "Gerenciar Perfis de Acesso", "Teste"
'                        .AddItem xlSubmenu, 18, "Resetar Senhas", "Teste"
'                .AddSubmenu xlMenu, 5, 19, "Personalização", , 170
'                        .AddItem xlSubmenu, 19, "Configurar Layout do Sistema", "Teste"
'                        .AddItem xlSubmenu, 19, "Ajustar Campos de Formulário", "Teste"
'                        .AddItem xlSubmenu, 19, "Notificações", "Teste"
'                .AddSubmenu xlMenu, 5, 20, "Sistema", , 150
'                        .AddItem xlSubmenu, 20, "Backup de Dados", "Teste"
'                        .AddItem xlSubmenu, 20, "Restaurar Backup", "Teste"
'                        .AddItem xlSubmenu, 20, "Consultar Logs de Sistema", "Teste"
'                        .AddItem xlSubmenu, 20, "Atualizações", "Teste"
'
'    End With
'
'End Sub
'
'
'
'Private Sub UserForm_Terminate()
'    Set Menu = Nothing
'End Sub



'======================================================================================================
'======================================================================================================

Public WithEvents xlMenuStrip As MSForms.Label
Attribute xlMenuStrip.VB_VarHelpID = -1
Public WithEvents xlSubmenus As MSForms.Label
Attribute xlSubmenus.VB_VarHelpID = -1
Public WithEvents xlItems As MSForms.Label
Attribute xlItems.VB_VarHelpID = -1
Public WithEvents xlClose As MSForms.Label
Attribute xlClose.VB_VarHelpID = -1
Public WithEvents xlImage As MSForms.Label
Attribute xlImage.VB_VarHelpID = -1
Public WithEvents xlArrow As MSForms.Label
Attribute xlArrow.VB_VarHelpID = -1
Public WithEvents xlListBox As MSForms.ListBox
Attribute xlListBox.VB_VarHelpID = -1
Public WithEvents xlFrame As MSForms.frame
Attribute xlFrame.VB_VarHelpID = -1
Public WithEvents xlFrameMenu As MSForms.frame
Attribute xlFrameMenu.VB_VarHelpID = -1
Public WithEvents xlMultiPage As MSForms.MultiPage
Attribute xlMultiPage.VB_VarHelpID = -1
Public WithEvents xlSeparator As MSForms.Label
Attribute xlSeparator.VB_VarHelpID = -1


Private form As Object
Private positionLeft As Single
Private colMenus As New Collection
Private colFrameMenu As New Collection
Private colFrameSubmenu As New Collection
Private colLabelNames As New Collection
Private colItemNames As New Collection



Enum Tipo
xlMenu
xlSubmenu
End Enum

Enum xlTemas
xlDracula
xlDefault
xlDarkBlue
xlDarkOrange
xlDarkGreen1
xlDarkGreen2
xlDarkYellow
xlDarkPurple
xlDarkWhite
xlDarkGray
xlDarkPink
xlDarkRed
xlWhiteBlue
xlWhiteOrange
xlWhiteGreen1
xlWhiteGreen2
xlWhiteYellow
xlWhitePurple
xlWhitePink
xlWhiteRed
xlEdge
xlExcel
xlWord
xlPowerPoint
xlAccess
xlOneNote
xlOutlook
xlSpotify
xlVisualStudio1
xlVisualStudio2
xlBlocoNotas
xlVBA1
xlVBA2
xlVBA3
End Enum

'====================================================================================================================
'===========================================| Propriedades |=========================================================
'====================================================================================================================


'Propriedade para criar a faixa do menu strip
Property Set UserForm(vForm As Object)
    
    Dim lbFechar As MSForms.Label
    Dim cls As clsMenuStrip
    
    'Define a variável publica para ser o control
    Set form = vForm
    
    'Manipular eventos dos controles que o controle label
    'responsável por fechar não consegue sobrepor
    Call ManipularEventos
    
    'Controle que representa a barra do menu strip
    With CreateLabel(form, "lb_FaixaMenu")
        .Width = form.Width
        .Height = 16
        .BackColor = vbWhite
        .Font.Size = 11
        .Font.Name = "Calibri"
        .Forecolor = vbBlack
    End With
    
    'Label que armazenará o design de seleção dos controles
    With CreateLabel(form, "lb_MenuCustom", False)
        .Forecolor = vbBlack
        .BackColor = RGB(240, 240, 240)
        .BorderColor = RGB(200, 200, 200)
        .Tag = RGB(200, 200, 200)
    End With
    
    
    'Label responsável para remover todos os efeitos
    Set lbFechar = form.Controls.Add("Forms.Label.1", "lb_Close", False)
    
    With lbFechar
        .Width = form.Width
        .Height = form.Height
        .BackStyle = fmBackStyleTransparent

        Set cls = New clsMenuStrip
        Set cls.xlClose = lbFechar
        
        colMenus.Add cls
        
    End With
    
    
    'Label para armazenar os nomes dos menus ativos.
    'Precisei fazer isso para corrigir o bug de dois submenus
    'ficarem aberto na mesma direção, para os casos quando tem
    'submenus consecutivos e dispara o evento do mouse move
    Call CreateLabel(form, "lb_MenusExibidos", False)
    
    
    'Label que guardará os nomes dos submenus
    Call CreateLabel(form, "lb_SubmenuNames", False)
    
End Property


Private Sub ManipularEventos()
    
    Dim ctrl As Control
    Dim cls As clsMenuStrip
    
    'Label que guardará os nomes das listbox para corrigir o bug
    Call CreateLabel(form, "lb_ListBox", False)
    
    For Each ctrl In form.Controls
        
        If TypeOf ctrl Is MSForms.ListBox Then
        
            Set cls = New clsMenuStrip
            Set cls.xlListBox = ctrl
            colMenus.Add cls
            
            'Salva os nomes das listbox para corrigir o bug
            If form.Controls("lb_ListBox").Caption = "" Then
                form.Controls("lb_ListBox").Caption = ctrl.Name
            Else
               form.Controls("lb_ListBox").Caption = _
               form.Controls("lb_ListBox").Caption & "|" & ctrl.Name
            End If
            
        ElseIf TypeOf ctrl Is MSForms.MultiPage Then
        
            Set cls = New clsMenuStrip
            Set cls.xlMultiPage = ctrl
            colMenus.Add cls
            
        ElseIf TypeOf ctrl Is MSForms.frame Then
        
            Set cls = New clsMenuStrip
            Set cls.xlFrame = ctrl
            colMenus.Add cls
            
        End If
        
    Next
    
End Sub


Property Let CustomBackColor(Color As Single)
    'Verificar se o userform foi instanciado
    Call CheckForm
    form.Controls("lb_FaixaMenu").BackColor = Color
End Property

Property Get CustomBackColor() As Single
    'Verificar se o userform foi instanciado
    Call CheckForm
    CustomBackColor = form.Controls("lb_FaixaMenu").BackColor
End Property

Property Let CustomFontSize(Size As Single)
    'Verificar se o userform foi instanciado
    Call CheckForm
    form.Controls("lb_FaixaMenu").Font.Size = Size
End Property

Property Get CustomFontSize() As Single
    'Verificar se o userform foi instanciado
    Call CheckForm
    CustomFontSize = form.Controls("lb_FaixaMenu").Font.Size
End Property

Property Let CustomFontName(Name As String)
    'Verificar se o userform foi instanciado
    Call CheckForm
    form.Controls("lb_FaixaMenu").Font.Name = Name
End Property

Property Get CustomFontName() As String
    'Verificar se o userform foi instanciado
    Call CheckForm
    CustomFontName = form.Controls("lb_FaixaMenu").Font.Name
End Property

Private Property Get CustomHeight() As Single
    CustomHeight = form.Controls("lb_FaixaMenu").Height
End Property

Property Let CustomForeColor(Color As Single)
    'Verificar se o userform foi instanciado
    Call CheckForm
    form.Controls("lb_FaixaMenu").Forecolor = Color
End Property

Property Get CustomForeColor() As Single
    'Verificar se o userform foi instanciado
    Call CheckForm
    CustomForeColor = form.Controls("lb_FaixaMenu").Forecolor
End Property

Property Let CustomSelectionForeColor(Color As Single)
    'Verificar se o userform foi instanciado
    Call CheckForm
    form.Controls("lb_MenuCustom").Forecolor = Color
End Property

Property Get CustomSelectionForeColor() As Single
    'Verificar se o userform foi instanciado
    Call CheckForm
    CustomSelectionForeColor = form.Controls("lb_MenuCustom").Forecolor
End Property

Property Let CustomSelectionBackColor(Color As Single)
    'Verificar se o userform foi instanciado
    Call CheckForm
    form.Controls("lb_MenuCustom").BackColor = Color
End Property

Property Get CustomSelectionBackColor() As Single
    'Verificar se o userform foi instanciado
    Call CheckForm
    CustomSelectionBackColor = form.Controls("lb_MenuCustom").BackColor
End Property

Property Let CustomBorderColor(Color As Single)
    'Verificar se o userform foi instanciado
    Call CheckForm
    form.Controls("lb_MenuCustom").BorderColor = Color
End Property

Property Get CustomBorderColor() As Single
    'Verificar se o userform foi instanciado
    Call CheckForm
    CustomBorderColor = form.Controls("lb_MenuCustom").BorderColor
End Property

Property Let CustomSeparatorColor(Color As Single)
    'Verificar se o userform foi instanciado
    Call CheckForm
    form.Controls("lb_MenuCustom").Tag = Color
End Property

Property Get CustomSeparatorColor() As Single
    'Verificar se o userform foi instanciado
    Call CheckForm
    CustomSeparatorColor = form.Controls("lb_MenuCustom").Tag
End Property


Property Let Themes(Tema As xlTemas)

    'Verificar se o userform foi instanciado
    Call CheckForm
    
    Select Case Tema
    '-----------------------------------| Temas Darks |-------------------------------
        Case xlDarkBlue
            Call ApplyThemes(3549736, 16640395, 3549736, 16640395, 3549736, 10777186)
            
        Case xlDarkOrange
            Call ApplyThemes(3549736, 1411582, 3549736, 1411582, 3549736, 10777186)
            
        Case xlDarkYellow
            Call ApplyThemes(3549736, 117759, 3549736, 117759, 3549736, 10777186)
            
        Case xlDarkGreen1
            Call ApplyThemes(3549736, 5490590, 3549736, 5490590, 3549736, 10777186)
            
        Case xlDarkGreen2
            Call ApplyThemes(3549736, 9364099, 3549736, 9364099, 3549736, 10777186)
            
        Case xlDarkPurple
            Call ApplyThemes(3549736, 16356285, 3549736, 16356285, 3549736, 10777186)
            
        Case xlDarkWhite
            Call ApplyThemes(3549736, 16777215, 3549736, 16777215, 3549736, 10777186)

        Case xlDarkPink
            Call ApplyThemes(3549736, 12940783, 3549736, 12940783, 3549736, 12940783)
        
        Case xlDarkRed
            Call ApplyThemes(3549736, 5592575, 3549736, 5592575, 3549736, 10777186)
            
    '------------------------------- | Temas Brancos | ----------------------------------
        Case xlWhiteBlue
            Call ApplyThemes(16777215, 10376717, 14559767, 10376717, 16777215, 10376717)
            
        Case xlWhiteOrange
            Call ApplyThemes(16777215, 1411582, 1411582, 1411582, 16777215, 1411582)
            
        Case xlWhiteYellow
            Call ApplyThemes(16777215, 37568, 37568, 37568, 16777215, 37568)
            
        Case xlWhiteGreen1
            Call ApplyThemes(16777215, 5490590, 5490590, 5490590, 16777215, 5490590)
            
        Case xlWhiteGreen2
            Call ApplyThemes(16777215, 5287936, 5287936, 5287936, 16777215, 5287936)
            
        Case xlWhitePurple
            Call ApplyThemes(16777215, 13978799, 13978799, 13978799, 16777215, 13978799)

        Case xlWhitePink
            Call ApplyThemes(16777215, 12540643, 12540643, 12540643, 16777215, 12540643)
        
        Case xlWhiteRed
            Call ApplyThemes(16777215, 5592575, 5592575, 5592575, 16777215, 5592575)
            
   '----------------------------------| Temas Diversos |-----------------------------------
        Case xlDracula
            Call ApplyThemes(3549736, 16356285, 10777186, 15923448, 0, 10777186)
            
        Case xlExcel
            Call ApplyThemes(3628056, 6726433, 4291600, 16777215, 16777215, 4291600)
            
        Case xlEdge
            Call ApplyThemes(9325584, 9754960, 9754960, 16777215, 16777215, 14387982)
            
        Case xlAccess
            Call ApplyThemes(2167944, 8021716, 6311881, 16777215, 16777215, 6311881)
            
        Case xlPowerPoint
            Call ApplyThemes(1851337, 4680942, 7049215, 16777215, 16777215, 7049215)
            
        Case xlWord
            Call ApplyThemes(9518864, 13859883, 15639873, 16777215, 16777215, 15639873)
            
        Case xlSpotify
            Call ApplyThemes(6346526, 985618, 985618, 1117970, 6346526, 985618)
            
        Case xlOutlook
            Call ApplyThemes(16767313, 13924096, 11758086, 9782784, 16777215, 11758086)
            
        Case xlOneNote
            Call ApplyThemes(11278456, 13978542, 15361226, 16777215, 16777215, 15361226)
        
        Case xlVisualStudio1
            Call ApplyThemes(14777772, 9257820, 16291531, 16777215, 16777215, 16291531)
    
        Case xlVisualStudio2
            Call ApplyThemes(9257820, 14777772, 16291531, 16777215, 16777215, 16291531)
            
        Case xlBlocoNotas
            Call ApplyThemes(9991434, 14343132, 14005837, 16777215, 9991434, 14866877)

        Case xlVBA1
            Call ApplyThemes(12748649, 8728216, 0, 16777215, 16777215, 0)
    
        Case xlVBA2
            Call ApplyThemes(12748649, 3571673, 0, 16777215, 16777215, 0)
        
        Case xlVBA3
            Call ApplyThemes(12748649, 579041, 0, 16777215, 0, 0)
            
        Case xlDefault
            Call ApplyThemes(16777215, 15790320, 13158600, 0, 0, 13158600)
            
        Case Else
            Call ApplyThemes(16777215, 15790320, 13158600, 0, 0, 13158600)
            
    End Select
    
End Property


Private Sub ApplyThemes(BackColor As Long, BackColorSelecion As Long, BoderColor As Long, _
                    Forecolor As Long, ForeColorSelection As Long, SeparatorColor As String)
                    
    CustomBackColor = BackColor
    CustomSelectionBackColor = BackColorSelecion
    CustomBorderColor = BoderColor
    CustomForeColor = Forecolor
    CustomSelectionForeColor = ForeColorSelection
    CustomSeparatorColor = SeparatorColor
    
End Sub


'====================================================================================================================
'==============================================| Métodos |===========================================================
'====================================================================================================================


Private Sub AddClassMenuItems(ctrl As Control)

    Dim cls As clsMenuStrip
    Set cls = New clsMenuStrip
    
    Set cls.xlItems = ctrl
    
    colMenus.Add cls
    
End Sub

Private Sub AddClassSubmenu(ctrl As Control)

    Dim cls As clsMenuStrip
    Set cls = New clsMenuStrip
    
    Set cls.xlSubmenus = ctrl
    
    colMenus.Add cls
    
End Sub

Private Sub AddClassImage(ctrl As Control)

    Dim cls As clsMenuStrip
    Set cls = New clsMenuStrip
    
    Set cls.xlImage = ctrl
    
    colMenus.Add cls
    
End Sub

Private Sub AddClassArrow(ctrl As Control)

    Dim cls As clsMenuStrip
    Set cls = New clsMenuStrip
    
    Set cls.xlArrow = ctrl
    
    colMenus.Add cls
    
End Sub

Private Sub AddClassMenuStrip(ctrl As Control)

    Dim cls As clsMenuStrip
    Set cls = New clsMenuStrip
    
    Set cls.xlMenuStrip = ctrl
    
    On Error GoTo Sair
    colMenus.Add cls, ctrl.Name
    
    Exit Sub
    
Sair:
    
    MsgBox "Duplicate menu names are not allowed!" & _
    vbLf & "----------------------------------------------------" & _
    vbLf & "Duplicate name : " & ctrl.Name, 16, "  Menu Strip"
    
    End
    
End Sub


Private Function CreateLabel(Parent As Object, Optional Name As String, Optional Visible As Boolean = True) As MSForms.Label
    
    Dim lb As MSForms.Label
    
    If Name = "" Then
        Set lb = Parent.Controls.Add("Forms.Label.1", , Visible)
    Else
        Set lb = Parent.Controls.Add("Forms.Label.1", Name, Visible)
    End If
    
    'Se a label for ficar visível será configurada
    If Visible Then
    
        'Algumas definições padrões para todas as labels
        With lb
            .Font.Size = CustomFontSize
            .Font.Name = CustomFontName
            .BackColor = CustomBackColor
            .Forecolor = CustomForeColor
            .Height = CustomHeight
        End With
        
    End If
    
    'Devolve a label pré-configurada
    Set CreateLabel = lb
    
End Function

Private Function CreateFrameMenu(Id As String, Parent As Object, Optional Name As String, _
                                 Optional Visible As Boolean = True) As MSForms.frame
    
    Dim frameMenu As MSForms.frame
    Dim cls As clsMenuStrip
    
    If Name = "" Then
        Set frameMenu = Parent.Controls.Add("Forms.Frame.1", , Visible)
    Else
        Set frameMenu = Parent.Controls.Add("Forms.Frame.1", "fr_" & Name, Visible)
    End If
    
    'Algumas definições padrões para todos os frames
    With frameMenu
        .BackColor = CustomBackColor
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CustomBorderColor
        
        Set cls = New clsMenuStrip
        Set cls.xlFrameMenu = frameMenu
        
        colMenus.Add cls
        
    End With

    
    On Error GoTo Sair
    colFrameMenu.Add frameMenu, Id
        
    'Devolve o frame pré-configurado
    Set CreateFrameMenu = frameMenu
    
    
    Exit Function
    
Sair:
        
    Call IdDupicate(Id, Name, "'menu'")

End Function

Private Function CreateFrameSubmenu(Id As String, Parent As Object, Optional Name As String, _
                                    Optional Visible As Boolean = True) As MSForms.frame
                                
    
    Dim frameSubmenu As MSForms.frame
    Dim cls As clsMenuStrip
    
    
    If Name = "" Then
        Set frameSubmenu = Parent.Controls.Add("Forms.Frame.1", , Visible)
    Else
        Set frameSubmenu = Parent.Controls.Add("Forms.Frame.1", "sub_" & Name, Visible)
    End If
    
    'Algumas definições padrões para todos os frames
    With frameSubmenu
        .BackColor = CustomBackColor
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = CustomBorderColor
        
        Set cls = New clsMenuStrip
        Set cls.xlFrameMenu = frameSubmenu
        
        colMenus.Add cls
        
    End With

    
    On Error GoTo Sair
    colFrameSubmenu.Add frameSubmenu, Id
        
    'Devolve o frame pré-configurado
    Set CreateFrameSubmenu = frameSubmenu
    
    
    Exit Function
    
Sair:
    
    Call IdDupicate(Id, Name, "'submenu'")
    
End Function

Private Sub IdDupicate(Id As String, Name As String, Parent As String)

    MsgBox "Duplicate ID in the " & Parent & " are not allowed!" _
    & vbLf & "---------------------------------------------------------" & _
    vbLf & "Id: " & Id & vbLf & "Label: " & Name, 16, "  Menu Strip"
    
    End
    
End Sub


Private Function CreateImage(Parent As Object, Optional Name As String, _
                             Optional Visible As Boolean = True) As MSForms.Label
                                    
    
    Dim xlImage As MSForms.Label
    
    If Name = "" Then
        Set xlImage = Parent.Controls.Add("Forms.Label.1", , Visible)
    Else
        Set xlImage = Parent.Controls.Add("Forms.Label.1", "img_" & Name, Visible)
    End If
    
    'Algumas definições padrões para todos os frames
    With xlImage
        .BackColor = CustomBackColor
        .BackStyle = 0
        .Height = CustomHeight
        .Width = 20
        .PicturePosition = fmPicturePositionLeftCenter
    End With
    
    
    'Adiciona a label a coleção e a classe
    Call AddClassImage(xlImage)
        
    'Devolve o label pré-configurada
    Set CreateImage = xlImage
    
End Function


Sub AddMenu(Id As String, Label As String, Optional Width As Single = 60, Optional WithItems As Single = 80)
    
    Dim lbMenuPrincipal As MSForms.Label
    
    'Verificar se o userform foi instanciado
    Call CheckForm
    
    'Criação das labels que representará a navegação do menu principal
    Set lbMenuPrincipal = CreateLabel(form, Label)
    
    With lbMenuPrincipal
        .Caption = Label
        .TextAlign = 2
        .Left = positionLeft
        .Width = Width
        
        'Variael incremental para posicionar o menu corretamente
        positionLeft = positionLeft + .Width
        
        'Adiciona a label a classe
        Call AddClassMenuStrip(lbMenuPrincipal)
        
    End With
    
    
    'Criação do frame do menu imediato
    With CreateFrameMenu(Id, form, Label, False)
        .Width = WithItems
        .Height = CustomHeight + 1
        .Top = CustomHeight
    End With
    
End Sub

Sub AddSubmenu(Parent As Tipo, IdParent As String, Id As String, Label As String, _
               Optional Separator As Boolean, Optional WithItems As Single = 80, _
               Optional Image As MSForms.Image)
              
    
    Dim lbSubmenu As MSForms.Label
    Dim frameMenu As MSForms.frame
    Dim xlArrow As MSForms.Label
    Dim frameSubmenu As MSForms.Label
    Dim xLabel As String, ultCtrl As Control
    
    'Verificar se o userform foi instanciado
    Call CheckForm
    
    
    On Error GoTo Sair
    
    If Parent = xlMenu Then
        Set frameMenu = colFrameMenu(IdParent)
    Else
        Set frameMenu = colFrameSubmenu(IdParent)
    End If
    
    On Error GoTo 0
    
    'Tratar itens com nomes duplicados
    xLabel = DuplicateLabel(Label)
    
    
    'Retorna o último controle adicionado no frame
    If frameMenu.Controls.Count > 0 Then
        Set ultCtrl = frameMenu.Controls(frameMenu.Controls.Count - 1)
    End If
    
    
    'Criação das labels que representará cada item do submenu
    Set lbSubmenu = CreateLabel(frameMenu, xLabel)
    
    With lbSubmenu
        .Caption = Label
        .Left = 20
        .Width = frameMenu.Width
        .Font.Size = CustomFontSize
        .ZOrder 1
        
        If Not ultCtrl Is Nothing Then
            If ultCtrl.Name = "lbSep" Then
                .Top = ultCtrl.Top + ultCtrl.Height * 2
            Else
                .Top = ultCtrl.Top + ultCtrl.Height
            End If

        End If
        
        'Adiciona a label a classe
        Call AddClassSubmenu(lbSubmenu)
        
    End With
    
    
    'Criação das imagens
    With CreateImage(frameMenu, xLabel)
        .Top = lbSubmenu.Top
        
        'Caso o usuário informe a imagem a mesma será passada
        If Not Image Is Nothing Then
            .Picture = Image.Picture
        End If
        
    End With
    
    'Criação das labels que representará cada item do submenu
    Set xlArrow = CreateLabel(frameMenu, "arrow_" & xLabel)
    
    With xlArrow
        .Width = 15
        .Top = lbSubmenu.Top
        .Font.Size = 10
        .Font.Name = "webdings"
        .Caption = 4
        .Left = frameMenu.Width - .Width
        .BackStyle = fmBackStyleTransparent
        .Tag = "sub_" & xLabel

        
        Call AddClassArrow(xlArrow)
        
    End With
    
    
    'Criação do frame que representará o submenu
    With CreateFrameSubmenu(Id, form, xLabel, False)
        .Width = WithItems
        .Height = CustomHeight + 1
        .Top = CustomHeight
    End With
    
    
    'Caso o usuário escolha o separdor, o mesmo será adicionado
    If Separator Then
        Call AddSeparatorMenu(frameMenu, lbSubmenu)
    End If

    
    'Ajustar a altura do frame
    With frameMenu.Controls(frameMenu.Controls.Count - 1)
        frameMenu.Height = .Top + .Height + 1
    End With
    
  
    
    Exit Sub
    
Sair:

    Call IdError(Parent, IdParent, Label)
    
End Sub


Sub AddItem(Parent As Tipo, IdParent As String, Label As String, _
            MacroName As String, Optional Separator As Boolean, Optional Image As MSForms.Image, _
            Optional ItemName As String)
              
    
    Dim lbItem As MSForms.Label
    Dim frame As MSForms.frame
    Dim xLabel As String
    Dim ultCtrl As Object
    
    'Verificar se o userform foi instanciado
    Call CheckForm
    
    If MacroName = "" Then GoTo Parar
    
    On Error GoTo Sair
    
    If Parent = xlMenu Then
        Set frame = colFrameMenu(IdParent)
    Else
        Set frame = colFrameSubmenu(IdParent)
    End If
    
    On Error GoTo 0
    
    'Tratar itens com nomes duplicados
    xLabel = DuplicateLabel(Label)
    
    'Nome do controle para controlar se vai está habilitado ou não
    If ItemName <> "" Then
        If Left(ItemName, 1) = "#" Then
            
            On Error GoTo Cancelar
            colItemNames.Add ItemName, ItemName
            xLabel = ItemName
            
        Else
            MsgBox "The item's name needs to start with '#'.", 16, "  Menu Strip"
            End
        End If
    End If
    
    
    'Retorna o último controle adicionado no frame
    If frame.Controls.Count > 0 Then
        Set ultCtrl = frame.Controls(frame.Controls.Count - 1)
    End If
    
    
    'Criação das labels que representará cada item do submenu
    Set lbItem = CreateLabel(frame, xLabel)
    
    With lbItem
        .Caption = Label
        .Left = 20
        .Width = frame.Width
        .Font.Size = CustomFontSize
        .Tag = MacroName
        .ZOrder 1
        
        
        If Not ultCtrl Is Nothing Then
            If ultCtrl.Name = "lbSep" Then
                .Top = ultCtrl.Top + ultCtrl.Height * 2
            Else
               .Top = ultCtrl.Top + ultCtrl.Height
            End If
        End If
   
        
        
        'Adiciona a label a classe
        Call AddClassMenuItems(lbItem)
        
    End With
    
    
    'Criação das imagens
    With CreateImage(frame, xLabel)
        .Top = lbItem.Top
        
        'Caso o usuário informe a imagem a mesma será passada
        If Not Image Is Nothing Then
            .Picture = Image.Picture
        End If
        
    End With
    
    
    'Caso o usuário escolha o separdor, o mesmo será adicionado
    If Separator Then
        Call AddSeparatorMenu(frame, lbItem)
    End If

    
    'Ajustar a altura do frame
    With frame.Controls(frame.Controls.Count - 1)
        frame.Height = .Top + .Height + 1
    End With
    
    
    Exit Sub
    
Sair:
    
    Call IdError(Parent, IdParent, Label)
    
    
    Exit Sub
    
Parar:
    
        MsgBox "The macro name cannot be passed as empty!" & _
        vbLf & "---------------------------------------------------------" & _
        vbLf & "Id: " & IdParent & vbLf & "Label: " & Label, 16, "  Menu Strip"
        End
        
Exit Sub

Cancelar:

        MsgBox "The item cannot have a duplicate name!" & _
        vbLf & "------------------------------------------------" & _
        vbLf & "Item Name: " & ItemName, 16, "  Menu Strip"
        End
    
    
End Sub

Private Function AddSeparatorMenu(Parent As MSForms.frame, Item As MSForms.Label) As MSForms.Label
    
    Dim lbSeparator As MSForms.Label
    Dim cls As clsMenuStrip
    
    Set lbSeparator = Parent.Controls.Add("Forms.Label.1", "lbSep")
    
    With lbSeparator
        .Width = Parent.Width - 24
        .Left = 20
        .Height = 0.7
        .BackColor = CustomSeparatorColor
        .ZOrder 0
        .Top = Item.Top + Item.Height + .Height
        
        Set cls = New clsMenuStrip
        Set cls.xlSeparator = lbSeparator
        
        colMenus.Add cls
        
    End With

    
End Function

Private Sub IdError(Parent As Tipo, IdParent, Label)

    If Parent = xlMenu Then
        MsgBox "Menu ID not found!" & vbLf & "-----------------------------------" & _
        vbLf & "Id: " & IdParent & vbLf & "Label: " & Label, 48, "  Menu Strip"
    Else
        MsgBox "Submenu ID not found!" & vbLf & "-----------------------------------" _
        & vbLf & "Id: " & IdParent & vbLf & "Label: " & Label, 48, "  Menu Strip"
    End If
    
    End
    
End Sub


Private Sub RemoverEfeitoMenuStrip()
    
    'A remoção do efeito será aplicada no menu que teve o seu nome guardado na tag
    If form.Controls("lb_FaixaMenu").Tag <> "" Then
        With form.Controls(form.Controls("lb_FaixaMenu").Tag)
        
            .Forecolor = CustomForeColor
            .BackColor = CustomBackColor
            
            'Oculta o frame que representa o menu
            form.Controls("fr_" & form.Controls("lb_FaixaMenu").Tag).Visible = False
            
            form.Controls("lb_FaixaMenu").Tag = ""
            
        End With
    End If
    
End Sub

Private Sub OcultarSubmenus(Optional Submenu As MSForms.Label)
    
    Dim arrSubmenus, i As Integer

    
    With form.Controls("lb_SubmenuNames")
        'Verifica se tem algum valor na tag
        If .Tag <> "" Then
            'Caso tenha, o valor será transformado em array
            arrSubmenus = Split(.Tag, "|")
            
            'Um loop será realizado pelo array pegando o nome
            'de cada submenu que foi exibido e fará o processo inverso
            
            'Mas antes será verificado se foi passado o submenu,
            'pois essa rotina também é chamada no evento mouse move
            'da label xlClose e label do menu strip.
     
             
            If Submenu Is Nothing Then
                'Se não for passado será ocultado todos os menus
                For i = 0 To UBound(arrSubmenus)
                    form.Controls(arrSubmenus(i)).Visible = False
                Next
                'Para facilitar a reutilização
                'o valor tag voltará a ficar vazia
                .Tag = ""
            Else
                'Caso seja informado, não será ocultado o submenu
                'que está sob o cursor do mouse
                For i = UBound(arrSubmenus) To 0 Step -1
                    If "sub_" & Submenu.Name <> arrSubmenus(i) Then
                        form.Controls(arrSubmenus(i)).Visible = False
                    Else
                        Exit For
                    End If
                Next
                
            End If
            
        End If
        
    End With
        
End Sub

Private Sub OcultarSubmenu(Optional MenuItem As MSForms.Label, Optional Parent As MSForms.frame)
        
    Dim arrSubmenus, i As Integer
    

    'Cria o array com os nomes dos submenus foram exibidos
    arrSubmenus = Split(form.Controls("lb_SubmenuNames").Tag, "|")
    
    'Um loop será feito por cada nome do submenu que foi exibido
    For i = UBound(arrSubmenus) To 0 Step -1
        
        
        If Not MenuItem Is Nothing Then
        
            'Se o nome do submenu for diferente do mome do menu/submenu
            'que o item pertence, esses submenus serão ocultados
            If MenuItem.Parent.Name <> arrSubmenus(i) Then
                form.Controls(arrSubmenus(i)).Visible = False
            Else
                Exit For
            End If
            
        Else
    
            If Parent.Name <> arrSubmenus(i) Then
                form.Controls(arrSubmenus(i)).Visible = False
            Else
                Exit For
            End If
        
        End If
        
    Next
    
End Sub


Private Sub RemoverSelecao(Item As MSForms.Label)
    
    'Verifica se tem alguma informação na tag
    If Item.Parent.Tag <> "" Then
    
        With form.Controls(Item.Parent.Tag)
            'Caso tenha, será verificado se o nome salvo
            'é diferente do item atual, caso não seja
            'o controle em questão será descelecionado
            If .Tag <> Item.Name Then
                .Forecolor = CustomForeColor
                .BackColor = CustomBackColor
                
                'O mesmo acontecerá para as imagens
                With form.Controls("img_" & Item.Parent.Tag)
                    .BackColor = CustomBackColor
                End With
                
                'O erro é para tratar o caso quanto o item não é um submenu
                'pois o mesmo não possui a setinha
                On Error Resume Next
                'O mesmo acontecerá para as imagens
                With form.Controls("arrow_" & Item.Parent.Tag)
                    .Forecolor = CustomForeColor
                End With
                
            End If
            
        End With
    End If

End Sub


Private Function DuplicateLabel(ByVal Label As String) As String
    
    On Error GoTo Tratar
    
    
    colLabelNames.Add Label, Label
    
    DuplicateLabel = Label
    
    
    Exit Function
    
Tratar:
    
    Label = Label & "_"
    Resume

End Function


Private Sub EventItems_MouveMove(xlItems As Object)
    
    Dim frameMenu As MSForms.frame
    Dim arrSubmenus, i As Integer
    
    Set frameMenu = xlItems.Parent
    Set form = frameMenu.Parent
  

    'Aplica a seleção se o controle não estiver selecionado
    If xlItems.BackColor <> CustomSelectionBackColor Then
    
        'Remove o feito da seleção anterior
        Call RemoverSelecao(xlItems)
        
        'Salva o nome do item selecionado na tag do submenu
        frameMenu.Tag = xlItems.Name
    
            
            'Destaca os itens
            With xlItems
                'Define a variável que representa o formulãrio
                Set form = .Parent.Parent
        
                .BackColor = CustomSelectionBackColor
                .Forecolor = CustomSelectionForeColor
                
            End With
            
            
            'Destaca as imagens
            With frameMenu.Controls("img_" & xlItems.Name)
                .BackColor = CustomSelectionBackColor
                .BackStyle = 1
            End With
            
            'Destaca a setinha caso seja um submenu
            On Error Resume Next
                With frameMenu.Controls("arrow_" & xlItems.Name)
                    .Forecolor = CustomSelectionForeColor
                End With
            On Error GoTo 0
            
            'Chama a macro para ocultar o submenu que foi aberto anteriormente
            Call OcultarSubmenu(xlItems)
            
            'Se o item estiver em um submenu, o nome do submenu
            'será salvo, para que o mesmo seja fechado quando
            'o mouse estiver sobre outro submenu
            If Left(xlItems.Parent.Name, 4) = "sub_" Then
                form.Controls("lb_SubmenuNames").Caption = xlItems.Parent.Name
            End If
            
            'Cria o array com os nomes dos submenus foram exibidos
            arrSubmenus = Split(form.Controls("lb_SubmenuNames").Tag, "|")
            
            'Um loop será feito por cada nome do submenu que foi exibido
            For i = UBound(arrSubmenus) To 0 Step -1
                
                'Se o nome do submenu for diferente do mome do menu/submenu
                'que o item pertence, esses submenus serão ocultados
                If xlItems.Parent.Name <> arrSubmenus(i) Then
                    form.Controls(arrSubmenus(i)).Visible = False
                Else
                    Exit For
                End If
                
            Next
        
    End If

End Sub

Private Sub EventItems_Click(xlItems As Object)


    On Error GoTo Sair
   
    form.Controls("lb_Close").Visible = False
    
    Call RemoverEfeitoMenuStrip
    Call OcultarSubmenus
    
    Application.Run xlItems.Tag

   Exit Sub
   
Sair:
   
   MsgBox "The '" & xlItems.Tag & "' macro was not found!", 48, "  Menu Strip"
   
End Sub

Private Sub EventSubmenuItems_MouseMove(xlSubmenus As Object)
    
    Dim frameSubmenu As MSForms.frame
    Dim arrSubmenu, i As Byte
    Dim ctrl As MSForms.Control
    
    Set frameSubmenu = xlSubmenus.Parent
    Set form = frameSubmenu.Parent
             
      
    'Remove o feito da seleção anterior
    Call RemoverSelecao(xlSubmenus)

    
    'Salva o nome do item selecionado na tag do submenu
    frameSubmenu.Tag = xlSubmenus.Name
    
    'Aplica a seleção se o controle não estiver selecionado
    If xlSubmenus.BackColor <> CustomSelectionBackColor Then
            
            'Destaca os itens
            With xlSubmenus
                'Define a variável que representa o formulãrio
                Set form = .Parent.Parent
                
                .BackColor = CustomSelectionBackColor
                .Forecolor = CustomSelectionForeColor
                
            End With
            
            
            'Destaca a imagem
            With frameSubmenu.Controls("img_" & xlSubmenus.Name)
                .BackColor = CustomSelectionBackColor
                .BackStyle = 1
            End With
            
            'Destaca a setinha se for item de submenu
            On Error Resume Next
                With frameSubmenu.Controls("arrow_" & xlSubmenus.Name)
                    .Forecolor = CustomSelectionForeColor
                End With
            On Error GoTo 0
            
            
            'Exibe o submenu correspondente
            With form.Controls("sub_" & xlSubmenus.Name)
                'Quando menu for fechado, ele fica com o item selecionado
                'A condição abaixo verifica a tag contém número ou texto,
                'se for texto, é o nome do item que estava salvo,
                'com isso volto ao estado padrão
                If .Tag <> "" And Not IsNumeric(.Tag) Then
                    form.Controls(.Tag).BackColor = CustomBackColor
                    form.Controls(.Tag).Forecolor = CustomForeColor
                    form.Controls("img_" & .Tag).BackColor = CustomBackColor
                    form.Controls("img_" & .Tag).BackStyle = 0
                    
                    'O erro é para tratar o caso quanto o item não é um submenu
                    'pois o mesmo não possui a setinha
                    On Error Resume Next
                    form.Controls("arrow_" & .Tag).Forecolor = CustomForeColor
                    On Error GoTo 0
                    
                End If
                
                .Visible = True
                .Left = frameSubmenu.Left + frameSubmenu.Width - 1
                .Top = frameSubmenu.Top + xlSubmenus.Top
                .Tag = ""
                
                'Loop para habilitar e desabilitar as imagens
                For Each ctrl In .Controls
                
                    If ctrl.Name <> "lbSep" And Left(ctrl.Name, 4) <> "img_" And Left(ctrl.Name, 6) <> "arrow_" Then
                        .Controls("img_" & ctrl.Name).Enabled = ctrl.Enabled
                    End If
            
                Next
                
                
            End With

            
            'Guarda os submenus que foram exibidos
            With form.Controls("lb_SubmenuNames")
            
                'Se a tag estiver vazia o nome será salvo normalmente
                If .Tag = "" Then
                    .Tag = "sub_" & xlSubmenus.Name
                Else
                
                   'Se não estiver salvo, será unificado com o nome que
                   'já contém, separado por um delimitador
                    If Not .Tag Like "*sub_" & xlSubmenus.Name & "*" Then
                        .Tag = .Tag & "|" & "sub_" & xlSubmenus.Name
                    End If
                    
                End If
                
                
                If .Caption <> "" Then
                    'Se a Caption que contém o nome do último submenu exibido
                    'for difente do submenu atual entraremos na condição para
                    'ocultar o mesmo
                    If .Caption <> "sub_" & xlSubmenus.Name Then
                    
                        'Só que antes precisamos verificar se o submenu que está sendo exibido
                        'pertence ao mesmo pai do submenu que está sendo aberto
                        'se tiver, o submenu anterior será fechado
                        If form.Controls("lb_MenusExibidos") = xlSubmenus.Parent.Name Then
                            
                            'Só que antes preciso verificar se o submenu selecionado
                            'não pertence ao submenu que será fechado,
                            'se não for eu não oculto, pois o submenu não pode fechar
                            'o submenu onde o mesmo pertence
                            If xlSubmenus.Parent.Name <> .Caption Then
                                form.Controls(.Caption).Visible = 0
                            End If
                            
                        End If
                        
                    End If
                End If
                
                
                'Guardo para ser utilizado para ocultar o nome do submenu
                'quando estou com o cursor do mouse sobre um item
                .Caption = "sub_" & xlSubmenus.Name
                form.Controls("lb_MenusExibidos") = xlSubmenus.Parent.Name
                
            End With

 
    End If
    

    
End Sub

Private Function ReturnForm(ByVal ctrl As Object) As Object
    
    On Error GoTo Parar
    
    While True
        Set ctrl = ctrl.Parent
    Wend
    
Parar:

    Set ReturnForm = ctrl
    
    
End Function


Private Sub OcultarSubmenuArrow(arrow As MSForms.Label)
        
    Dim arrSubmenus, i As Integer
            
    'Cria o array com os nomes dos submenus foram exibidos
    arrSubmenus = Split(form.Controls("lb_SubmenuNames").Tag, "|")
    
    'Um loop será feito por cada nome do submenu que foi exibido
    For i = UBound(arrSubmenus) To 0 Step -1
        
        'Se o nome do submenu for diferente do mome do menu/submenu
        'que o item pertence, esses submenus serão ocultados
        If arrow.Parent.Name <> arrSubmenus(i) Then
        
            'O submenu não pode fechar o seu próprio nível
            'por isso verifico se os nomes são distintos para que o mesmo
            'possa realizar o fechamento
            If form.Controls("lb_SubmenuNames") <> arrSubmenus(i) Then
                form.Controls(arrSubmenus(i)).Visible = False
            End If
        Else
            Exit For
        End If
        
    Next
    
End Sub

Private Sub CheckForm()
    
    If form Is Nothing Then
        MsgBox "The 'UserForm' property of the" & vbLf & "class has not been instantiated!" & _
        vbNewLine & "----------------------------------------------" & _
        vbNewLine & "Ex: Set Menu.UserForm = Me", 16, "  Menu Strip"
        End
    End If
    
End Sub


'====================================================================================================================
'==============================================| Eventos |===========================================================
'====================================================================================================================

Private Sub xlMenuStrip_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim ctrl As MSForms.Control
    
    'Define a variavel form antes de chamar a macro de remover o efeito
    Set form = xlMenuStrip.Parent
    
    
    'Macro para retirar o efeito de seleção da faixa do menu
    'ocultar os menus e submenus que estivem exibidos
    If form.Controls("lb_FaixaMenu").Tag <> xlMenuStrip.Name Then
        Call RemoverEfeitoMenuStrip
    End If
    
    'Oculta todos os submenus que estiverem exibidos
    Call OcultarSubmenus
        
    'Destaca o menu que o mouse está posicionado
    With xlMenuStrip
        .Forecolor = CustomSelectionForeColor
        .BackColor = CustomSelectionBackColor
    End With
    
    'Salva o nome do menu que foi destacado para retirar
    'o efeito quando o mesmo não estiver sob o ponteiro do mouse
    form.Controls("lb_FaixaMenu").Tag = xlMenuStrip.Name
    
    'Habilita a label que será responsável por retirar o efeito
    'e fechar o menu quando não estiverem sob o ponteiro do mouse
    form.Controls("lb_Close").Visible = True
    
    
    'Exibe o menu e alinha
    With form.Controls("fr_" & xlMenuStrip.Name)
        
        .Visible = True
        .Left = xlMenuStrip.Left
        
        'Se a tag não for varia, ela contem o nome do menu
        'nesse caso, removo a seleção do item
        If .Tag <> "" Then
            Call RemoverSelecao(form.Controls(.Tag))
        End If
        
        'Loop para habilitar e desabilitar as imagens
        For Each ctrl In .Controls
        
            If ctrl.Name <> "lbSep" And Left(ctrl.Name, 4) <> "img_" And Left(ctrl.Name, 6) <> "arrow_" Then
                .Controls("img_" & ctrl.Name).Enabled = ctrl.Enabled
            End If
    
        Next
        
        
    End With

    
End Sub


Private Sub xlClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim arrListBox, i As Byte
    
    'Define a variavel form antes de chamar a macro de remover o efeito
    Set form = xlClose.Parent
    
    'Macro para retirar o efeito de seleção da faixa do menu
    Call RemoverEfeitoMenuStrip
    Call OcultarSubmenus
    
    'Para evitar repetição, a label ficará invisível após
    'redefinir o menu para o padrão
    xlClose.Visible = False
    
    
    'Processo para corrigir o bug da listbox
    If form.Controls("lb_ListBox").Caption <> "" Then
    
        arrListBox = Split(form.Controls("lb_ListBox").Caption, "|")
        
        For i = 0 To UBound(arrListBox)
            form.Controls(arrListBox(i)).Left = form.Controls(arrListBox(i)).Left + 1
            form.Controls(arrListBox(i)).Left = form.Controls(arrListBox(i)).Left - 1
        Next
        
    End If
    
End Sub




Private Sub xlSubmenus_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Set form = xlSubmenus.Parent.Parent
    
    'Verifica se o item já não está selecionada
    'para evitar a execução do código de forma desnecessária
    If xlSubmenus.BackColor <> CustomSelectionBackColor Then
        Call EventSubmenuItems_MouseMove(xlSubmenus)
    End If
    
End Sub

Private Sub xlItems_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Set form = xlItems.Parent.Parent
    
    'Verifica se o item já não está selecionada
    'para evitar a execução do código de forma desnecessária
    If xlItems.BackColor <> CustomSelectionBackColor Then
       Call EventItems_MouveMove(xlItems)
    End If
    
End Sub

Private Sub xlItems_Click()
    Call EventItems_Click(xlItems)
End Sub

Private Sub xlImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim lbItem As MSForms.Label
    
    Set form = xlImage.Parent.Parent
    
    'Verifica se a imagem já não está selecionada
    'para evitar a execução do código de forma desnecessária
    If xlImage.BackColor <> CustomSelectionBackColor Then
    
        'Remove o prexixo "img_" para criar a label correspondente
        Set lbItem = xlImage.Parent(Mid(xlImage.Name, 5))
        
         'Chama a macro informando a label
        If lbItem.Tag <> "" Then
            'Se tivar valor na tag, o evento de um item
            Call EventItems_MouveMove(lbItem)
        Else
            'Se tiver vazio é a sub de um submenu
            Call EventSubmenuItems_MouseMove(lbItem)
        End If
    
    End If
    
End Sub

Private Sub xlImage_Click()

    Dim lbItem As MSForms.Label
    
    'Remove o prexixo "img_" para criar a label correspondente
    Set lbItem = xlImage.Parent(Mid(xlImage.Name, 5))
    
    'Verifica se a tag não está vazia, pois poder ser
    'a imagem relacionada ao submenu que não executa macros
    If lbItem.Tag <> "" Then
        'Chama a macro informando a label
        Call EventItems_Click(lbItem)
    End If
    
End Sub

Private Sub xlArrow_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim lbSeta As MSForms.Label
    
    
    
    Set form = xlArrow.Parent.Parent
    
    'Verifica se a cor do preenchimento não é a cor de seleção
    'assim evita a execução do código de forma desnecessária
    If xlArrow.BackColor <> CustomSelectionBackColor Then
        
        With form.Controls("lb_SubmenuNames")
        
            If .Caption <> "" Then
                'Verifica se o submenu salvo é diferente do atual
                'Esse caso é específico para quando o usuário estiver
                'voltado para o menu anterior
                If .Caption <> xlArrow.Tag Then
                
                    'Caso seja, o submenu anterior será fechado
                    'Entretanto, preciso verificar se o item pertence
                    'ao submenu que está salvo, se pertencer não posso
                    'ocultar, pois o item não pode fechar o submenu
                    'onde o mesmo se encontra
                    If .Caption <> xlArrow.Parent.Name Then
                        form.Controls(.Caption).Visible = 0
                    End If
                    
                End If
            End If
            
        End With
        
        
        'Remove o prexixo "arrow_" para criar a label correspondente
        Set lbSeta = xlArrow.Parent(Mid(xlArrow.Name, 7))
        
        'Chama a macro informando a label
        Call EventSubmenuItems_MouseMove(lbSeta)
        
        
        'Fechamento especial
        Call OcultarSubmenuArrow(xlArrow)

    
    End If

End Sub

Private Sub xlMultiPage_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Set form = ReturnForm(xlMultiPage)
    
    If form.Controls("lb_FaixaMenu").Tag <> "" Then
    
        form.Controls("lb_Close").Visible = False
        Call RemoverEfeitoMenuStrip
        Call OcultarSubmenus
    
    End If
    
End Sub

Private Sub xlListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        
    Set form = ReturnForm(xlListBox)

    If form.Controls("lb_FaixaMenu").Tag <> "" Then
    
        form.Controls("lb_Close").Visible = False
        Call RemoverEfeitoMenuStrip
        Call OcultarSubmenus
        xlListBox.Left = xlListBox.Left + 1
        xlListBox.Left = xlListBox.Left - 1
    End If
    
End Sub

Private Sub xlFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Set form = ReturnForm(xlFrame)
    
    If form.Controls("lb_FaixaMenu").Tag <> "" Then
    
        form.Controls("lb_Close").Visible = False
        Call RemoverEfeitoMenuStrip
        Call OcultarSubmenus
    
    End If
    
End Sub

Private Sub xlFrameMenu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        
    Set form = xlFrameMenu.Parent
    
    With xlFrameMenu
        
        If .Tag <> "" Then
        
            If .Controls(.Tag).BackColor <> CustomBackColor Then
                
                'Remover a seleção do item
                .Controls(.Tag).BackColor = CustomBackColor
                .Controls(.Tag).Forecolor = CustomForeColor
                .Controls("img_" & .Tag).BackColor = CustomBackColor
                
                On Error Resume Next
                .Controls("arrow_" & .Tag).Forecolor = CustomForeColor
                
                'Fechar os submenus
                Call OcultarSubmenu(, xlFrameMenu)
                
            End If
        
        End If
        
    End With
    
End Sub

Private Sub xlSeparator_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim xlFrameMenu As MSForms.frame
    
    Set form = xlSeparator.Parent.Parent
    Set xlFrameMenu = xlSeparator.Parent
    
    With xlFrameMenu
        
        If .Tag <> "" Then
        
            If .Controls(.Tag).BackColor <> CustomBackColor Then
                
                'Remover a seleção do item
                .Controls(.Tag).BackColor = CustomBackColor
                .Controls(.Tag).Forecolor = CustomForeColor
                .Controls("img_" & .Tag).BackColor = CustomBackColor
                
                On Error Resume Next
                .Controls("arrow_" & .Tag).Forecolor = CustomForeColor
                
                'Fechar os submenus
                Call OcultarSubmenu(xlSeparator)
                
            End If
        
        End If
        
    End With
    
End Sub

