'====================================================================
' MÓDULO LOG MANAGER - VERSÃO REFATORADA PROFISSIONAL
' Responsável por todas as operações de log do sistema
' Data/Hora: 2025-01-27
' Versão: PROFESSIONAL v3.0
' Desenvolvedor: Sistema PDV Enterprise
'====================================================================

Option Explicit

' ===== CONSTANTES =====
Private Const LOG_SHEET_NAME As String = "SystemLog"
Private Const LOG_FILE_PATH As String = "C:\Logs\PDV_System\"
Private Const MAX_LOG_ENTRIES As Long = 10000
Private Const LOG_RETENTION_DAYS As Long = 30

' ===== ENUMERAÇÕES =====
Private Enum LogLevel
    Debug = 0
    Info = 1
    Warning = 2
    Error = 3
    Critical = 4
End Enum

Private Enum LogCategory
    System = 0
    User = 1
    Database = 2
    Security = 3
    Performance = 4
    Business = 5
End Enum

' ===== ESTRUTURAS DE DADOS =====
Private Type LogEntry
    Timestamp As Date
    Level As LogLevel
    Category As LogCategory
    Module As String
    Procedure As String
    Message As String
    UserID As String
    SessionID As String
    AdditionalData As String
End Type

Private Type LogConfiguration
    EnableFileLogging As Boolean
    EnableSheetLogging As Boolean
    EnableConsoleLogging As Boolean
    MinLogLevel As LogLevel
    MaxEntries As Long
    RetentionDays As Long
    LogFilePath As String
End Type

' ===== VARIÁVEIS DE INSTÂNCIA =====
Private mLogConfig As LogConfiguration
Private mSessionID As String
Private mCurrentUser As String
Private mLogBuffer As Collection
Private mIsInitialized As Boolean

'====================================================================
' INICIALIZAÇÃO DO MÓDULO
'====================================================================
Private Sub Class_Initialize()
    Call InitializeLogConfiguration
    Call GenerateSessionID
    Set mLogBuffer = New Collection
    mIsInitialized = True
    
    ' Log de inicialização
    Call LogInfo("LogManager inicializado", "Class_Initialize")
End Sub

Private Sub InitializeLogConfiguration()
    With mLogConfig
        .EnableFileLogging = True
        .EnableSheetLogging = True
        .EnableConsoleLogging = False
        .MinLogLevel = LogLevel.Info
        .MaxEntries = MAX_LOG_ENTRIES
        .RetentionDays = LOG_RETENTION_DAYS
        .LogFilePath = LOG_FILE_PATH
    End With
End Sub

Private Sub GenerateSessionID()
    ' Gerar ID único para a sessão
    mSessionID = "SESS_" & Format(Now, "yyyymmdd_hhnnss") & "_" & Int(Rnd() * 10000)
End Sub

'====================================================================
' MÉTODOS PRINCIPAIS DE LOG
'====================================================================

' === LOG DE INFORMAÇÃO ===
Public Sub LogInfo(message As String, procedureName As String, Optional moduleName As String = "", Optional additionalData As String = "")
    Call WriteLog(LogLevel.Info, LogCategory.System, moduleName, procedureName, message, additionalData)
End Sub

' === LOG DE AVISO ===
Public Sub LogWarning(message As String, procedureName As String, Optional moduleName As String = "", Optional additionalData As String = "")
    Call WriteLog(LogLevel.Warning, LogCategory.System, moduleName, procedureName, message, additionalData)
End Sub

' === LOG DE ERRO ===
Public Sub LogError(procedureName As String, err As ErrObject, Optional moduleName As String = "", Optional additionalData As String = "")
    Dim message As String
    message = "Erro: " & err.Description & " (Código: " & err.Number & ")"
    
    Call WriteLog(LogLevel.Error, LogCategory.System, moduleName, procedureName, message, additionalData)
End Sub

' === LOG DE ERRO CRÍTICO ===
Public Sub LogCriticalError(procedureName As String, err As ErrObject, Optional moduleName As String = "", Optional additionalData As String = "")
    Dim message As String
    message = "ERRO CRÍTICO: " & err.Description & " (Código: " & err.Number & ")"
    
    Call WriteLog(LogLevel.Critical, LogCategory.System, moduleName, procedureName, message, additionalData)
End Sub

' === LOG DE DEBUG ===
Public Sub LogDebug(message As String, procedureName As String, Optional moduleName As String = "", Optional additionalData As String = "")
    Call WriteLog(LogLevel.Debug, LogCategory.System, moduleName, procedureName, message, additionalData)
End Sub

' === LOG DE USUÁRIO ===
Public Sub LogUserAction(action As String, userID As String, Optional additionalData As String = "")
    Call WriteLog(LogLevel.Info, LogCategory.User, "UserManager", "LogUserAction", action, additionalData)
End Sub

' === LOG DE SEGURANÇA ===
Public Sub LogSecurityEvent(eventType As String, userID As String, Optional additionalData As String = "")
    Call WriteLog(LogLevel.Warning, LogCategory.Security, "SecurityManager", "LogSecurityEvent", eventType, additionalData)
End Sub

' === LOG DE PERFORMANCE ===
Public Sub LogPerformance(operation As String, duration As Double, Optional additionalData As String = "")
    Dim message As String
    message = "Operação: " & operation & " - Duração: " & Format(duration, "0.000") & " segundos"
    
    Call WriteLog(LogLevel.Info, LogCategory.Performance, "PerformanceManager", "LogPerformance", message, additionalData)
End Sub

' === LOG DE NEGÓCIO ===
Public Sub LogBusinessEvent(eventType As String, eventData As String, Optional additionalData As String = "")
    Call WriteLog(LogLevel.Info, LogCategory.Business, "BusinessManager", "LogBusinessEvent", eventType & ": " & eventData, additionalData)
End Sub

'====================================================================
' MÉTODO PRINCIPAL DE ESCRITA DE LOG
'====================================================================
Private Sub WriteLog(level As LogLevel, category As LogCategory, moduleName As String, procedureName As String, message As String, additionalData As String)
    On Error GoTo ErrorHandler
    
    ' Verificar se o nível de log é suficiente
    If level < mLogConfig.MinLogLevel Then
        Exit Sub
    End If
    
    ' Criar entrada de log
    Dim logEntry As LogEntry
    With logEntry
        .Timestamp = Now
        .Level = level
        .Category = category
        .Module = IIf(moduleName = "", "Unknown", moduleName)
        .Procedure = IIf(procedureName = "", "Unknown", procedureName)
        .Message = message
        .UserID = mCurrentUser
        .SessionID = mSessionID
        .AdditionalData = additionalData
    End With
    
    ' Adicionar ao buffer
    Call AddToLogBuffer(logEntry)
    
    ' Escrever logs baseado na configuração
    If mLogConfig.EnableSheetLogging Then
        Call WriteToSheet(logEntry)
    End If
    
    If mLogConfig.EnableFileLogging Then
        Call WriteToFile(logEntry)
    End If
    
    If mLogConfig.EnableConsoleLogging Then
        Call WriteToConsole(logEntry)
    End If
    
    Exit Sub

ErrorHandler:
    ' Em caso de erro no log, tentar escrever em arquivo de emergência
    Call WriteEmergencyLog("Erro no sistema de log: " & Err.Description)
End Sub

'====================================================================
' MÉTODOS DE ESCRITA DE LOG
'====================================================================

' === ESCREVER NO WORKSHEET ===
Private Sub WriteToSheet(logEntry As LogEntry)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetLogWorksheet
    
    If ws Is Nothing Then
        Call CreateLogWorksheet
        Set ws = GetLogWorksheet
    End If
    
    ' Obter próxima linha
    Dim nextRow As Long
    nextRow = GetNextLogRow(ws)
    
    ' Escrever dados
    With ws
        .Cells(nextRow, 1).Value = logEntry.Timestamp
        .Cells(nextRow, 2).Value = GetLogLevelString(logEntry.Level)
        .Cells(nextRow, 3).Value = GetLogCategoryString(logEntry.Category)
        .Cells(nextRow, 4).Value = logEntry.Module
        .Cells(nextRow, 5).Value = logEntry.Procedure
        .Cells(nextRow, 6).Value = logEntry.Message
        .Cells(nextRow, 7).Value = logEntry.UserID
        .Cells(nextRow, 8).Value = logEntry.SessionID
        .Cells(nextRow, 9).Value = logEntry.AdditionalData
    End With
    
    ' Aplicar formatação
    Call FormatLogRow(ws, nextRow, logEntry.Level)
    
    ' Verificar limite de entradas
    Call CheckLogLimits(ws)
    
    Exit Sub

ErrorHandler:
    Call WriteEmergencyLog("Erro ao escrever log na planilha: " & Err.Description)
End Sub

' === ESCREVER NO ARQUIVO ===
Private Sub WriteToFile(logEntry As LogEntry)
    On Error GoTo ErrorHandler
    
    ' Criar diretório se não existir
    Call CreateLogDirectory
    
    ' Obter nome do arquivo
    Dim fileName As String
    fileName = GetLogFileName
    
    ' Formatar linha de log
    Dim logLine As String
    logLine = FormatLogLine(logEntry)
    
    ' Escrever no arquivo
    Call AppendToFile(fileName, logLine)
    
    Exit Sub

ErrorHandler:
    Call WriteEmergencyLog("Erro ao escrever log no arquivo: " & Err.Description)
End Sub

' === ESCREVER NO CONSOLE ===
Private Sub WriteToConsole(logEntry As LogEntry)
    On Error GoTo ErrorHandler
    
    ' Usar Debug.Print para console
    Debug.Print FormatLogLine(logEntry)
    
    Exit Sub

ErrorHandler:
    ' Ignorar erros de console
End Sub

' === ESCREVER LOG DE EMERGÊNCIA ===
Private Sub WriteEmergencyLog(message As String)
    On Error Resume Next
    
    ' Tentar escrever em arquivo de emergência
    Dim emergencyFile As String
    emergencyFile = "C:\emergency_log.txt"
    
    Dim logLine As String
    logLine = Format(Now, "yyyy-mm-dd hh:mm:ss") & " - EMERGENCY: " & message & vbCrLf
    
    Call AppendToFile(emergencyFile, logLine)
End Sub

'====================================================================
' MÉTODOS AUXILIARES
'====================================================================

' === ADICIONAR AO BUFFER ===
Private Sub AddToLogBuffer(logEntry As LogEntry)
    On Error Resume Next
    
    mLogBuffer.Add logEntry
    
    ' Manter apenas as últimas entradas no buffer
    If mLogBuffer.Count > 100 Then
        mLogBuffer.Remove 1
    End If
End Sub

' === OBTER WORKSHEET DE LOG ===
Private Function GetLogWorksheet() As Worksheet
    On Error Resume Next
    Set GetLogWorksheet = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0
End Function

' === CRIAR WORKSHEET DE LOG ===
Private Sub CreateLogWorksheet()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    
    With ws
        .Name = LOG_SHEET_NAME
        
        ' Definir cabeçalhos
        .Cells(1, 1).Value = "Timestamp"
        .Cells(1, 2).Value = "Level"
        .Cells(1, 3).Value = "Category"
        .Cells(1, 4).Value = "Module"
        .Cells(1, 5).Value = "Procedure"
        .Cells(1, 6).Value = "Message"
        .Cells(1, 7).Value = "UserID"
        .Cells(1, 8).Value = "SessionID"
        .Cells(1, 9).Value = "AdditionalData"
        
        ' Formatar cabeçalhos
        Call FormatLogHeaders(ws)
    End With
    
    Exit Sub

ErrorHandler:
    Call WriteEmergencyLog("Erro ao criar worksheet de log: " & Err.Description)
End Sub

' === OBTER PRÓXIMA LINHA DE LOG ===
Private Function GetNextLogRow(ws As Worksheet) As Long
    GetNextLogRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
End Function

' === FORMATAR LINHA DE LOG ===
Private Function FormatLogLine(logEntry As LogEntry) As String
    Dim line As String
    
    line = Format(logEntry.Timestamp, "yyyy-mm-dd hh:mm:ss") & " | " & _
           GetLogLevelString(logEntry.Level) & " | " & _
           GetLogCategoryString(logEntry.Category) & " | " & _
           logEntry.Module & " | " & _
           logEntry.Procedure & " | " & _
           logEntry.Message & " | " & _
           logEntry.UserID & " | " & _
           logEntry.SessionID
    
    If logEntry.AdditionalData <> "" Then
        line = line & " | " & logEntry.AdditionalData
    End If
    
    FormatLogLine = line
End Function

' === OBTER STRING DO NÍVEL DE LOG ===
Private Function GetLogLevelString(level As LogLevel) As String
    Select Case level
        Case LogLevel.Debug
            GetLogLevelString = "DEBUG"
        Case LogLevel.Info
            GetLogLevelString = "INFO"
        Case LogLevel.Warning
            GetLogLevelString = "WARNING"
        Case LogLevel.Error
            GetLogLevelString = "ERROR"
        Case LogLevel.Critical
            GetLogLevelString = "CRITICAL"
        Case Else
            GetLogLevelString = "UNKNOWN"
    End Select
End Function

' === OBTER STRING DA CATEGORIA ===
Private Function GetLogCategoryString(category As LogCategory) As String
    Select Case category
        Case LogCategory.System
            GetLogCategoryString = "SYSTEM"
        Case LogCategory.User
            GetLogCategoryString = "USER"
        Case LogCategory.Database
            GetLogCategoryString = "DATABASE"
        Case LogCategory.Security
            GetLogCategoryString = "SECURITY"
        Case LogCategory.Performance
            GetLogCategoryString = "PERFORMANCE"
        Case LogCategory.Business
            GetLogCategoryString = "BUSINESS"
        Case Else
            GetLogCategoryString = "UNKNOWN"
    End Select
End Function

' === FORMATAR CABEÇALHOS ===
Private Sub FormatLogHeaders(ws As Worksheet)
    With ws.Range("A1:I1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Ajustar largura das colunas
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 20
    ws.Columns("F").ColumnWidth = 40
    ws.Columns("G").ColumnWidth = 15
    ws.Columns("H").ColumnWidth = 20
    ws.Columns("I").ColumnWidth = 30
End Sub

' === FORMATAR LINHA DE LOG ===
Private Sub FormatLogRow(ws As Worksheet, row As Long, level As LogLevel)
    ' Aplicar cor baseada no nível
    Select Case level
        Case LogLevel.Error, LogLevel.Critical
            ws.Rows(row).Interior.Color = RGB(255, 200, 200)
        Case LogLevel.Warning
            ws.Rows(row).Interior.Color = RGB(255, 255, 200)
        Case LogLevel.Info
            ws.Rows(row).Interior.Color = RGB(200, 255, 200)
        Case LogLevel.Debug
            ws.Rows(row).Interior.Color = RGB(240, 240, 240)
    End Select
    
    ' Aplicar bordas
    ws.Rows(row).Borders.LineStyle = xlContinuous
End Sub

' === VERIFICAR LIMITES DE LOG ===
Private Sub CheckLogLimits(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Se excedeu o limite, remover entradas antigas
    If lastRow > mLogConfig.MaxEntries Then
        Dim rowsToDelete As Long
        rowsToDelete = lastRow - mLogConfig.MaxEntries
        
        ' Manter cabeçalho e remover linhas antigas
        If rowsToDelete > 0 Then
            ws.Rows("2:" & (rowsToDelete + 1)).Delete
        End If
    End If
End Sub

' === CRIAR DIRETÓRIO DE LOG ===
Private Sub CreateLogDirectory()
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(mLogConfig.LogFilePath) Then
        fso.CreateFolder mLogConfig.LogFilePath
    End If
    
    Set fso = Nothing
End Sub

' === OBTER NOME DO ARQUIVO DE LOG ===
Private Function GetLogFileName() As String
    GetLogFileName = mLogConfig.LogFilePath & "PDV_Log_" & Format(Date, "yyyy-mm-dd") & ".txt"
End Function

' === ADICIONAR AO ARQUIVO ===
Private Sub AppendToFile(fileName As String, content As String)
    On Error GoTo ErrorHandler
    
    Dim fileNumber As Integer
    fileNumber = FreeFile
    
    Open fileName For Append As fileNumber
    Print #fileNumber, content
    Close fileNumber
    
    Exit Sub

ErrorHandler:
    Call WriteEmergencyLog("Erro ao escrever no arquivo " & fileName & ": " & Err.Description)
End Sub

'====================================================================
' MÉTODOS DE CONFIGURAÇÃO
'====================================================================

' === DEFINIR USUÁRIO ATUAL ===
Public Sub SetCurrentUser(userID As String)
    mCurrentUser = userID
    Call LogInfo("Usuário definido: " & userID, "SetCurrentUser")
End Sub

' === CONFIGURAR LOG ===
Public Sub ConfigureLogging(enableFileLogging As Boolean, enableSheetLogging As Boolean, minLogLevel As LogLevel)
    mLogConfig.EnableFileLogging = enableFileLogging
    mLogConfig.EnableSheetLogging = enableSheetLogging
    mLogConfig.MinLogLevel = minLogLevel
    
    Call LogInfo("Configuração de log atualizada", "ConfigureLogging")
End Sub

' === OBTER CONFIGURAÇÃO ===
Public Function GetLogConfiguration() As LogConfiguration
    GetLogConfiguration = mLogConfig
End Function

'====================================================================
' MÉTODOS DE MANUTENÇÃO
'====================================================================

' === LIMPAR LOGS ANTIGOS ===
Public Sub CleanOldLogs()
    On Error GoTo ErrorHandler
    
    ' Limpar logs da planilha
    Call CleanOldSheetLogs
    
    ' Limpar arquivos de log antigos
    Call CleanOldLogFiles
    
    Call LogInfo("Limpeza de logs antigos concluída", "CleanOldLogs")
    
    Exit Sub

ErrorHandler:
    Call LogError("CleanOldLogs", Err)
End Sub

Private Sub CleanOldSheetLogs()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetLogWorksheet
    
    If ws Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then Exit Sub
    
    ' Remover entradas mais antigas que o período de retenção
    Dim cutoffDate As Date
    cutoffDate = Date - mLogConfig.RetentionDays
    
    Dim i As Long
    For i = lastRow To 2 Step -1
        If CDate(ws.Cells(i, 1).Value) < cutoffDate Then
            ws.Rows(i).Delete
        End If
    Next i
    
    Exit Sub

ErrorHandler:
    Call WriteEmergencyLog("Erro ao limpar logs da planilha: " & Err.Description)
End Sub

Private Sub CleanOldLogFiles()
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(mLogConfig.LogFilePath) Then Exit Sub
    
    Dim folder As Object
    Set folder = fso.GetFolder(mLogConfig.LogFilePath)
    
    Dim file As Object
    For Each file In folder.Files
        If file.Name Like "PDV_Log_*.txt" Then
            If file.DateLastModified < (Date - mLogConfig.RetentionDays) Then
                file.Delete
            End If
        End If
    Next file
    
    Set fso = Nothing
    Exit Sub

ErrorHandler:
    Call WriteEmergencyLog("Erro ao limpar arquivos de log: " & Err.Description)
End Sub

' === EXPORTAR LOGS ===
Public Sub ExportLogs(filePath As String, Optional startDate As Date, Optional endDate As Date)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = GetLogWorksheet
    
    If ws Is Nothing Then
        Call LogWarning("Worksheet de log não encontrado", "ExportLogs")
        Exit Sub
    End If
    
    ' Criar workbook temporário
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    ' Copiar dados filtrados
    Call CopyFilteredLogs(ws, wb.Sheets(1), startDate, endDate)
    
    ' Salvar arquivo
    wb.SaveAs filePath
    wb.Close
    
    Call LogInfo("Logs exportados para: " & filePath, "ExportLogs")
    
    Exit Sub

ErrorHandler:
    Call LogError("ExportLogs", Err)
End Sub

Private Sub CopyFilteredLogs(sourceWs As Worksheet, targetWs As Worksheet, startDate As Date, endDate As Date)
    ' Copiar cabeçalhos
    sourceWs.Rows(1).Copy targetWs.Rows(1)
    
    Dim lastRow As Long
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    
    Dim targetRow As Long
    targetRow = 2
    
    Dim i As Long
    For i = 2 To lastRow
        Dim logDate As Date
        logDate = CDate(sourceWs.Cells(i, 1).Value)
        
        ' Verificar se está no período
        If (startDate = 0 Or logDate >= startDate) And (endDate = 0 Or logDate <= endDate) Then
            sourceWs.Rows(i).Copy targetWs.Rows(targetRow)
            targetRow = targetRow + 1
        End If
    Next i
End Sub

' === OBTER ESTATÍSTICAS DE LOG ===
Public Function GetLogStatistics() As Dictionary
    On Error GoTo ErrorHandler
    
    Set GetLogStatistics = New Dictionary
    
    Dim ws As Worksheet
    Set ws = GetLogWorksheet
    
    If ws Is Nothing Then
        GetLogStatistics.Add "TotalEntries", 0
        GetLogStatistics.Add "ErrorCount", 0
        GetLogStatistics.Add "WarningCount", 0
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim totalEntries As Long
    Dim errorCount As Long
    Dim warningCount As Long
    
    totalEntries = lastRow - 1 ' Excluir cabeçalho
    
    ' Contar por nível
    Dim i As Long
    For i = 2 To lastRow
        Dim level As String
        level = UCase(Trim(CStr(ws.Cells(i, 2).Value)))
        
        Select Case level
            Case "ERROR", "CRITICAL"
                errorCount = errorCount + 1
            Case "WARNING"
                warningCount = warningCount + 1
        End Select
    Next i
    
    GetLogStatistics.Add "TotalEntries", totalEntries
    GetLogStatistics.Add "ErrorCount", errorCount
    GetLogStatistics.Add "WarningCount", warningCount
    
    Exit Function

ErrorHandler:
    Call LogError("GetLogStatistics", Err)
    Set GetLogStatistics = New Dictionary
End Function