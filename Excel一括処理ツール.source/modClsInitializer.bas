Attribute VB_Name = "modClsInitializer"
Option Explicit

Public Function GetLoggerType(prmLoggerTypeStr As String) As LOGGER_TYPE
    Dim loggerType As LOGGER_TYPE
    
    Select Case prmLoggerTypeStr
        Case "�f�o�b�O"
            loggerType = LOGGER_TYPE_DEBUGPRINT
        Case "�t�@�C��"
            loggerType = LOGGER_TYPE_LOGFILE
        Case "�V�[�g"
            loggerType = LOGGER_TYPE_LOGSHEET
        Case Else
            loggerType = ""
    End Select
    GetLoggerType = loggerType
End Function

Public Function GetLogOutputMode(prmLogOutputModeStr As String) As LogOutputMode
    Dim outputMode As LogOutputMode
    
    Select Case prmLogOutputModeStr
        Case "�t�@�C�����Œ�"
            outputMode = FileLogFixedLogName
        Case "�����t�t�@�C����"
            outputMode = FileLogTimeBasedName
        Case "�V���v�����O"
            outputMode = SheetLogSimple
        Case "�Œ�񃍃O"
            outputMode = SheetLogFormatted
    End Select
    GetLogOutputMode = outputMode
End Function

Public Function InitLogger(prmLoggerType As LOGGER_TYPE, param1 As String, param2 As String, outputMode As LogOutputMode, maxLogSize As Long, useModuleName As Boolean) As clsLogger
    Dim log As New clsLogger
    Dim ret As VbMsgBoxResult
    
    If prmLoggerType = LOGGER_TYPE_DEBUGPRINT Then
        Call log.InitializeLogger(LOGGER_TYPE_DEBUGPRINT, "", "", DebugLogNormal, maxLogSize, useModuleName)
    ElseIf prmLoggerType = LOGGER_TYPE_LOGFILE Then
        Dim filePath As String
        Dim fileName As String
        If param1 <> "" And param2 <> "" Then
            filePath = param1
            fileName = param2
        Else
            ret = MsgBox("�p�����[�^���s���ł�" & vbCrLf & "param1:" & param1 & vbCrLf & "param2:" & param2, vbOKOnly, "�G���[")
            Exit Function
        End If
        Call log.InitializeLogger(LOGGER_TYPE_LOGFILE, filePath, fileName, DebugLogNormal, maxLogSize, useModuleName)
    ElseIf prmLoggerType = LOGGER_TYPE_LOGSHEET Then
        Dim sheetName As String
        Dim startLogCell As String
        If sheetName <> "" And startLogCell <> "" Then
            sheetName = param1
            startLogCell = param2
        Else
            ret = MsgBox("�p�����[�^���s���ł�" & vbCrLf & "param1:" & param1 & vbCrLf & "param2:" & param2, vbOKOnly, "�G���[")
            Exit Function
        End If
        Call log.InitializeLogger(LOGGER_TYPE_LOGSHEET, sheetName, startLogCell, outputMode, maxLogSize, useModuleName)
    Else
        Exit Function
    End If
    Set InitLogger = log
End Function




