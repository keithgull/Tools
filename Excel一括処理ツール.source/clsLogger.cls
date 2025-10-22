VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ====== Class Module: Logger ==================
' VBA 汎用Logger
'  依存関係
'   clsLogger
'    + iLoggerOutput - clsLoggerDebugPrintOutput    ：debugprint用のロガー
'    + iLoggerOutput - clsLoggerFileOutput          ：ファイル出力用のロガー
'    + iLoggerOutput - clsLoggerSheetOutput         ；シート出力用のロガー
'
'  使い方
'    Dim log As New clsLogger
'
'    ' Enumを使って初期化
'    log.InitializeLogger LOGGER_TYPE_DEBUGPRINT, "", "", DebugLogNormal, 10                                      ' Debug.printを使ったログ
'    ' log.InitializeLogger LOGGER_TYPE_LOGFILE, "C:\Work", "vbalogout.log",  FileLogTimeBasedName, 10000000      ' ログファイルへ出力するログ
'    ' log.InitializeLogger LOGGER_TYPE_LOGSHEET, "Sheet1", "C10", SheetLogFormatted, 100000                      ' シートに出力するログ
'
'    log.SetLogLevel LOGLEVEL_DEBUG, True                                               ' ログレベルの個別設定
'    log.SetLogLevelAll true, true, true, true, true                                    ' ログレベルの一括設定（INFO、ERROR、WARN、DEBUG、TRACEの順で出力・非出力を設定）
'
'    log.InfoLog "開始"
'    log.DebugLog "デバッグ情報"
'

#If VBA7 Then
    Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
    Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Enum LogOutputMode
    DebugLogNormal
    FileLogFixedLogName
    FileLogTimeBasedName
    SheetLogSimple
    SheetLogFormatted
End Enum

Private logInstance As clsLogger
Private outputHandler As iLoggerOutput
Private logLevels As Scripting.Dictionary
Private formatTemplate_ As String
Private timeFormat_ As String
Private useModuleName_ As Boolean
Private moduleNameLen_ As Integer

Private Sub Class_Initialize()
    'Set logInstance = GetLogger()
    Set logLevels = New Scripting.Dictionary
    logLevels.Add LOGLEVEL_INFO, True
    logLevels.Add LOGLEVEL_DEBUG, False
    logLevels.Add LOGLEVEL_WARN, True
    logLevels.Add LOGLEVEL_ERROR, True
    logLevels.Add LOGLEVEL_TRACE, False
    useModuleName_ = False
End Sub

' シングルトン取得
Public Function GetLogger() As clsLogger
    If logInstance Is Nothing Then
        Set logInstance = New clsLogger
    End If
    Set GetLogger = logInstance
End Function

' 出力先の初期化
Public Sub InitializeLogger(ByVal loggerType As LOGGER_TYPE, ByVal param1 As String, ByVal param2 As String, param3 As LogOutputMode, Optional maxLogSize As Long = 10000000, Optional useModuleName As Boolean = False)
    Select Case loggerType
        Case LOGGER_TYPE_DEBUGPRINT
            Dim dlog As New clsLoggerDebugPrintOutput
            dlog.LogFormat = FILE_LOG_FORMAT_WITHOUT_MODULENAME
            useModuleName_ = useModuleName
            Set outputHandler = dlog

        Case LOGGER_TYPE_LOGFILE
            Dim flog As New clsLoggerFileOutput
            Dim logFileName As String
            flog.LogFormat = FILE_LOG_FORMAT_WITHOUT_MODULENAME
            logFileName = GetLogFileName(param3, param2)
            flog.Init param1, logFileName, param3, maxLogSize
            
            Set outputHandler = flog

        Case LOGGER_TYPE_LOGSHEET
            Dim slog As New clsLoggerSheetOutput
            Dim ws As Worksheet
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(param1)
            On Error GoTo 0
            If ws Is Nothing Then Err.Raise vbObjectError + 1, , "指定されたシートが存在しません: " & param1

            Dim col As Long, row As Long
            col = Range(param2).Column
            row = Range(param2).row
            slog.LogFormat = SHEET_LOG_FORMAT_WITHOUT_MODULENAME
            slog.TimeFormat = ("yyyy/MM/dd hh:mm:ss.000")
            slog.Init ws, col, row, param3, maxLogSize
            Set outputHandler = slog

        Case Else
            Err.Raise vbObjectError + 2, , "不明なログ種別: " & loggerType
    End Select
End Sub

' logファイル名の作成
Public Function GetLogFileName(fileLogOutputMode As LogOutputMode, baseFileName As String) As String
    Dim logFileName As String
    Dim dotpos As Integer
    If fileLogOutputMode = FileLogTimeBasedName Then
        ' 時間ベースのログ・ファイル名の場合
        Dim logExtension As String
        dotpos = InStrRev(baseFileName, ".")
        If dotpos > 0 Then
            ' 拡張子ありの場合
            logExtension = Right(baseFileName, Len(baseFileName) - dotpos)
            logFileName = Left(baseFileName, dotpos - 1) & "_" & Format(Now, "yyyymmdd_hhmm") & "." & logExtension
        Else
            ' 拡張子なしの場合
            logFileName = baseFileName & "_" & Format(Now, "yyyymmdd_hhmm") & ".log"
        End If
    ElseIf fileLogOutputMode = FileLogFixedLogName Then
        dotpos = InStrRev(baseFileName, ".")
        If dotpos > 0 Then
            logFileName = baseFileName
        Else
            logFileName = baseFileName & ".log"
        End If
    End If
    GetLogFileName = logFileName
End Function


Public Sub SetOutput(ByVal handler As iLoggerOutput)
    Set outputHandler = handler
End Sub

Public Sub SetLogLevel(ByVal level As String, ByVal enabled As Boolean)
    If logLevels.Exists(level) Then
        logLevels(level) = enabled
    End If
End Sub

Public Sub SetLogLevelAll(ByVal infoLvl As Boolean, ByVal errorLvl As Boolean, warnLvl As Boolean, debugLvl As Boolean, traceLvl As Boolean)
    SetLogLevel LOGLEVEL_INFO, infoLvl
    SetLogLevel LOGLEVEL_ERROR, errorLvl
    SetLogLevel LOGLEVEL_WARN, warnLvl
    SetLogLevel LOGLEVEL_DEBUG, debugLvl
    SetLogLevel LOGLEVEL_TRACE, traceLvl
End Sub

Public Sub SetUseModuleName(ByVal useFlag As Boolean, Optional nameLen As Integer = 0)
    useModuleName_ = useFlag
    moduleNameLen_ = nameLen
End Sub

Public Sub SetLogFormat(ByVal template As String)
    formatTemplate_ = template
End Sub

Public Sub log(ByVal level As String, ByVal moduleName As String, ByVal message As String)
    If Not logLevels.Exists(level) Or Not logLevels(level) Then
        Exit Sub
    End If

    Dim formatted As String
    Dim formattedDatetime As String
    Dim formattedLogLevel As String
    formattedDatetime = Format(Now, timeFormat_)
    formattedLogLevel = "[" & Format(level, "@@@@@") & "]" ' 5文字で桁揃え
    formatted = outputHandler.LogFormat
    formatted = Replace(formatted, "{time}", formattedDatetime)
    formatted = Replace(formatted, "{level}", formattedLogLevel)
    formatted = IIf(useModuleName_, Replace(formatted, "{module}", moduleName), "")
    formatted = Replace(formatted, "{message}", message)
    
    If Not outputHandler Is Nothing Then
        outputHandler.OutputLogMsg formatted
    End If
End Sub

' 各ログレベル用メソッド
Public Sub InfoLogM(ByVal moduleName As String, ByVal message As String)
    OutputLog LOGLEVEL_INFO, moduleName, message
End Sub

Public Sub DebugLogM(ByVal moduleName As String, ByVal message As String)
    OutputLog LOGLEVEL_DEBUG, moduleName, message
End Sub

Public Sub WarnLogM(ByVal moduleName As String, ByVal message As String)
    OutputLog LOGLEVEL_WARN, moduleName, message
End Sub

Public Sub ErrorLogM(ByVal moduleName As String, ByVal message As String)
    OutputLog LOGLEVEL_ERROR, moduleName, message
End Sub

Public Sub TraceLogM(ByVal moduleName As String, ByVal message As String)
    OutputLog LOGLEVEL_TRACE, moduleName, message
End Sub

Public Sub InfoLog(ByVal message As String)
    OutputLog LOGLEVEL_INFO, "", message
End Sub

Public Sub DebugLog(ByVal message As String)
    OutputLog LOGLEVEL_DEBUG, "", message
End Sub

Public Sub WarnLog(ByVal message As String)
    OutputLog LOGLEVEL_WARN, "", message
End Sub

Public Sub ErrorLog(ByVal message As String)
    OutputLog LOGLEVEL_ERROR, "", message
End Sub

Public Sub TraceLog(ByVal message As String)
    OutputLog LOGLEVEL_TRACE, "", message
End Sub


' 共通ログ出力処理
Private Sub OutputLog(ByVal level As String, ByVal moduleName As String, ByVal message As String)
    If Not logLevels.Exists(level) Or Not logLevels(level) Then Exit Sub

    Dim paddedLevel As String
    paddedLevel = "[" & Format(level, "@@@@@") & "]" ' 5文字で桁揃え

    Dim formatted As String
    formatted = outputHandler.LogFormat
    formatted = Replace(formatted, "{time}", GetPreciseTime("-"))
    formatted = Replace(formatted, "{level}", paddedLevel)
    formatted = IIf(useModuleName_, Replace(formatted, "{module}", moduleName), formatted)
    formatted = Replace(formatted, "{message}", message)

    If Not outputHandler Is Nothing Then
        outputHandler.OutputLogMsg formatted
    End If
End Sub

Function GetPreciseTime(datedelim As String) As String
    Dim st As SYSTEMTIME
    GetLocalTime st
    GetPreciseTime = Format(st.wYear, "0000") & datedelim & _
                     Format(st.wMonth, "00") & datedelim & _
                     Format(st.wDay, "00") & " " & _
                     Format(st.wHour, "00") & ":" & _
                     Format(st.wMinute, "00") & ":" & _
                     Format(st.wSecond, "00") & "." & _
                     Format(st.wMilliseconds, "000")
End Function

Property Let FormatTemplate(formatStr As String)
    formatTemplate_ = formatStr
End Property

Property Get FormatTemplate() As String
    FormatTemplate = formatTemplate_
End Property

Property Let TimeFormat(formatTimeStr As String)
    timeFormat_ = formatTimeStr
End Property

Property Get TimeFormat() As String
    TimeFormat = timeFormat_
End Property
