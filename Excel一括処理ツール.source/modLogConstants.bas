Attribute VB_Name = "modLogConstants"
' Module: LoggerConstants
Option Explicit

' ログレベル定数（表示用に桁揃え済み）
Public Const LOGLEVEL_INFO  As String = "INFO "
Public Const LOGLEVEL_DEBUG As String = "DEBUG"
Public Const LOGLEVEL_WARN  As String = "WARN "
Public Const LOGLEVEL_ERROR As String = "ERROR"
Public Const LOGLEVEL_TRACE As String = "TRACE"

' ログレベルキー（内部制御用）
'Public Const LOGKEY_INFO  As String = "INFO"
'Public Const LOGKEY_DEBUG As String = "DEBUG"
'Public Const LOGKEY_WARN  As String = "WARN"
'Public Const LOGKEY_ERROR As String = "ERROR"
'Public Const LOGKEY_TRACE As String = "TRACE"

' Enum定義：ログ種別
Public Enum LOGGER_TYPE
    LOGGER_TYPE_DEBUGPRINT = 0
    LOGGER_TYPE_LOGFILE = 1
    LOGGER_TYPE_LOGSHEET = 2
End Enum

Public Const FILE_LOG_FORMAT_DEFAULT As String = "{time} {level} {module} {message}"
Public Const SHEET_LOG_FORMAT_DEFAULT As String = "{time}\t{level}\t{module}\t{message}"

Public Const FILE_LOG_FORMAT_WITHOUT_MODULENAME As String = "{time} {level} {message}"
Public Const SHEET_LOG_FORMAT_WITHOUT_MODULENAME As String = "{time}\t{level}\t{message}"

