VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLoggerFileOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: FileOutput
Option Explicit
Implements iLoggerOutput

Private logFilePath_ As String
Private logPath_ As String
Private logFile_ As String
Private maxLogSize_ As Long
Private useModuleName_ As Boolean
Private logFileCount_ As Integer
Private fileLogOutputMode_ As LogOutputMode
Private logFormat_  As String
Private startLogTime_ As String

Public Sub Init(ByVal path As String, ByVal fileName As String, fileLogOutputMode As LogOutputMode, Optional maxLogSize As Long = 10000000, Optional useModuleName As Boolean = False)
    If Right(path, 1) <> "\" Then
        path = path & "\"
    End If
    maxLogSize_ = maxLogSize
    logFileCount_ = 0
    fileLogOutputMode_ = fileLogOutputMode
    useModuleName_ = useModuleName
    logPath_ = path
    logFile_ = fileName
    logFilePath_ = path & fileName
End Sub

Private Sub ILoggerOutput_OutputLog(ByVal timestamp As String, ByVal level As String, ByVal moduleName As String, ByVal message As String)
    ILoggerOutput_OutputLogMsg timestamp & " " & level & " " & IIf(moduleName <> "", moduleName & " ", "") & message
End Sub

Private Sub ILoggerOutput_OutputLogMsg(ByVal message As String)
    Dim fnum As Integer
    If Dir(logFilePath_) <> "" Then
        If FileLen(logFilePath_) >= maxLogSize_ Then
            Call SwitchNewLogFile
        End If
    Else
        ' ファイルがなく新規に作成する場合、ログの最初の時間を記録
        startLogTime_ = Format(Now, "yyyymmdd_hhmm")
    End If
    fnum = FreeFile
    Open logFilePath_ For Append As #fnum
    Print #fnum, message
    Close #fnum
End Sub

Private Sub ILoggerOutput_ClearLog(ByVal Target As String)
     
End Sub

Public Sub RenameLogToOld()
    Dim logFileName As String
    Dim oldLogFilePathName As String
    Dim pos As Long
    Dim availableLogFileNum As Integer
    pos = InStrRev(logFile_, ".")
    logFileName = Left(logFile_, pos - 1)     ' "."より前の部分を取得
    availableLogFileNum = FindLogNum(logFileName, logFileCount_)
    oldLogFilePathName = logPath_ & logFileName & "." & Format((availableLogFileNum + 1), "00") & ".log"       ' xxxxxx.nn.log形式にリネーム
On Error GoTo RenameError:
    Name logFilePath_ As oldLogFilePathName
On Error GoTo 0
    Exit Sub
RenameError:
    availableLogFileNum = availableLogFileNum + 1
    oldLogFilePathName = logPath_ & logFileName & "." & Format((availableLogFileNum + 1), "00") & ".log"
    Resume
End Sub

Public Sub RenameLogToTime()
    Dim logFileName As String
    Dim oldLogFilePathName As String
    Dim oldLogFileName As String
    Dim pos As Long
    Dim availableLogFileNum As Integer
    pos = InStrRev(logFile_, ".")
    logFileName = Left(logFile_, pos - 1)     ' "."より前の部分を取得
    oldLogFileName = logFileName & "-" & Format(Now, "yyyymmdd_hhmm") & ".log"
    oldLogFilePathName = logPath_ & oldLogFileName
On Error GoTo RenameError:
    Name logFilePath_ As oldLogFilePathName
On Error GoTo 0
    Exit Sub
RenameError:
    Dim logCount As Integer
    logCount = FindLogCount(logFileName, 1)
    oldLogFilePathName = logPath_ & logFileName & "-" & Format(Now, "yyyymmdd_hhmm") & "_" & Format((logCount), "00") & ".log"
    Resume
End Sub

Public Sub SwitchNewLogFile()
    If fileLogOutputMode_ = FileLogTimeBasedName Then
        Call RenameLogToTime
    ElseIf fileLogOutputMode_ = FileLogFixedLogName Then
        Call RenameLogToOld
    End If
End Sub

Public Function FindLogCount(oldLogFileName As String, logCount As Integer) As Integer
    Dim logCountEnabled As Boolean
    logCountEnabled = True
    Do
        Dim renameLogFilePathName As String
        renameLogFilePathName = logPath_ & oldLogFileName & "_" & Format(logCount, "00") & ".log"
        If Dir(renameLogFilePathName) = "" Then
            FindLogCount = logCount
            Exit Function
        End If
        logCount = logCount + 1
    Loop While logCountEnabled
End Function

Public Function FindLogNum(logFileName As String, logNum As Integer) As Integer
    Dim logNumEnabled As Boolean
    logNumEnabled = True
    Do
        Dim newLogFileName As String
        newLogFileName = logPath_ & logFileName & "." & Format((logNum + 1), "00") & ".log"
        If Dir(newLogFileName) = "" Then
            FindLogNum = logNum
            logFileCount_ = logNum - 1
            Exit Function
        End If
        logNum = logNum + 1
    Loop While logNumEnabled
End Function

Private Sub ChangeLogFileName(newLogFileName As String)
    logFile_ = newLogFileName
    logFilePath_ = logPath_ & newLogFileName
End Sub

Public Property Let maxLogSize(logSize As Long)
    maxLogSize_ = logSize
End Property

Public Property Get maxLogSize() As Long
    maxLogSize = maxLogSize_
End Property

Public Property Let ILoggerOutput_LogFormat(formatStr As String)
    logFormat_ = formatStr
End Property

Public Property Get ILoggerOutput_LogFormat() As String
    ILoggerOutput_LogFormat = logFormat_
End Property

Public Property Let LogFormat(formatStr As String)
    logFormat_ = formatStr
End Property


