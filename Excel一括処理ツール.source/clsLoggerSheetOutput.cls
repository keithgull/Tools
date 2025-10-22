VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLoggerSheetOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: SheetOutput
Option Explicit
Implements iLoggerOutput

Private targetSheet_ As Worksheet
Private currentRow_ As Long
Private currentCol_ As Long
Private lgLevelCol_ As Long
Private mdlnameCol_ As Long
Private messageCol_ As Long
Private currentIdx_ As Integer
Private lgLevelIdx_ As Integer
Private mdlnameIdx_ As Integer
Private messageIdx_ As Integer
Private logFormat_ As String
Private logTimeFormat_ As String
Private useModuleName_ As Boolean
Private sheetLogOutputMode_  As LogOutputMode
Private sheetMaxRowSize_ As Long
Private enableClearLogOnInit_ As Boolean
Private targetLogArea_ As Range

Public Sub Init(ByVal ws As Worksheet, ByVal logStartRow As Long, ByVal logStartCol As Long, sheetLogOutputMode As LogOutputMode, Optional maxLogSize As Long = 10000000, Optional useModuleName As Boolean = False)
    Set targetSheet_ = ws
    currentRow_ = logStartRow
    sheetMaxRowSize_ = maxLogSize
    sheetLogOutputMode_ = sheetLogOutputMode
    useModuleName_ = useModuleName
    currentIdx_ = -1
    lgLevelIdx_ = -1
    mdlnameIdx_ = -1
    messageIdx_ = -1
    
    Set targetLogArea_ = ws.Range(ws.Cells(logStartRow, logStartCol), ws.Cells(logStartRow + maxLogSize, logStartCol + 4))
    
    If sheetLogOutputMode = SheetLogSimple Then
        currentCol_ = logStartCol
    ElseIf sheetLogOutputMode = SheetLogFormatted Then
        Call initSheetCol(logStartCol)
        ws.Range(ws.Cells(logStartRow, currentCol_), ws.Cells(logStartRow + maxLogSize, currentCol_)).NumberFormatLocal = logTimeFormat_
    End If
    If enableClearLogOnInit_ Then
        targetLogArea_.ClearContents
    End If
End Sub

Private Sub initSheetCol(logStartCol As Long)
    Dim logArray() As String
    Dim i As Integer
    If Not useModuleName_ Then
        logFormat_ = Replace(logFormat_, "{module}", "")
    End If
    logArray = Split(logFormat_, "\t")
    For i = 0 To UBound(logArray)
        If InStr(logArray(i), "{time}") > 0 Then
            currentCol_ = logStartCol + i
            currentIdx_ = i
        End If
        If InStr(logArray(i), "{level}") > 0 Then
            lgLevelCol_ = logStartCol + i
            lgLevelIdx_ = i
        End If
        If useModuleName_ And InStr(logArray(i), "{module}") > 0 Then
            mdlnameCol_ = logStartCol + i
            mdlnameIdx_ = i
        End If
        If InStr(logArray(i), "{message}") > 0 Then
            messageCol_ = logStartCol + i
            messageIdx_ = i
        End If
    Next
End Sub


Private Sub ILoggerOutput_OutputLog(ByVal timestamp As String, ByVal level As String, ByVal moduleName As String, ByVal message As String)
    If sheetLogOutputMode_ = SheetLogSimple Then
        targetSheet_.Cells(currentRow_, currentCol_).Value = timestamp & " " & level & " " & IIf(moduleName <> "", moduleName & " ", "") & message
    ElseIf sheetLogOutputMode_ = SheetLogFormatted Then

    End If
End Sub

Private Sub ILoggerOutput_OutputLogMsg(ByVal message As String)
    
    If sheetLogOutputMode_ = SheetLogSimple Then
        targetSheet_.Cells(currentRow_, currentCol_).Value = Replace(message, "\t", " ")
    ElseIf sheetLogOutputMode_ = SheetLogFormatted Then
        Dim logArray() As String
        logArray = Split(message, "\t")
        If currentIdx_ >= 0 Then targetSheet_.Cells(currentRow_, currentCol_).Value = logArray(currentIdx_)
        If lgLevelIdx_ >= 0 Then targetSheet_.Cells(currentRow_, lgLevelCol_).Value = logArray(lgLevelIdx_)
        If mdlnameIdx_ >= 0 Then targetSheet_.Cells(currentRow_, mdlnameCol_).Value = logArray(mdlnameIdx_)
        If messageIdx_ >= 0 Then targetSheet_.Cells(currentRow_, messageCol_).Value = logArray(messageIdx_)
        
    End If
    currentRow_ = currentRow_ + 1
End Sub

Private Sub ILoggerOutput_ClearLog(ByVal target As String)
    targetSheet_.Range(targetLogArea_).ClearContents
End Sub

Public Property Let ILoggerOutput_LogFormat(formatStr As String)
    
End Property

Public Property Get ILoggerOutput_LogFormat() As String
    ILoggerOutput_LogFormat = logFormat_
End Property

Public Property Let LogFormat(formatStr As String)
    logFormat_ = formatStr
End Property

Public Property Let TimeFormat(formatStr As String)
    logTimeFormat_ = formatStr
End Property

Public Property Let LogOutputMode(sLogOutputMode As LogOutputMode)
    sheetLogOutputMode_ = sLogOutputMode
End Property

Public Property Get LogOutputMode() As LogOutputMode
    LogOutputMode = sheetLogOutputMode_
End Property

Public Property Let EnableClearLogOnInit(flag As Boolean)
    EnableClearLogOnInit = flag
End Property
