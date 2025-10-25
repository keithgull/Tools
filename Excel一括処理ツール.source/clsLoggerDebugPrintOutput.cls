VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLoggerDebugPrintOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: DebugOutput
Option Explicit
Implements iLoggerOutput

Private logFormat_ As String

Private Sub ILoggerOutput_OutputLog(ByVal timestamp As String, ByVal level As String, ByVal moduleName As String, ByVal message As String)
    ILoggerOutput_OutputLogMsg timestamp & " " & level & " " & IIf(moduleName <> "", moduleName & " ", "") & message
End Sub

Private Sub ILoggerOutput_OutputLogMsg(ByVal message As String)
    Debug.Print message
End Sub

Private Sub ILoggerOutput_ClearLog(ByVal Target As String)
     Debug.Print String(200, vbCrLf)
     Debug.Print String(200, Chr(8))
End Sub

Private Property Let ILoggerOutput_LogFormat(formatStr As String)
    logFormat_ = formatStr
End Property

Private Property Get ILoggerOutput_LogFormat() As String
    ILoggerOutput_LogFormat = logFormat_
End Property


Public Property Let LogFormat(formatStr As String)
    logFormat_ = formatStr
End Property
