VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iLoggerOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: ILoggerOutput
Option Explicit

Public Sub OutputLog(ByVal timestamp As String, ByVal level As String, ByVal moduleName As String, ByVal message As String)
    ' 実装は各出力クラスで行う
End Sub

Public Sub OutputLogMsg(ByVal message As String)
    ' 実装は各出力クラスで行う
End Sub

Public Sub ClearLog(ByVal target As String)
    ' 実装は各出力クラスで行う
End Sub

Public Property Let LogFormat(formatStr As String)

End Property

Public Property Get LogFormat() As String

End Property
