VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutorGrepAndReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsExecutorGrepAndReplace
Option Explicit
Implements iExecutor

Public Sub IExecutor_InitExecutor(ws As Worksheet)
    ' 列構造体の初期化
    'udtLC = InitiListColsTpl()

End Sub

Public Function IExecutor_ReadFile(ByRef ws As Worksheet, ByRef wb As Workbook, ByVal baseFolder As String, ByVal folderName As String, ByVal fileName As String, ByRef tplRow As Long, Optional log As clsLogger = Nothing) As Integer
    
End Function

Public Sub IExecutor_Execute(ByRef ws As Worksheet, ByVal caller As String, ByRef curRow As Integer, ByVal folderName As String, ByVal fileName As String, Optional log As clsLogger = Nothing)
    ' 実装は各出力クラスで行う
End Sub

Public Function IExecutor_GetTemplateSheetName() As String

End Function

