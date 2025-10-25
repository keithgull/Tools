VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutorSheetName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsExecutorReplace
Option Explicit
Implements iExecutor

Private Type tpListColSheetNmTpl
    headerStRow As Long
    folderNmCol As Integer
    fileNameCol As Integer
    sheetNamCol As Integer
    shVisiblCol As Integer
    shRenameCol As Integer
End Type

Private tplCurRow As Long
Private udtLC As tpListColSheetNmTpl

Public Sub IExecutor_InitExecutor(ws As Worksheet)
    ' 列構造体の初期化
    udtLC = InitiListColsTpl(ws)

End Sub

Public Function IExecutor_ReadFile(ByRef ws As Worksheet, ByRef wb As Workbook, ByVal baseFolder As String, ByVal folderName As String, ByVal fileName As String, ByRef tplRow As Long, Optional log As clsLogger = Nothing) As Integer
    Dim sheetName As String
    Dim sh As Worksheet
    
    ws.Range("TPL_BASE_FOLDER").Value = baseFolder
            
    For Each sh In wb.Worksheets
        ws.Cells(tplRow, udtLC.folderNmCol).Value = folderName
        ws.Cells(tplRow, udtLC.fileNameCol).Value = fileName
        ws.Cells(tplRow, udtLC.sheetNamCol).Value = sh.Name
        ws.Cells(tplRow, udtLC.shVisiblCol).Value = IIf(Not sh.Visible, "◯", "")
        tplRow = tplRow + 1
    Next
    tplRow = tplRow - 1
    
End Function

Public Sub IExecutor_Execute(ByVal folderName As String, ByVal fileName As String, Optional log As clsLogger = Nothing)
    ' 実装は各出力クラスで行う
End Sub

Public Function IExecutor_GetTemplateSheetName() As String
    IExecutor_GetTemplateSheetName = "TemplateSheetList"
End Function


Private Function InitiListColsTpl(ws As Worksheet) As tpListColSheetNmTpl
    Dim listCol As tpListColSheetNmTpl
    listCol.headerStRow = ws.Range("TPL_HEADER_ROW").row
    listCol.folderNmCol = ws.Range("TPL_HEADER_FOLDER").Column
    listCol.fileNameCol = ws.Range("TPL_HEADER_FILENAME").Column
    listCol.sheetNamCol = ws.Range("TPL_HEADER_SHEETNAME").Column
    listCol.shVisiblCol = ws.Range("TPL_HEADER_SHEET_VISIBLE").Column
    listCol.shRenameCol = ws.Range("TPL_HEADER_SHEET_RENAME").Column
    InitiListColsTpl = listCol
End Function
