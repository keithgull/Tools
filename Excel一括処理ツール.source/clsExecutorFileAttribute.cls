VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecutorFileAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsExecutorFileAttribute
Option Explicit
Implements iExecutor

Private Type tpListColAttrTpl
    headerStRow As Long
    folderNmCol As Integer
    fileNameCol As Integer
    sheetNamCol As Integer
    LastAuthCol As Integer
End Type

Private tplCurRow As Long
Private udtLC As tpListColAttrTpl


Public Sub IExecutor_InitExecutor(ws As Worksheet)
    ' óÒç\ë¢ëÃÇÃèâä˙âª
    udtLC = InitiListColsTpl(ws)

End Sub

Public Function IExecutor_ReadFile(ByRef ws As Worksheet, ByRef wb As Workbook, ByVal baseFolder As String, ByVal folderName As String, ByVal fileName As String, ByRef tplRow As Long, Optional log As clsLogger = Nothing) As Integer
    Dim lastAuthor As String
    Dim bdp As DocumentProperties

    Set bdp = wb.BuiltinDocumentProperties
    lastAuthor = bdp.Item("Last Author")
    
    ws.Range("TPL_BASE_FOLDER").Value = baseFolder
        
    ws.Cells(tplRow, udtLC.folderNmCol).Value = folderName
    ws.Cells(tplRow, udtLC.fileNameCol).Value = fileName
    ws.Cells(tplRow, udtLC.LastAuthCol).Value = lastAuthor
    
End Function

Public Sub IExecutor_Execute(ByRef ws As Worksheet, ByVal caller As String, ByRef curRow As Long, ByVal folderName As String, ByVal fileName As String, Optional log As clsLogger = Nothing)
    
End Sub


Public Function IExecutor_GetTemplateSheetName() As String
    IExecutor_GetTemplateSheetName = "TemplateFileAttrList"
End Function

Private Function InitiListColsTpl(ws As Worksheet) As tpListColAttrTpl
    Dim listCol As tpListColAttrTpl
    listCol.headerStRow = ws.Range("TPL_HEADER_ROW").row
    listCol.folderNmCol = ws.Range("TPL_HEADER_FOLDER").Column
    listCol.fileNameCol = ws.Range("TPL_HEADER_FILENAME").Column
    listCol.LastAuthCol = ws.Range("TPL_HEADER_LASTAUTHOR").Column
    InitiListColsTpl = listCol
End Function

