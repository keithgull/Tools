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

Enum SearchValueType
    VALUETYPE_VALUE = 0
    VALUETYPE_FORMULA = 1
    VALUETYPE_COL_WIDTH = 2
    VALUETYPE_ROW_HEIGHT = 3
    VALUETYPE_OFFSET_VALUE = 4
    VALUETYPE_OFFSET_FORMULA = 5
End Enum

Private Type tpSearchConditionTpl
    SrCondNum As String
    STVColPos As Integer
    STVRowPos As Long
    CelRowPos As Long
    GetVlType As Integer
    CellPosNm As String
End Type

Private Type tpListColSheetNmTpl
    headerStRow As Long
    folderNmCol As Integer
    fileNameCol As Integer
    sheetNamCol As Integer
    shVisiblCol As Integer
    shRenameCol As Integer
    checkFlgCol As Integer
End Type

Private tplCurRow As Long
Private udtLC As tpListColSheetNmTpl
Private searchCondAry() As tpSearchConditionTpl

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

Public Sub IExecutor_Execute(ByRef ws As Worksheet, ByVal caller As String, ByRef curRow As Long, ByVal folderName As String, ByVal fileName As String, Optional log As clsLogger = Nothing)
    Select Case caller
        Case "ExecuteReadFile"
            Call ReadFilePropAndValue(ws, caller, curRow, log)
        Case "ExecuteReflect"
    End Select
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
    listCol.checkFlgCol = ws.Range("TPL_ALL_CHECK_FLG").Column
    InitiListColsTpl = listCol
End Function

Private Function InitSearchValueTypeArray(ws As Worksheet, maxCnt As Integer) As tpSearchConditionTpl()
    Dim posCol As Integer
    Dim i As Integer
    Dim posRow As Long
    ReDim searchCondAry(maxCnt)
    
    For i = 0 To maxCnt
        Dim srchCond As tpSearchConditionTpl
        Dim tmpName As String
        tmpName = "TPL_SEARCH_COND_" + Format(i + 1, "00")
        srchCond.SrCondNum = "[" + Format(i + 1, "00") + "]"
        srchCond.STVColPos = ws.Range(tmpName).Column
        srchCond.STVRowPos = ws.Range("TPL_SEARCH_VALUE_TYPE_ROW").row
        srchCond.CelRowPos = ws.Range("TPL_SEARCH_VALUE_POS_ROW").row
        srchCond.GetVlType = ConvSTV(ws.Cells(srchCond.STVRowPos, srchCond.STVColPos).Value)
        srchCond.CellPosNm = ws.Cells(srchCond.CelRowPos, srchCond.STVColPos).Value
        searchCondAry(i) = srchCond
    Next
End Function

Private Function ConvSTV(stvStr As String) As SearchValueType
    Select Case stvStr
        Case "セル値"
            ConvSTV = VALUETYPE_VALUE
        Case "数式"
            ConvSTV = VALUETYPE_FORMULA
        Case "列幅"
            ConvSTV = VALUETYPE_COL_WIDTH
        Case "行の高さ"
            ConvSTV = VALUETYPE_ROW_HEIGHT
    End Select
End Function


Public Sub ReadFilePropAndValue(ByRef ws As Worksheet, ByVal caller As String, ByRef curRow As Long, Optional log As clsLogger = Nothing)
    Dim startRow As Long
    Dim endRow As Long
    Dim maxCount As Integer
    Dim currentRow As Long
    Dim basePath As String
    Dim tplRow As Long
    
    maxCount = 5
    Call InitSearchValueTypeArray(ws, maxCount)
    
    startRow = ws.Range("TPL_HEADER_FOLDER").row + 1
    endRow = ws.Range("TPL_HEADER_FOLDER").row + ws.Range("AVAILABLE_DATA_COUNT").Value
    basePath = ws.Range("BASEDIR").Value
    
    ' メインループ
    For tplRow = startRow To endRow
        Dim targetFilePath As String
        Dim targetFileName As String
        Dim targetFilePathName As String
        Dim targetSheetName As String
        Dim targetEnabled As Boolean
        Dim tmpWB As Workbook
        Dim ret As Integer
        
        targetEnabled = ws.Cells(tplRow, udtLC.checkFlgCol).Value
        
        If targetEnabled Then
            targetFilePath = ws.Cells(tplRow, udtLC.folderNmCol).Value
            targetFileName = ws.Cells(tplRow, udtLC.fileNameCol).Value
            targetSheetName = ws.Cells(tplRow, udtLC.sheetNamCol).Value
            targetFilePathName = basePath & targetFilePath & targetFileName
        
            ' ファイルオープン
            If ws.Cells(tplRow, udtLC.checkFlgCol).Value Then
                Dim retVal As String
                Set tmpWB = Workbooks.Open(fileName:=targetFilePathName, ReadOnly:=True)
                retVal = GetCellValue(ws, tmpWB.Worksheets(targetSheetName), maxCount, tplRow)
                tmpWB.Close saveChanges:=False
                tplRow = tplRow + 1
            End If
        End If
    Next
    
End Sub

Function GetCellValue(ws As Worksheet, tmpWs As Worksheet, maxCnt As Integer, tplRow As Long) As String
    Dim i As Integer
    For i = 0 To maxCnt
        If searchCondAry(i).CellPosNm <> "" Then
            Select Case (searchCondAry(i).GetVlType)
                Case VALUETYPE_VALUE
                    ws.Cells(tplRow, searchCondAry(i).STVColPos).Value = tmpWs.Range(searchCondAry(i).CellPosNm).Value
                Case VALUETYPE_FORMULA
                    ws.Cells(tplRow, searchCondAry(i).STVColPos).Value = "'" & tmpWs.Range(searchCondAry(i).CellPosNm).Formula
                Case VALUETYPE_COL_WIDTH
                    ws.Cells(tplRow, searchCondAry(i).STVColPos).Value = tmpWs.Range(searchCondAry(i).CellPosNm).ColumnWidth
                Case VALUETYPE_ROW_HEIGHT
                    ws.Cells(tplRow, searchCondAry(i).STVColPos).Value = tmpWs.Range(searchCondAry(i).CellPosNm).RowHeight
            End Select
        End If
    Next
End Function
