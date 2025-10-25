Attribute VB_Name = "MacroMain"
Option Explicit

Public Const SHEET_TOOL As String = "ツール"

Public Const EXCLUDE_SHEETNAMES As String = "ShConfig,ShToolMain,TemplateSh0,TemplateSh1,TemplateSh2"

' 名前の定義
Public Const RNG_FILE_LIST As String = "RANGE_FILE_LIST"

Type tpExecParams
    recurse As Boolean
    attrib  As Boolean
End Type

Type tpListCol
    folderNmCol As Integer
    fileNameCol As Integer
    fileSizeCol As Integer
    fileCrdtCol As Integer
    fileUpdtCol As Integer
End Type

Type tpExcludeList
    excludeFileNmArray() As String
    excludeFolderArray() As String
End Type

' クリアボタンを押したときの処理
Public Sub ClearFileList()
    Dim target As Range
    Set target = ThisWorkbook.Worksheets(SHEET_TOOL).Range(RNG_FILE_LIST)
    target.ClearContents
End Sub

' 参照ボタンを押したときの処理
Public Sub ReferenceFolder()
    Dim ws As Worksheet
    Dim ret As String
    
    Set ws = ThisWorkbook.ActiveSheet
    ret = SelectFolderAndSetPath("Excelブックを検索するフォルダを設定してください。", "", False)
    If ret <> "" Then
        ws.Range("TARGET_FOLDER").Value = ret
    End If
    
End Sub

' 「ファイル一覧取得」ボタンを押したときの処理
'  RetrieveFileList
'   →ListFiles
Public Sub RetrieveFileList()
    Dim baseFolder As String
    Dim ws As Worksheet
    Dim row As Long
    Dim recurse As Boolean
    Dim udtExePrm As tpExecParams
    Dim udtLstCol As tpListCol
    Dim udtExcLst As tpExcludeList
    
    ' ベースフォルダパスの取得 (セルC2を基準)
    baseFolder = Range("TARGET_FOLDER").Value
    
    ' ワークシート設定
    Set ws = ThisWorkbook.Sheets("ツール")
    
    ' 過去の一覧をクリア
    ws.Range("RANGE_FILE_LIST").ClearContents
    
    ' 表示を開始する行
    row = ws.Range("HEADER_FOLDER").row + 1
    
    ' 各パラメータ初期化
    udtExePrm = InitExecParams(ws)         ' 実行パラメータの初期化
    udtLstCol = InitListCols(ws)           ' リスト列構造体の初期化
    udtExcLst = InitExcludeList(ws)        ' 除外リストの初期化
    
    ' 再帰的にファイルを取得
    Call ListFiles(baseFolder, ws, row, udtExePrm, udtLstCol, udtExcLst)
End Sub

' リスト一覧を作成する処理
Sub ListFiles(folderPath As String, ws As Worksheet, ByRef row As Long, udtExePrm As tpExecParams, ByRef udtLstCol As tpListCol, udtExcLst As tpExcludeList)
    Dim fm As New clsFSOManager    ' FileSystemObject の作成
    Dim fol As Scripting.folder
    Dim subfolder As Object
    Dim fi As Object
           
    ' 指定フォルダの取得
    Set fol = fm.GetFolder(folderPath)
        
    ' フォルダ内のファイルをリストアップ
    For Each fi In fol.files
        If fi.Name Like "*.xls*" Then ' Excelファイルのみ
            If Not IsExclude(fi.Name, udtExcLst.excludeFileNmArray) Then
                ws.Cells(row, udtLstCol.folderNmCol).Value = Replace(fol.path & "\", folderPath, "")  ' 相対フォルダ
                ws.Cells(row, udtLstCol.fileNameCol).Value = fi.Name              ' ファイル名
                ws.Cells(row, udtLstCol.fileSizeCol).Value = fi.Size              ' ファイルサイズ
                ws.Cells(row, udtLstCol.fileCrdtCol).Value = fi.DateCreated       ' 作成日時
                ws.Cells(row, udtLstCol.fileUpdtCol).Value = fi.DateLastModified  ' 更新日時
                row = row + 1
            End If
        End If
    Next fi
    
    If udtExePrm.recurse Then
        ' サブフォルダも再帰的に検索
        For Each subfolder In fol.Subfolders
            If IsExclude(subfolder.Name, udtExcLst.excludeFolderArray) Then
                Call ListFiles(subfolder.path, ws, row, udtExePrm, udtLstCol, udtExcLst)
            End If
        Next subfolder
    End If
End Sub

Public Sub ReadListFiles()
    Dim ws As Worksheet
    Dim configWs As Worksheet
    Dim filecount As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim basePath As String
    Dim udtListCols As tpListCol
    Dim executor As iExecutor
    Dim execFunc As String
    Dim newShtNm As String
    Dim tmpSheet As String
    Dim msgRet As VbMsgBoxResult
    Dim newWs As Worksheet
    Dim currentRow As Long
    Dim loggerType As LOGGER_TYPE
    Dim logOutMode As LogOutputMode
    Dim param1 As String
    Dim param2 As String
    Dim log As clsLogger
    Dim excludeSheetArray() As String
    Dim tplRow As Long
    Dim tplStartRow As Long
    excludeSheetArray = Split(EXCLUDE_SHEETNAMES, ",")
    
    Set ws = ThisWorkbook.Worksheets("ツール")
    Set configWs = ThisWorkbook.Worksheets("設定")
    filecount = ws.Range("FILE_COUNT").Value
    
    startRow = ws.Range("HEADER_FOLDER").row + 1
    endRow = ws.Range("HEADER_FOLDER").row + filecount
    udtListCols = InitListCols(ws)
    basePath = ws.Range("TARGET_FOLDER").Value
    execFunc = ws.Range("EXEC_FUNCTION").Value
    newShtNm = ws.Range("NEWSHEETNAME").Value
    
    loggerType = GetLoggerType(configWs.Range("LOGGER_TYPE").Value)
    logOutMode = GetLogOutputMode(configWs.Range("LOG_OUTPUTMODE").Value)
    param1 = Replace(configWs.Range("LOGFILE_PATH").Value, "{WORKBOOK_PATH}", ThisWorkbook.path)
    param2 = configWs.Range("LOGFILE_NAME").Value
    
    If newShtNm = "" Then
        msgRet = MsgBox("シート名が設定されていません", vbOKOnly, "エラー")
        Exit Sub
    End If
    
    Set executor = GetExecutor(execFunc)
    tmpSheet = executor.GetTemplateSheetName()
    
    Set newWs = AddSheet(tmpSheet, newShtNm, excludeSheetArray)
    If newWs Is Nothing Then
        msgRet = MsgBox("シートが正しく追加できませんでした。", vbOKOnly, "エラー")
        Exit Sub
    End If
    
    ' executorの初期化
    Call executor.InitExecutor(newWs)
    tplStartRow = newWs.Range("TPL_HEADER_ROW").row + 1
    tplRow = tplStartRow
    
    Set log = InitLogger(loggerType, param1, param2, logOutMode, 100000, False)
    ' メインループ
    For currentRow = startRow To endRow
        Dim targetFilePath As String
        Dim targetFileName As String
        Dim targetFilePathName As String
        Dim newWB As Workbook
        Dim ret As Integer
        
        targetFilePath = ws.Cells(currentRow, udtListCols.folderNmCol).Value
        targetFileName = ws.Cells(currentRow, udtListCols.fileNameCol).Value
        targetFilePathName = basePath & targetFilePath & targetFileName
        
        ' ファイルオープン
        Set newWB = Workbooks.Open(fileName:=targetFilePathName, ReadOnly:=True)
        ret = executor.ReadFile(newWs, newWB, basePath, targetFilePath, targetFileName, tplRow, log)
        newWB.Close saveChanges:=False
        tplRow = tplRow + 1
    Next
    
End Sub

Function InitExecParams(ws As Worksheet) As tpExecParams
    Dim params As tpExecParams
    params.recurse = IIf(ws.Range("SUBFOLDER_ENABLED").Value = "はい", True, False)
    params.attrib = IIf(ws.Range("ATTRIBUTE_ENABLED").Value = "はい", True, False)
    InitExecParams = params
End Function

Function InitListCols(ws As Worksheet) As tpListCol
    Dim listCols As tpListCol
    listCols.folderNmCol = ws.Range("HEADER_FOLDER").Column
    listCols.fileNameCol = ws.Range("HEADER_FILENAME").Column
    listCols.fileSizeCol = ws.Range("HEADER_FILESIZE").Column
    listCols.fileUpdtCol = ws.Range("HEADER_CREATETIME").Column
    listCols.fileCrdtCol = ws.Range("HEADER_UPDATETIME").Column
    InitListCols = listCols
End Function

Function InitExcludeList(ws As Worksheet) As tpExcludeList
    Dim ecList As tpExcludeList
    ecList.excludeFolderArray = GetCellRangeToArray("EXCLUDE_FOLDER_LIST") '除外フォルダのリストを設定シートより取得
    ecList.excludeFileNmArray = GetCellRangeToArray("EXCLUDE_FILE_LIST")   '除外ファイル（拡張子）のリストを設定シートより取得
    InitExcludeList = ecList
End Function


Function GetExecutor(func As String) As iExecutor
    Dim executor As iExecutor
    Select Case func
        Case "属性変更"
            Set executor = New clsExecutorFileAttribute
        Case "シート名変更"
            Set executor = New clsExecutorSheetName
        Case "ファイル内容置換"
            Set executor = New clsExecutorGrepAndReplace
        Case Else
            Set GetExecutor = Nothing
    End Select
    Set GetExecutor = executor
End Function

' 除外フォルダおよび除外ファイル（拡張子）の処理
Function IsExclude(fileFolderName As String, excludePatterns() As String) As Boolean
    Dim i As Long
    
    ' 初期状態: 一致しないとする
    IsExclude = False

    For i = LBound(excludePatterns) To UBound(excludePatterns)
        ' ワイルドカードを活用し、一致判定を実施
        If fileFolderName Like excludePatterns(i) Then
            IsExclude = True ' 一致した場合
            Exit Function
        End If
    Next i

End Function


