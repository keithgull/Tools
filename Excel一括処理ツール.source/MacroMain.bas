Attribute VB_Name = "MacroMain"
Option Explicit

Public Const SHEET_TOOL As String = "ツール"
Public Const SHEET_CONFIG As String = "設定"

Public Const EXCLUDE_SHEETNAMES As String = "ShConfig,ShToolMain,TemplateSh0,TemplateSh1,TemplateSh2"

' 名前の定義
Public Const RNG_FILE_LIST As String = "RANGE_FILE_LIST"

Type tpExecParams
    recurse As Boolean
    append  As Boolean
End Type

Type tpListCol
    folderNmCol As Integer
    fileNameCol As Integer
    fileSizeCol As Integer
    fileCrdtCol As Integer
    fileUpdtCol As Integer
    fileChekCol As Integer
    checkFlgCol As Integer
End Type

Type tpExcludeList
    excludeFileNmArray() As String
    excludeFolderArray() As String
End Type

Public Function GetMainWs() As Worksheet
    Set GetMainWs = ThisWorkbook.Worksheets(SHEET_TOOL)
End Function

Public Function GetConfigWs() As Worksheet
    Set GetConfigWs = ThisWorkbook.Worksheets(SHEET_CONFIG)
End Function

' クリアボタンを押したときの処理
Public Sub ClearFileList()
    Call ClearData(GetMainWs(), RNG_FILE_LIST, True, "一覧をクリアしてもいいですか？")
End Sub

' 参照ボタンを押したときの処理
Public Sub ReferenceFolder()
    Call CommonFolderRef(GetMainWs(), "TARGET_FOLDER", ThisWorkbook.path, "Excelブックを検索するフォルダを設定してください。", "", False)
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
    Set ws = GetMainWs()
    
    ' 各パラメータ初期化
    udtExePrm = InitExecParams(ws)         ' 実行パラメータの初期化
    udtLstCol = InitListCols(ws)           ' リスト列構造体の初期化
    udtExcLst = InitExcludeList(ws)        ' 除外リストの初期化
        
    ' 表示を開始する行
    If udtExePrm.append = True Then
        row = ws.Range("HEADER_FOLDER").row + ws.Range("FILE_COUNT").Value + 1
    Else
        ' 過去の一覧をクリア
        ws.Range("RANGE_FILE_LIST").ClearContents
        row = ws.Range("HEADER_FOLDER").row + 1
    End If
    
    ' 再帰的にファイルを取得
    Call ListFiles(baseFolder, ws, row, udtExePrm, udtLstCol, udtExcLst)
End Sub

' リスト一覧を作成する処理
Sub ListFiles(folderPath As String, ws As Worksheet, ByRef row As Long, udtExePrm As tpExecParams, ByRef udtLstCol As tpListCol, udtExcLst As tpExcludeList)
    Dim fm As New clsFSOManager    ' FileSystemObject の作成
    Dim fol As Scripting.folder
    Dim subfolder As Object
    Dim fi As Object
    Dim basePath As String
           
    ' 指定フォルダの取得
    Set fol = fm.GetFolder(folderPath)
    basePath = ws.Range("TARGET_FOLDER").Value
        
    ' フォルダ内のファイルをリストアップ
    For Each fi In fol.files
        If fi.Name Like "*.xls*" Then ' Excelファイルのみ
            If Not IsExclude(fi.Name, udtExcLst.excludeFileNmArray) Then
                ws.Cells(row, udtLstCol.folderNmCol).Value = Replace(fol.path & "\", basePath, "")  ' 相対フォルダ
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
            If Not IsExclude(subfolder.Name, udtExcLst.excludeFolderArray) Then
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
    
    Set ws = GetMainWs()
    Set configWs = GetConfigWs()
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
    
    Call modifySheetButton(newWs, "ExecuteSheetTask")
    
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
        If ws.Cells(currentRow, udtListCols.checkFlgCol).Value Then
            Set newWB = Workbooks.Open(fileName:=targetFilePathName, ReadOnly:=True)
            ret = executor.ReadFile(newWs, newWB, basePath, targetFilePath, targetFileName, tplRow, log)
            newWB.Close saveChanges:=False
            tplRow = tplRow + 1
        End If
    Next
    
End Sub

Public Sub ExecuteSheetTask(Optional ByVal wsName As String = "")
    Dim mainWs As Worksheet
    Dim taskWs As Worksheet
    Dim btnName As String
    Dim executor As iExecutor
    Dim tplStartRow As Long
    Dim tplRow As Long
    Dim currentRow As Long
    
    Dim startRow As Long
    Dim endRow As Long
    
    
    Dim log As clsLogger
    Dim loggerType As LOGGER_TYPE
    Dim logOutMode As LogOutputMode
    Dim param1 As String
    Dim param2 As String
    Dim execRet As Integer
    
    Set mainWs = GetMainWs()
    If wsName <> "" Then
        Set taskWs = ThisWorkbook.Worksheets(wsName)
    Else
        Set taskWs = ActiveSheet
    End If
    
    Set executor = GetExecutor(mainWs.Range("EXEC_FUNCTION").Value)
    Call executor.InitExecutor(taskWs)
    tplStartRow = taskWs.Range("TPL_HEADER_ROW").row + 1
    tplRow = tplStartRow
    
    Set log = InitLogger(loggerType, param1, param2, logOutMode, 100000, False)
    
    executor.Execute taskWs, Application.caller, tplStartRow, "", "", log
    
    Debug.Print taskWs.Name
End Sub


Function InitExecParams(ws As Worksheet) As tpExecParams
    Dim params As tpExecParams
    params.recurse = IIf(ws.Range("SUBFOLDER_ENABLED").Value = "はい", True, False)
    params.append = IIf(ws.Range("APPEND_ENABLED").Value = "はい", True, False)
    InitExecParams = params
End Function

Function InitListCols(ws As Worksheet) As tpListCol
    Dim listCols As tpListCol
    listCols.folderNmCol = ws.Range("HEADER_FOLDER").Column
    listCols.fileNameCol = ws.Range("HEADER_FILENAME").Column
    listCols.fileSizeCol = ws.Range("HEADER_FILESIZE").Column
    listCols.fileUpdtCol = ws.Range("HEADER_CREATETIME").Column
    listCols.fileCrdtCol = ws.Range("HEADER_UPDATETIME").Column
    listCols.fileChekCol = ws.Range("HEADER_SELECT").Column
    listCols.checkFlgCol = ws.Range("HEADER_ALL_CHECK_FLG").Column
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


