Attribute VB_Name = "MacroMain"
Option Explicit

Public Const SHEET_TOOL As String = "�c�[��"
Public Const SHEET_CONFIG As String = "�ݒ�"

Public Const EXCLUDE_SHEETNAMES As String = "ShConfig,ShToolMain,TemplateSh0,TemplateSh1,TemplateSh2"

' ���O�̒�`
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

' �N���A�{�^�����������Ƃ��̏���
Public Sub ClearFileList()
    Call ClearData(GetMainWs(), RNG_FILE_LIST, True, "�ꗗ���N���A���Ă������ł����H")
End Sub

' �Q�ƃ{�^�����������Ƃ��̏���
Public Sub ReferenceFolder()
    Call CommonFolderRef(GetMainWs(), "TARGET_FOLDER", ThisWorkbook.path, "Excel�u�b�N����������t�H���_��ݒ肵�Ă��������B", "", False)
End Sub

' �u�t�@�C���ꗗ�擾�v�{�^�����������Ƃ��̏���
'  RetrieveFileList
'   ��ListFiles
Public Sub RetrieveFileList()
    Dim baseFolder As String
    Dim ws As Worksheet
    Dim row As Long
    Dim recurse As Boolean
    Dim udtExePrm As tpExecParams
    Dim udtLstCol As tpListCol
    Dim udtExcLst As tpExcludeList
    
    ' �x�[�X�t�H���_�p�X�̎擾 (�Z��C2���)
    baseFolder = Range("TARGET_FOLDER").Value
    
    ' ���[�N�V�[�g�ݒ�
    Set ws = GetMainWs()
    
    ' �e�p�����[�^������
    udtExePrm = InitExecParams(ws)         ' ���s�p�����[�^�̏�����
    udtLstCol = InitListCols(ws)           ' ���X�g��\���̂̏�����
    udtExcLst = InitExcludeList(ws)        ' ���O���X�g�̏�����
        
    ' �\�����J�n����s
    If udtExePrm.append = True Then
        row = ws.Range("HEADER_FOLDER").row + ws.Range("FILE_COUNT").Value + 1
    Else
        ' �ߋ��̈ꗗ���N���A
        ws.Range("RANGE_FILE_LIST").ClearContents
        row = ws.Range("HEADER_FOLDER").row + 1
    End If
    
    ' �ċA�I�Ƀt�@�C�����擾
    Call ListFiles(baseFolder, ws, row, udtExePrm, udtLstCol, udtExcLst)
End Sub

' ���X�g�ꗗ���쐬���鏈��
Sub ListFiles(folderPath As String, ws As Worksheet, ByRef row As Long, udtExePrm As tpExecParams, ByRef udtLstCol As tpListCol, udtExcLst As tpExcludeList)
    Dim fm As New clsFSOManager    ' FileSystemObject �̍쐬
    Dim fol As Scripting.folder
    Dim subfolder As Object
    Dim fi As Object
    Dim basePath As String
           
    ' �w��t�H���_�̎擾
    Set fol = fm.GetFolder(folderPath)
    basePath = ws.Range("TARGET_FOLDER").Value
        
    ' �t�H���_���̃t�@�C�������X�g�A�b�v
    For Each fi In fol.files
        If fi.Name Like "*.xls*" Then ' Excel�t�@�C���̂�
            If Not IsExclude(fi.Name, udtExcLst.excludeFileNmArray) Then
                ws.Cells(row, udtLstCol.folderNmCol).Value = Replace(fol.path & "\", basePath, "")  ' ���΃t�H���_
                ws.Cells(row, udtLstCol.fileNameCol).Value = fi.Name              ' �t�@�C����
                ws.Cells(row, udtLstCol.fileSizeCol).Value = fi.Size              ' �t�@�C���T�C�Y
                ws.Cells(row, udtLstCol.fileCrdtCol).Value = fi.DateCreated       ' �쐬����
                ws.Cells(row, udtLstCol.fileUpdtCol).Value = fi.DateLastModified  ' �X�V����
                row = row + 1
            End If
        End If
    Next fi
    
    If udtExePrm.recurse Then
        ' �T�u�t�H���_���ċA�I�Ɍ���
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
        msgRet = MsgBox("�V�[�g�����ݒ肳��Ă��܂���", vbOKOnly, "�G���[")
        Exit Sub
    End If
    
    Set executor = GetExecutor(execFunc)
    tmpSheet = executor.GetTemplateSheetName()
    
    Set newWs = AddSheet(tmpSheet, newShtNm, excludeSheetArray)
    If newWs Is Nothing Then
        msgRet = MsgBox("�V�[�g���������ǉ��ł��܂���ł����B", vbOKOnly, "�G���[")
        Exit Sub
    End If
    
    Call modifySheetButton(newWs, "ExecuteSheetTask")
    
    ' executor�̏�����
    Call executor.InitExecutor(newWs)
    tplStartRow = newWs.Range("TPL_HEADER_ROW").row + 1
    tplRow = tplStartRow
    
    Set log = InitLogger(loggerType, param1, param2, logOutMode, 100000, False)
    ' ���C�����[�v
    For currentRow = startRow To endRow
        Dim targetFilePath As String
        Dim targetFileName As String
        Dim targetFilePathName As String
        Dim newWB As Workbook
        Dim ret As Integer
        
        targetFilePath = ws.Cells(currentRow, udtListCols.folderNmCol).Value
        targetFileName = ws.Cells(currentRow, udtListCols.fileNameCol).Value
        targetFilePathName = basePath & targetFilePath & targetFileName
        
        ' �t�@�C���I�[�v��
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
    params.recurse = IIf(ws.Range("SUBFOLDER_ENABLED").Value = "�͂�", True, False)
    params.append = IIf(ws.Range("APPEND_ENABLED").Value = "�͂�", True, False)
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
    ecList.excludeFolderArray = GetCellRangeToArray("EXCLUDE_FOLDER_LIST") '���O�t�H���_�̃��X�g��ݒ�V�[�g���擾
    ecList.excludeFileNmArray = GetCellRangeToArray("EXCLUDE_FILE_LIST")   '���O�t�@�C���i�g���q�j�̃��X�g��ݒ�V�[�g���擾
    InitExcludeList = ecList
End Function


Function GetExecutor(func As String) As iExecutor
    Dim executor As iExecutor
    Select Case func
        Case "�����ύX"
            Set executor = New clsExecutorFileAttribute
        Case "�V�[�g���ύX"
            Set executor = New clsExecutorSheetName
        Case "�t�@�C�����e�u��"
            Set executor = New clsExecutorGrepAndReplace
        Case Else
            Set GetExecutor = Nothing
    End Select
    Set GetExecutor = executor
End Function

' ���O�t�H���_����я��O�t�@�C���i�g���q�j�̏���
Function IsExclude(fileFolderName As String, excludePatterns() As String) As Boolean
    Dim i As Long
    
    ' �������: ��v���Ȃ��Ƃ���
    IsExclude = False

    For i = LBound(excludePatterns) To UBound(excludePatterns)
        ' ���C���h�J�[�h�����p���A��v��������{
        If fileFolderName Like excludePatterns(i) Then
            IsExclude = True ' ��v�����ꍇ
            Exit Function
        End If
    Next i

End Function


