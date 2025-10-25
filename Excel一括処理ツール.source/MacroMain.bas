Attribute VB_Name = "MacroMain"
Option Explicit

Public Const SHEET_TOOL As String = "�c�[��"

Public Const EXCLUDE_SHEETNAMES As String = "ShConfig,ShToolMain,TemplateSh0,TemplateSh1,TemplateSh2"

' ���O�̒�`
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

' �N���A�{�^�����������Ƃ��̏���
Public Sub ClearFileList()
    Dim target As Range
    Set target = ThisWorkbook.Worksheets(SHEET_TOOL).Range(RNG_FILE_LIST)
    target.ClearContents
End Sub

' �Q�ƃ{�^�����������Ƃ��̏���
Public Sub ReferenceFolder()
    Dim ws As Worksheet
    Dim ret As String
    
    Set ws = ThisWorkbook.ActiveSheet
    ret = SelectFolderAndSetPath("Excel�u�b�N����������t�H���_��ݒ肵�Ă��������B", "", False)
    If ret <> "" Then
        ws.Range("TARGET_FOLDER").Value = ret
    End If
    
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
    Set ws = ThisWorkbook.Sheets("�c�[��")
    
    ' �ߋ��̈ꗗ���N���A
    ws.Range("RANGE_FILE_LIST").ClearContents
    
    ' �\�����J�n����s
    row = ws.Range("HEADER_FOLDER").row + 1
    
    ' �e�p�����[�^������
    udtExePrm = InitExecParams(ws)         ' ���s�p�����[�^�̏�����
    udtLstCol = InitListCols(ws)           ' ���X�g��\���̂̏�����
    udtExcLst = InitExcludeList(ws)        ' ���O���X�g�̏�����
    
    ' �ċA�I�Ƀt�@�C�����擾
    Call ListFiles(baseFolder, ws, row, udtExePrm, udtLstCol, udtExcLst)
End Sub

' ���X�g�ꗗ���쐬���鏈��
Sub ListFiles(folderPath As String, ws As Worksheet, ByRef row As Long, udtExePrm As tpExecParams, ByRef udtLstCol As tpListCol, udtExcLst As tpExcludeList)
    Dim fm As New clsFSOManager    ' FileSystemObject �̍쐬
    Dim fol As Scripting.folder
    Dim subfolder As Object
    Dim fi As Object
           
    ' �w��t�H���_�̎擾
    Set fol = fm.GetFolder(folderPath)
        
    ' �t�H���_���̃t�@�C�������X�g�A�b�v
    For Each fi In fol.files
        If fi.Name Like "*.xls*" Then ' Excel�t�@�C���̂�
            If Not IsExclude(fi.Name, udtExcLst.excludeFileNmArray) Then
                ws.Cells(row, udtLstCol.folderNmCol).Value = Replace(fol.path & "\", folderPath, "")  ' ���΃t�H���_
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
    
    Set ws = ThisWorkbook.Worksheets("�c�[��")
    Set configWs = ThisWorkbook.Worksheets("�ݒ�")
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
        Set newWB = Workbooks.Open(fileName:=targetFilePathName, ReadOnly:=True)
        ret = executor.ReadFile(newWs, newWB, basePath, targetFilePath, targetFileName, tplRow, log)
        newWB.Close saveChanges:=False
        tplRow = tplRow + 1
    Next
    
End Sub

Function InitExecParams(ws As Worksheet) As tpExecParams
    Dim params As tpExecParams
    params.recurse = IIf(ws.Range("SUBFOLDER_ENABLED").Value = "�͂�", True, False)
    params.attrib = IIf(ws.Range("ATTRIBUTE_ENABLED").Value = "�͂�", True, False)
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


