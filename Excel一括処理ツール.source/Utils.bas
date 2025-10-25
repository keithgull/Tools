Attribute VB_Name = "Utils"
Option Explicit


' ���ʂ̃t�@�C���Q�Ə���
'  ���[�N�V�[�g�̓���̃Z���ɑ΂��ăt�@�C���I�����s���A�I�����ꂽ�t�@�C���p�X��ݒ肵�܂��B
'   CommonFileRef
'    ��SelectFileAndSetPath
'
Public Function CommonFileRef(ws As Worksheet, rngName As String, defaultpath As String, fileType As String, fileFilter As String, dialogTitle As String, cancelMsg As String, Optional silentMode As Boolean = True) As String
    Dim ret As String
    Dim defPath As String
    Dim rngTarget As Range

    Set rngTarget = ws.Range(rngName)
    defPath = rngTarget.Value
    If defPath = "" Then
        defPath = defaultpath
    End If
    ret = SelectFileAndSetPath(defPath, fileType, fileFilter, dialogTitle, cancelMsg, silentMode)
    If ret <> "" Then
        rngTarget.Value = ret
    End If

    Set rngTarget = Nothing
    CommonFileRef = ret
End Function

' ���ʂ̃t�H���_�Q�Ə���
'  ���[�N�V�[�g�̓���̃Z���ɑ΂��ăt�H���_�I�����s���A�I�����ꂽ�t�H���_�p�X��ݒ肵�܂��B
'   CommonFolderRef
'    ��SelectFolderAndSetPath
'
Public Function CommonFolderRef(ws As Worksheet, rngName As String, defaultpath As String, dialogTitle As String, cancelMsg As String, Optional silentMode As Boolean = True) As String
    Dim ret As String
    Dim defPath As String
    Dim rngTarget As Range

    Set rngTarget = ws.Range(rngName)
    
    defPath = rngTarget.Value
    If defPath = "" Then
        defPath = defaultpath
    End If
    ret = SelectFolderAndSetPath(defPath, dialogTitle, cancelMsg, silentMode)
    If ret <> "" Then
        rngTarget.Value = ret
    End If
    
    Set rngTarget = Nothing
    CommonFolderRef = ret
End Function


Public Function SelectFolderAndSetPath(defaultpath As String, dialogTitle As String, cancelMsg As String, Optional silentMode As Boolean = True) As String
    Dim folderPath As String
    Dim dialog As FileDialog
    
    ' �t�@�C���_�C�A���O�̍쐬
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' �_�C�A���O�̐ݒ�
    dialog.Title = IIf(dialogTitle <> "", dialogTitle, "�t�H���_��I�����Ă�������")
    dialog.AllowMultiSelect = False
    dialog.InitialFileName = defaultpath
    cancelMsg = IIf(silentMode = False And cancelMsg <> "", cancelMsg, "�t�H���_�I�����L�����Z������܂����B")
    
    ' �_�C�A���O��\�����đI��
    If dialog.Show = -1 Then
        ' �I�����ꂽ�t�H���_�p�X���擾
        folderPath = dialog.SelectedItems(1)
        
        ' �Z��A1�Ƀt�H���_�p�X��ݒ�
        SelectFolderAndSetPath = folderPath
        Exit Function
    Else
        If silentMode = False Then
            MsgBox cancelMsg, vbExclamation
        End If
    End If
    SelectFolderAndSetPath = ""
End Function

Public Function SelectFileAndSetPath(defaultpath As String, fileType As String, fileFilter As String, dialogTitle As String, cancelMsg As String, Optional silentMode As Boolean = True) As String
    Dim filePath As String
    Dim dialog As FileDialog
    
    ' �t�@�C���_�C�A���O�̍쐬
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    
    ' �_�C�A���O�̐ݒ�
    dialog.Filters.Clear
    dialog.Filters.Add fileType, fileFilter
    dialog.Title = IIf(dialogTitle <> "", dialogTitle, "�t�@�C����I�����Ă�������")
    dialog.AllowMultiSelect = False
    dialog.InitialFileName = defaultpath
    cancelMsg = IIf(silentMode = False And cancelMsg <> "", cancelMsg, "�t�@�C���I�����L�����Z������܂����B")
    
    ' �_�C�A���O��\�����đI��
    If dialog.Show = -1 Then
        ' �I�����ꂽ�t�H���_�p�X���擾
        filePath = dialog.SelectedItems(1)
        
        ' �Z��A1�Ƀt�H���_�p�X��ݒ�
        SelectFileAndSetPath = filePath
        Exit Function
    Else
        If silentMode = False Then
            MsgBox cancelMsg, vbExclamation
        End If
    End If
    SelectFileAndSetPath = ""
End Function


Function GetCellRangeToArray(rangeName As String) As String()
    Dim rng As Range
    Dim cell As Range
    Dim cellList As Collection
    Dim cellArray() As String
    Dim i As Long
    
    On Error Resume Next
    Set rng = ThisWorkbook.Names(rangeName).RefersToRange
    On Error GoTo 0
    
    Set cellList = New Collection

    ' �󔒃Z�������O���Ēl�����W
    For Each cell In rng
        If Trim(cell.Value) <> "" Then
            cellList.Add Trim(cell.Value)
        End If
    Next cell
    
    ' �R���N�V������z��ɕϊ�
    ReDim cellArray(0 To cellList.count - 1)
    For i = 1 To cellList.count
        cellArray(i - 1) = cellList(i)
    Next i
    
    GetCellRangeToArray = cellArray
End Function


Function AddPathLastDelimiter(path As String) As String
    If Right(path, 1) <> "\" Then
        path = path & "\"
    End If
    AddPathLastDelimiter = path
End Function

Sub ActivateApp(val As Boolean)
    Application.ScreenUpdating = val
    Application.EnableEvents = val
    Application.Calculation = IIf(val, xlCalculationAutomatic, xlCalculationManual)
End Sub

