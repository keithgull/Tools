Attribute VB_Name = "Utils"
Option Explicit


' 共通のファイル参照処理
'  ワークシートの特定のセルに対してファイル選択を行い、選択されたファイルパスを設定します。
'   CommonFileRef
'    →SelectFileAndSetPath
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

' 共通のフォルダ参照処理
'  ワークシートの特定のセルに対してフォルダ選択を行い、選択されたフォルダパスを設定します。
'   CommonFolderRef
'    →SelectFolderAndSetPath
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
    
    ' ファイルダイアログの作成
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' ダイアログの設定
    dialog.Title = IIf(dialogTitle <> "", dialogTitle, "フォルダを選択してください")
    dialog.AllowMultiSelect = False
    dialog.InitialFileName = defaultpath
    cancelMsg = IIf(silentMode = False And cancelMsg <> "", cancelMsg, "フォルダ選択がキャンセルされました。")
    
    ' ダイアログを表示して選択
    If dialog.Show = -1 Then
        ' 選択されたフォルダパスを取得
        folderPath = dialog.SelectedItems(1)
        
        ' セルA1にフォルダパスを設定
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
    
    ' ファイルダイアログの作成
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    
    ' ダイアログの設定
    dialog.Filters.Clear
    dialog.Filters.Add fileType, fileFilter
    dialog.Title = IIf(dialogTitle <> "", dialogTitle, "ファイルを選択してください")
    dialog.AllowMultiSelect = False
    dialog.InitialFileName = defaultpath
    cancelMsg = IIf(silentMode = False And cancelMsg <> "", cancelMsg, "ファイル選択がキャンセルされました。")
    
    ' ダイアログを表示して選択
    If dialog.Show = -1 Then
        ' 選択されたフォルダパスを取得
        filePath = dialog.SelectedItems(1)
        
        ' セルA1にフォルダパスを設定
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

    ' 空白セルを除外して値を収集
    For Each cell In rng
        If Trim(cell.Value) <> "" Then
            cellList.Add Trim(cell.Value)
        End If
    Next cell
    
    ' コレクションを配列に変換
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

