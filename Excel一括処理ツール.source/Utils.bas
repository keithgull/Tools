Attribute VB_Name = "Utils"
Option Explicit

Public Function SelectFolderAndSetPath(defaultPath As String, dialogTitle As String, cancelMsg As String, Optional silentmode As Boolean = True) As String
    Dim folderPath As String
    Dim dialog As FileDialog
    
    ' ファイルダイアログの作成
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' ダイアログの設定
    dialog.Title = IIf(dialogTitle <> "", dialogTitle, "フォルダを選択してください")
    dialog.AllowMultiSelect = False
    dialog.InitialFileName = defaultPath
    cancelMsg = IIf(silentmode = False And cancelMsg <> "", cancelMsg, "フォルダ選択がキャンセルされました。")
    
    ' ダイアログを表示して選択
    If dialog.Show = -1 Then
        ' 選択されたフォルダパスを取得
        folderPath = dialog.SelectedItems(1)
        
        ' セルA1にフォルダパスを設定
        SelectFolderAndSetPath = folderPath
        Exit Function
    Else
        If silentmode = False Then
            MsgBox cancelMsg, vbExclamation
        End If
    End If
    SelectFolderAndSetPath = ""
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

