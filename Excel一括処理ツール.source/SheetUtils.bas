Attribute VB_Name = "SheetUtils"
Option Explicit

Public Function ClearData(ByRef ws As Worksheet, ByVal rangeName As String, Optional useMessage As Boolean, Optional message As String = "データをクリアしますか？") As Boolean
    Dim msgRet As VbMsgBoxResult
    If (useMessage) Then
        msgRet = MsgBox(message, vbYesNo, "クリアの確認")
        If (msgRet = vbNo) Then
            ClearData = False
            Exit Function
        End If
    End If
   
    ws.Range(rangeName).ClearContents
    ClearData = True
End Function

Function AddSheet(templateSheetName As String, newSheetName As String, excludeSheetArray() As String) As Worksheet
    Dim ws As Worksheet
    Dim template As Worksheet
    Dim i As Integer
    
    ' テンプレートシートの取得
    Set template = GetWorksheetByName(templateSheetName, False)
    
    If template Is Nothing Then
        MsgBox "テンプレートシートが見つかりません。", vbExclamation
        Set AddSheet = Nothing
        Exit Function
    End If

    ' 新しいシート名が既に存在するか確認
    If WorksheetExists(newSheetName) Then
        MsgBox "シートが既に存在するためスキップします。", vbExclamation
        Set AddSheet = Nothing
        Exit Function
    End If
    
    ' テンプレートシートをコピーして新しいシートを追加
    template.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
    'Set ws = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set ws = ThisWorkbook.ActiveSheet
    
    Debug.Print "add:" & ws.Name
    For i = 1 To ThisWorkbook.Worksheets.count
        'If Worksheets(i).CodeName = "shTemplate1" Then
        If IsExcludeSheet(excludeSheetArray, Worksheets(i).CodeName) Then
            Set ws = Worksheets(i)
            Exit For
        End If
    Next
    ws.Name = newSheetName
    
    ' 追加されたシートを戻り値として返す
    Set AddSheet = ws
End Function

Private Function IsExcludeSheet(excludeSheetArray() As String, sheetName As String) As Boolean
    Dim sh As Worksheet
    Dim shName As String
    Dim i As Integer
    For i = 0 To UBound(excludeSheetArray)
        shName = excludeSheetArray(i)
        If sheetName = shName Then
            IsExcludeSheet = False
            Exit Function
        End If
    Next
    IsExcludeSheet = True
End Function

Sub DeleteSheet(targetSheetName As String)
    Dim ws As Worksheet
    
    ' 対象シートの取得
    Set ws = GetWorksheetByName(targetSheetName, False)
    
    If ws Is Nothing Then
        'MsgBox "対象シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Application.DisplayAlerts = False
    Debug.Print "del:" & ws.Name
    ws.Delete
    
    Application.DisplayAlerts = True
End Sub

Function GetWorksheetByName(sheetName As String, silentMode As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "対象シートが見つかりません。", vbExclamation
        Exit Function
    End If
    Set GetWorksheetByName = ws
End Function

Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    WorksheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function

' セルをチェックボックスとした際に、全チェック処理・全解除を行うための処理
'
' 引数：
'       ws                      : ワークシート
'       maxCount                : データの処理件数
'       checkHeaderRngName      : チェックボックス列とした列のヘッダ行の名前
'       dataExistCheckRngName   : データの存在チェックを行う列の名前
Public Sub CheckAll(ByRef ws As Worksheet, val As String, maxCount As Long, checkHeaderRngName As String, dataExistCheckRngName As String, useFilter As Boolean)
Application.EnableEvents = False
Application.ScreenUpdating = False
    Dim i As Integer
    Dim startRow As Integer
    Dim checkCol As Integer
    Dim fileCol As Integer
    
    startRow = ws.Range(checkHeaderRngName).row + 1
    checkCol = ws.Range(checkHeaderRngName).Column
    fileCol = ws.Range(dataExistCheckRngName).Column
    
    Dim curRow As Integer
    For i = 0 To maxCount
        curRow = startRow + i
        If useFilter Then
            If ws.Cells(curRow, fileCol).EntireRow.Hidden = False Then
                If ws.Cells(curRow, fileCol).Value <> "" Then
                    ws.Cells(curRow, checkCol).Value = val
                End If
            End If
        Else
            If ws.Cells(curRow, fileCol).Value <> "" Then
                ws.Cells(curRow, checkCol).Value = val
            End If
        End If
    Next
Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub

' ボタンが押されたワークシートを特定し、ワークシートを返す
Function GetButtonParentSheet(buttonName As String) As Worksheet
    Dim ws As Worksheet
    Dim shp As Shape
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set shp = ws.Shapes(buttonName)
        If Err.Number = 0 Then
            Set GetButtonParentSheet = ws
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    Next
    Set GetButtonParentSheet = Nothing ' 見つからなかった場合
End Function


Function CommonDoubleClickAndCheck(ByRef ws As Worksheet, ByVal rngCheckHeaderName As String, dataCount As Long, rngAllChecked As String, ByVal checkboxRange As String, ByVal dataCheckHeaderRng As String, _
                                   ByVal checkStr As String, ByVal target As Range, cancel As Boolean, useFilter As Boolean)
    Dim targetRow As Long
    Dim targetCol As Integer
    Dim headerRow As Long
    Dim checkCol As Integer
    Dim allChecked As Boolean
    Dim maxRow As Long
    
    headerRow = ws.Range(rngCheckHeaderName).row
    checkCol = ws.Range(rngCheckHeaderName).Column
    allChecked = ws.Range(rngAllChecked).Value
    maxRow = headerRow + dataCount
    
    targetRow = target.row
    targetCol = target.Column
    If Not Intersect(target, ws.Range(checkboxRange)) Is Nothing Then
        If target.Value = checkStr Then
            target.Value = ""
        Else
            target.Value = checkStr ' check を表示
        End If
        cancel = True
    End If
        
    ' ヘッダの領域でダブルクリックが行われた場合
    If targetRow = headerRow And targetCol = checkCol Then
        If (Not allChecked) Then
            Call CheckAll(ws, checkStr, maxRow, rngCheckHeaderName, dataCheckHeaderRng, useFilter)
            ws.Range(rngAllChecked).Value = True
        Else
            Call CheckAll(ws, "", maxRow, rngCheckHeaderName, dataCheckHeaderRng, useFilter)
            ws.Range(rngAllChecked).Value = False
        End If
        cancel = True
    End If
    
End Function


