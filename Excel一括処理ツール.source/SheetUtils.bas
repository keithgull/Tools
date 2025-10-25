Attribute VB_Name = "SheetUtils"
Option Explicit

Public Function ClearData(ByRef ws As Worksheet, ByVal rangeName As String, Optional useMessage As Boolean, Optional message As String = "データをクリアしますか？") As Boolean
    Dim msgRet As VbMsgBoxResult
    If (useMessage) Then
        msgRet = MsgBox(message, "クリアの確認", vbYesNo)
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

Function GetWorksheetByName(sheetName As String, silentmode As Boolean) As Worksheet
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

