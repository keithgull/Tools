Attribute VB_Name = "SheetUtils"
Option Explicit

Public Function ClearData(ByRef ws As Worksheet, ByVal rangeName As String, Optional useMessage As Boolean, Optional message As String = "�f�[�^���N���A���܂����H") As Boolean
    Dim msgRet As VbMsgBoxResult
    If (useMessage) Then
        msgRet = MsgBox(message, vbYesNo, "�N���A�̊m�F")
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
    
    ' �e���v���[�g�V�[�g�̎擾
    Set template = GetWorksheetByName(templateSheetName, False)
    
    If template Is Nothing Then
        MsgBox "�e���v���[�g�V�[�g��������܂���B", vbExclamation
        Set AddSheet = Nothing
        Exit Function
    End If

    ' �V�����V�[�g�������ɑ��݂��邩�m�F
    If WorksheetExists(newSheetName) Then
        MsgBox "�V�[�g�����ɑ��݂��邽�߃X�L�b�v���܂��B", vbExclamation
        Set AddSheet = Nothing
        Exit Function
    End If
    
    ' �e���v���[�g�V�[�g���R�s�[���ĐV�����V�[�g��ǉ�
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
    
    ' �ǉ����ꂽ�V�[�g��߂�l�Ƃ��ĕԂ�
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
    
    ' �ΏۃV�[�g�̎擾
    Set ws = GetWorksheetByName(targetSheetName, False)
    
    If ws Is Nothing Then
        'MsgBox "�ΏۃV�[�g��������܂���B", vbExclamation
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
        MsgBox "�ΏۃV�[�g��������܂���B", vbExclamation
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

' �Z�����`�F�b�N�{�b�N�X�Ƃ����ۂɁA�S�`�F�b�N�����E�S�������s�����߂̏���
'
' �����F
'       ws                      : ���[�N�V�[�g
'       maxCount                : �f�[�^�̏�������
'       checkHeaderRngName      : �`�F�b�N�{�b�N�X��Ƃ�����̃w�b�_�s�̖��O
'       dataExistCheckRngName   : �f�[�^�̑��݃`�F�b�N���s����̖��O
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

' �{�^���������ꂽ���[�N�V�[�g����肵�A���[�N�V�[�g��Ԃ�
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
    Set GetButtonParentSheet = Nothing ' ������Ȃ������ꍇ
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
            target.Value = checkStr ' check ��\��
        End If
        cancel = True
    End If
        
    ' �w�b�_�̗̈�Ń_�u���N���b�N���s��ꂽ�ꍇ
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


