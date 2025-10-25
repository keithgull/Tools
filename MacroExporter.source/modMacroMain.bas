Attribute VB_Name = "modMacroMain"
Option Explicit

Const SHEET_MAIN As String = "マクロツール"


Sub auto_open()
    Dim ws As Worksheet
    Dim pathRng As Range
    
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Set pathRng = ws.Range("OUTPUT_PATH")
    If pathRng.Value = "" Then
        pathRng.Value = ThisWorkbook.path
    End If
    Set ws = Nothing
    Set pathRng = Nothing
End Sub


Sub Export()

    Dim wb As Workbook
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .AllowMultiSelect = False
        .Filters.Add "Macro-enabled", "*.xlsm"
        .Filters.Add "Macro-enabled (bin)", "*.xlsb"
        .Filters.Add "Macro-enabled (old)", "*.xls"
        .Filters.Add "Add-in", "*.xlsa"
        .Filters.Add "Add-in (old)", "*.xla"
        .InitialFileName = ThisWorkbook.path
        .Show
        If .SelectedItems.Count <= 0 Then
            MsgBox "Operation Cancelled", vbExclamation, "Operation Cancelled"
            Exit Sub
        Else
            Set wb = Workbooks.Open(.SelectedItems.Item(1), False, True)
            If wb.FullName <> ThisWorkbook.FullName Then
                ActiveWindow.Visible = False
                ThisWorkbook.Activate
            End If
        End If
    End With
    
    Dim vbc As VBComponent
    Dim strExt As String
    Dim blnExport As Boolean
    Dim strPath As String
    Dim ws As Worksheet
    Dim pathRng As Range
    Dim chkVal As Boolean
    Dim path As String
    Dim dirName As String
    Dim chkNoMacro As Boolean
    Dim bookName As String
    Dim clsCnt As Integer
    Dim modCnt As Integer
    Dim frmCnt As Integer
    Dim dcoCnt As Integer
    
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Set pathRng = ws.Range("OUTPUT_PATH")
    chkVal = ws.Range("CHECK_VALUE").Value
    chkNoMacro = ws.Range("CHECK_NOMACRO").Value
    
    path = pathRng.Value
    bookName = Left(wb.Name, InStrRev(wb.Name, ".") - 1)
    
    If chkVal Then
        dirName = path & "\" & bookName & "\"
        If Dir(dirName) = "" Then
            MkDir (dirName)
        End If
    Else
        dirName = path & "\"
    End If
    
    
    Dim logstr As String
    
    For Each vbc In wb.VBProject.VBComponents
        blnExport = False
        Select Case vbc.Type
            Case VBIDE.vbext_ComponentType.vbext_ct_Document:
                blnExport = True
                strExt = ".dco"
                logstr = "Exporting Document object: " & vbc.Name & strExt
                dcoCnt = dcoCnt + 1
            Case VBIDE.vbext_ComponentType.vbext_ct_ActiveXDesigner:
                logstr = "Skipping Active X Designer type: " & vbc.Name
            Case VBIDE.vbext_ComponentType.vbext_ct_StdModule:
                blnExport = True
                strExt = ".bas"
                logstr = "Exporting Standard Module: " & vbc.Name & strExt
                modCnt = modCnt + 1
            Case VBIDE.vbext_ComponentType.vbext_ct_MSForm:
                blnExport = True
                strExt = ".frm"
                logstr = "Exporting MS Form: " & vbc.Name & strExt
                frmCnt = frmCnt + 1
            Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule:
                blnExport = True
                strExt = ".cls"
                logstr = "Exporting Class module: " & vbc.Name & strExt
                clsCnt = clsCnt + 1
            Case Else:
                logstr = "Skipping unsupported type: " & vbc.Name & "(" & CStr(vbc.Type) & ")"
        End Select
        
        If blnExport Then
            strPath = dirName & vbc.Name & strExt
            vbc.Export strPath
            Debug.Print logstr & " Saved As " & strPath
        End If
    Next
    
    If chkNoMacro Then
        For Each vbc In wb.VBProject.VBComponents
            Select Case vbc.Type
                Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
                    wb.VBProject.VBComponents.Remove vbc
                Case vbext_ct_Document
                    vbc.CodeModule.DeleteLines 1, vbc.CodeModule.CountOfLines
            End Select
        Next
        wb.Windows(1).Visible = True
        wb.SaveAs dirName & bookName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    End If
    
    If wb.FullName <> ThisWorkbook.FullName Then
        'Stop
        wb.Close False
    End If
    MsgBox "エクスポート完了：モジュールファイルを出力しました。" & vbCrLf & _
        "　クラスモジュール" & Chr(9) & Chr(9) & "： " & clsCnt & vbCrLf & _
        "　標準モジュール" & Chr(9) & Chr(9) & "： " & modCnt & vbCrLf & _
        "　フォームモジュール" & Chr(9) & Chr(9) & "： " & frmCnt & vbCrLf & _
        "　Documentモジュール" & Chr(9) & "： " & dcoCnt & vbCrLf & _
        "　トータル" & Chr(9) & Chr(9) & Chr(9) & "： " & (clsCnt + modCnt + frmCnt + dcoCnt) & vbCrLf
    
End Sub

Sub Import()
    Dim folderPath As String
    Dim fso As Scripting.FileSystemObject
    Dim fileObj As Scripting.file
    Dim moduleFile As Scripting.file
    Dim xlsxFile As String
    Dim wb As Workbook
    Dim ext As String
    Dim ws As Worksheet
    Dim bookpath As String
    Dim importPath As String
    Dim outputPath As String
    
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    bookpath = ws.Range("ORIGINAL_BOOK").Value
    importPath = ws.Range("MACRO_FOLDER").Value
    outputPath = ws.Range("OUTPUT_MACRO_PATH").Value
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' xlsxファイルを特定
    Set fileObj = fso.GetFile(bookpath)
    If LCase(fso.GetExtensionName(fileObj.Name)) = "xlsx" Then
        xlsxFile = fileObj.path
    End If


    If xlsxFile = "" Then
        MsgBox "xlsxファイルが見つかりません"
        Exit Sub
    End If

    ' xlsxファイルを開く
    Set wb = Workbooks.Open(xlsxFile)

    ' モジュールファイルをインポート
    For Each moduleFile In fso.GetFolder(importPath).Files
        ext = LCase(fso.GetExtensionName(moduleFile.Name))
        If ext = "bas" Or _
           ext = "cls" Or _
           ext = "frm" Then
            wb.VBProject.VBComponents.Import moduleFile.path
        ElseIf ext = "dco" Then
            Dim moduleFileName As String
            Dim moduleName As String
            moduleFileName = Right(moduleFile.path, Len(moduleFile.path) - InStrRev(moduleFile.path, "\"))
            moduleName = Replace(moduleFileName, ".dco", "")
            'wb.VBProject.VBComponents.Import moduleFile.path
            Call InjectDocumentModuleCode(wb, moduleFile.path, moduleName)
        End If
    Next

    ' xlsmとして保存
    wb.SaveAs Replace(xlsxFile, ".xlsx", "_withMacro.xlsm"), FileFormat:=xlOpenXMLWorkbookMacroEnabled
    wb.Close SaveChanges:=False

    MsgBox "統合完了：xlsmとして保存しました"
    
    Set fso = Nothing
    Set wb = Nothing
    Set ws = Nothing
    Set fileObj = Nothing
End Sub

Sub InjectDocumentModuleCode(wb As Workbook, modulePath As String, moduleName As String)
    Dim targetComp As VBIDE.VBComponent
    Dim fso As Scripting.FileSystemObject
    Dim ts As Object
    Dim line As String, code As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(modulePath, 1)

    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        If Trim(line) Like "VERSION*" Or _
           Trim(line) = "BEGIN" Or _
           Trim(line) = "END" Or _
           Trim(line) Like "MultiUse*" Or _
           Trim(line) Like "Attribute VB_*" Then
            ' スキップ
        Else
            code = code & line & vbCrLf
        End If
    Loop
    ts.Close

    Set targetComp = wb.VBProject.VBComponents(moduleName) ' ← 対象モジュール名

    If code <> "" Then
        With targetComp.CodeModule
            .DeleteLines 1, .CountOfLines   ' 既存コードを削除
            .InsertLines 1, code            ' 新しいコードを注入
        End With
    End If
End Sub

Sub ExcelFileRef()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Call CommonFileRef(ws, "ORIGINAL_BOOK", ThisWorkbook.path, "Excel book", "*.xlsx", "マクロ取込元のExcelブックを設定してください。", "", False)
    
    Set ws = Nothing
End Sub


Sub ImportFolderRef()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Call CommonFolderRef(ws, "MACRO_FOLDER", ThisWorkbook.path, "インポート対象のマクロモジュールのフォルダを設定してください。", "", False)
        
    Set ws = Nothing
End Sub

Sub OutputFolderRef()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Call CommonFolderRef(ws, "OUTPUT_MACRO_PATH", ThisWorkbook.path, "マクロ統合ブックの出力先のフォルダを設定してください。", "", False)
        
    Set ws = Nothing
End Sub


