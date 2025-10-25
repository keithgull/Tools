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
            Case VBIDE.vbext_ComponentType.vbext_ct_ActiveXDesigner:
                logstr = "Skipping Active X Designer type: " & vbc.Name
            Case VBIDE.vbext_ComponentType.vbext_ct_StdModule:
                blnExport = True
                strExt = ".bas"
                logstr = "Exporting Standard Module: " & vbc.Name & strExt
            Case VBIDE.vbext_ComponentType.vbext_ct_MSForm:
                blnExport = True
                strExt = ".frm"
                logstr = "Exporting MS Form: " & vbc.Name & strExt
            Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule:
                blnExport = True
                strExt = ".cls"
                logstr = "Exporting Class module: " & vbc.Name & strExt
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
        
        wb.SaveAs dirName & bookName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    End If
    
    If wb.FullName <> ThisWorkbook.FullName Then
        'Stop
        wb.Close False
    End If
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
           ext = "frm" Or _
           ext = "dco" Then
            wb.VBProject.VBComponents.Import moduleFile.path
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

Sub ExcelFileRef()
    Dim ws As Worksheet
    Dim rngTargetBook As Range
    Dim ret As String
    Dim defPath As String
    
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Set rngTargetBook = ws.Range("ORIGINAL_BOOK")
    defPath = rngTargetBook.Value
    If defPath = "" Then
        defPath = ThisWorkbook.path
    End If
    ret = SelectFileAndSetPath(defPath, "Excel book", "*.xlsx", "マクロ取込元のExcelブックを設定してください。", "", False)
    If ret <> "" Then
        rngTargetBook.Value = ret
    End If
    
    Set ws = Nothing
    Set rngTargetBook = Nothing
    
End Sub

Sub ImportFolderRef()
    Dim ws As Worksheet
    Dim ret As String
    Dim defPath As String
    Dim rngImportPath As Range
    
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Set rngImportPath = ws.Range("MACRO_FOLDER")
    
    defPath = rngImportPath.Value
    If defPath = "" Then
        defPath = ThisWorkbook.path
    End If
    ret = SelectFolderAndSetPath(defPath, "インポート対象のマクロモジュールのフォルダを設定してください。", "", False)
    If ret <> "" Then
        rngImportPath.Value = ret
    End If
    
    Set ws = Nothing
    Set rngImportPath = Nothing
    
End Sub

Sub OutputFolderRef()
    Dim ws As Worksheet
    Dim ret As String
    Dim defPath As String
    Dim rngOutputPath As Range
    
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    Set rngOutputPath = ws.Range("OUTPUT_MACRO_PATH")
    
    defPath = rngOutputPath.Value
    If defPath = "" Then
        defPath = ThisWorkbook.path
    End If
    ret = SelectFolderAndSetPath(defPath, "マクロ統合ブックの出力先のフォルダを設定してください。", "", False)
    If ret <> "" Then
        rngOutputPath.Value = ret
    End If
    
    Set ws = Nothing
    Set rngOutputPath = Nothing

End Sub
