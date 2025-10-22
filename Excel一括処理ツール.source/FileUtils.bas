Attribute VB_Name = "FileUtils"
Option Explicit

' ファイル検索
Function SearchFiles(fso As Object, folderPath As String, pattern As String) As Collection
    Dim files As Collection
    Dim folder As Object, file As Object
    
    Set files = New Collection
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.files
        If file.Name Like pattern Then
            files.Add file.path
        End If
    Next file
    
    Set SearchFiles = files
End Function


' 条件に基づいたファイル検索
Function SearchFilesByCondition(fso As Object, folderPath As String, Optional keyword As String = "", Optional extension As String = "", Optional minSize As Long = 0, Optional maxSize As Long = -1, Optional modifiedAfter As Date = #1/1/1900#, Optional modifiedBefore As Date = #12/31/9999#) As Collection
    Dim files As Collection
    Dim folder As Object, file As Object
    
    Set files = New Collection
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.files
        ' 検索ワード
        If keyword <> "" And InStr(1, file.Name, keyword, vbTextCompare) = 0 Then
            GoTo NextFile
        End If
        
        ' 拡張子
        If extension <> "" And Right(file.Name, Len(extension)) <> extension Then
            GoTo NextFile
        End If
        
        ' サイズ
        If (file.Size < minSize) Or (maxSize >= 0 And file.Size > maxSize) Then
            GoTo NextFile
        End If
        
        ' 更新日時
        If (file.DateLastModified < modifiedAfter) Or (file.DateLastModified > modifiedBefore) Then
            GoTo NextFile
        End If
        
        ' 条件を満たす場合にコレクションに追加
        files.Add file.path
        
NextFile:
    Next file
    
    Set SearchFilesByCondition = files
End Function

' ファイル存在チェック
Function FileExists(fso As Object, filePath As String) As Boolean
    FileExists = fso.FileExists(filePath)
End Function

' ファイル種別確認
Function IsFolder(fso As Object, path As String) As Boolean
    IsFolder = fso.FolderExists(path)
End Function

' ファイルサイズ確認
Function GetFileSize(fso As Object, filePath As String) As Long
    Dim file As Object
    If FileExists(fso, filePath) Then
        Set file = fso.GetFile(filePath)
        GetFileSize = file.Size
    Else
        GetFileSize = -1 ' ファイルが存在しない場合
    End If
End Function

' ファイル属性確認
Function GetFileAttributes(fso As Object, filePath As String) As String
    Dim file As Object
    If FileExists(fso, filePath) Then
        Set file = fso.GetFile(filePath)
        GetFileAttributes = file.Attributes
    Else
        GetFileAttributes = "File not found"
    End If
End Function

' ファイル移動
Sub moveFile(fso As Object, sourcePath As String, destinationPath As String)
    If FileExists(fso, sourcePath) Then
        fso.moveFile sourcePath, destinationPath
    End If
End Sub

' ファイル削除
Sub DeleteFile(fso As Object, filePath As String)
    If FileExists(fso, filePath) Then
        fso.DeleteFile filePath
    End If
End Sub

' ファイルリネーム
Sub renameFile(fso As Object, filePath As String, newName As String)
    Dim file As Object
    If FileExists(fso, filePath) Then
        Set file = fso.GetFile(filePath)
        file.Name = newName
    End If
End Sub

' ファイル複製
Sub copyFile(fso As Object, sourcePath As String, destinationPath As String)
    If FileExists(fso, sourcePath) Then
        fso.copyFile sourcePath, destinationPath
    End If
End Sub

' ファイルプロパティ値の取得（例: 作成日と更新日）
Function GetFileProperties(fso As Object, filePath As String) As String
    Dim file As Object
    If FileExists(fso, filePath) Then
        Set file = fso.GetFile(filePath)
        GetFileProperties = "Created: " & file.DateCreated & ", Last Modified: " & file.DateLastModified
    Else
        GetFileProperties = "File not found"
    End If
End Function

' フォルダ取得
Function GetFolder(fso As Object, folderPath As String) As Object
    On Error Resume Next
    Set GetFolder = fso.GetFolder(folderPath)
    On Error GoTo 0
End Function

' ファイル取得
Function GetFile(fso As Object, filePath As String) As Object
    On Error Resume Next
    Set GetFile = fso.GetFile(filePath)
    On Error GoTo 0
End Function

Function AddFolderDelimiter(ByVal filePath As String) As String
    If Right(filePath, 1) = "\" Or Right(filePath, 1) = "/" Then
        AddFolderDelimiter = filePath
    Else
        AddFolderDelimiter = filePath & "\"
    End If
End Function
