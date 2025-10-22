VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSOManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private fso As Object

' FSOの初期化
Private Sub Class_Initialize()
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
End Sub

' FSOの終了
Public Sub Class_Terminate()
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
End Sub

' ファイル検索
Public Function SearchFiles(folderPath As String, Optional keyword As String = "", Optional extension As String = "", Optional minSize As Long = 0, Optional maxSize As Long = -1, Optional modifiedAfter As Date = #1/1/1900#, Optional modifiedBefore As Date = #12/31/9999#) As Collection
    If fso Is Nothing Then Call Class_Initialize
    Set SearchFiles = SearchFilesByCondition(fso, folderPath, keyword, extension, minSize, maxSize, modifiedAfter, modifiedBefore)
End Function

' ファイル存在チェック
Public Function FileExists(filePath As String) As Boolean
    If fso Is Nothing Then Call Class_Initialize
    FileExists = FileUtils.FileExists(fso, filePath)
End Function

' ファイル種別確認
Public Function IsFolder(path As String) As Boolean
    If fso Is Nothing Then Call Class_Initialize
    IsFolder = FileUtils.IsFolder(fso, path)
End Function

' ファイルサイズ確認
Public Function GetFileSize(filePath As String) As Long
    If fso Is Nothing Then Call Class_Initialize
    GetFileSize = FileUtils.GetFileSize(fso, filePath)
End Function

' ファイル属性確認
Public Function GetFileAttributes(filePath As String) As String
    If fso Is Nothing Then Call Class_Initialize
    GetFileAttributes = FileUtils.GetFileAttributes(fso, filePath)
End Function

' ファイル移動
Public Sub moveFile(sourcePath As String, destinationPath As String)
    If fso Is Nothing Then Call Class_Initialize
    FileUtils.moveFile fso, sourcePath, destinationPath
End Sub

' ファイル削除
Public Sub DeleteFile(filePath As String)
    If fso Is Nothing Then Call Class_Initialize
    FileUtils.DeleteFile fso, filePath
End Sub

' ファイルリネーム
Public Sub renameFile(filePath As String, newName As String)
    If fso Is Nothing Then Call Class_Initialize
    FileUtils.renameFile fso, filePath, newName
End Sub

' ファイルコピー
Public Sub copyFile(sourcePath As String, destinationPath As String)
    If fso Is Nothing Then Call Class_Initialize
    FileUtils.copyFile fso, sourcePath, destinationPath
End Sub

' ファイルプロパティの取得
Public Function GetFileProperties(filePath As String) As String
    If fso Is Nothing Then Call Class_Initialize
    GetFileProperties = FileUtils.GetFileProperties(fso, filePath)
End Function

' フォルダ取得
Function GetFolder(folderPath As String) As Object
    If fso Is Nothing Then Call Class_Initialize
    Set GetFolder = FileUtils.GetFolder(fso, folderPath)
End Function

' ファイル取得
Function GetFile(filePath As String) As Object
    If fso Is Nothing Then Call Class_Initialize
    Set GetFile = FileUtils.GetFile(fso, filePath)
End Function


