Attribute VB_Name = "FileUtils"
Option Explicit

' �t�@�C������
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


' �����Ɋ�Â����t�@�C������
Function SearchFilesByCondition(fso As Object, folderPath As String, Optional keyword As String = "", Optional extension As String = "", Optional minSize As Long = 0, Optional maxSize As Long = -1, Optional modifiedAfter As Date = #1/1/1900#, Optional modifiedBefore As Date = #12/31/9999#) As Collection
    Dim files As Collection
    Dim folder As Object, file As Object
    
    Set files = New Collection
    Set folder = fso.GetFolder(folderPath)
    
    For Each file In folder.files
        ' �������[�h
        If keyword <> "" And InStr(1, file.Name, keyword, vbTextCompare) = 0 Then
            GoTo NextFile
        End If
        
        ' �g���q
        If extension <> "" And Right(file.Name, Len(extension)) <> extension Then
            GoTo NextFile
        End If
        
        ' �T�C�Y
        If (file.Size < minSize) Or (maxSize >= 0 And file.Size > maxSize) Then
            GoTo NextFile
        End If
        
        ' �X�V����
        If (file.DateLastModified < modifiedAfter) Or (file.DateLastModified > modifiedBefore) Then
            GoTo NextFile
        End If
        
        ' �����𖞂����ꍇ�ɃR���N�V�����ɒǉ�
        files.Add file.path
        
NextFile:
    Next file
    
    Set SearchFilesByCondition = files
End Function

' �t�@�C�����݃`�F�b�N
Function FileExists(fso As Object, filePath As String) As Boolean
    FileExists = fso.FileExists(filePath)
End Function

' �t�@�C����ʊm�F
Function IsFolder(fso As Object, path As String) As Boolean
    IsFolder = fso.FolderExists(path)
End Function

' �t�@�C���T�C�Y�m�F
Function GetFileSize(fso As Object, filePath As String) As Long
    Dim file As Object
    If FileExists(fso, filePath) Then
        Set file = fso.GetFile(filePath)
        GetFileSize = file.Size
    Else
        GetFileSize = -1 ' �t�@�C�������݂��Ȃ��ꍇ
    End If
End Function

' �t�@�C�������m�F
Function GetFileAttributes(fso As Object, filePath As String) As String
    Dim file As Object
    If FileExists(fso, filePath) Then
        Set file = fso.GetFile(filePath)
        GetFileAttributes = file.Attributes
    Else
        GetFileAttributes = "File not found"
    End If
End Function

' �t�@�C���ړ�
Sub moveFile(fso As Object, sourcePath As String, destinationPath As String)
    If FileExists(fso, sourcePath) Then
        fso.moveFile sourcePath, destinationPath
    End If
End Sub

' �t�@�C���폜
Sub DeleteFile(fso As Object, filePath As String)
    If FileExists(fso, filePath) Then
        fso.DeleteFile filePath
    End If
End Sub

' �t�@�C�����l�[��
Sub renameFile(fso As Object, filePath As String, newName As String)
    Dim file As Object
    If FileExists(fso, filePath) Then
        Set file = fso.GetFile(filePath)
        file.Name = newName
    End If
End Sub

' �t�@�C������
Sub copyFile(fso As Object, sourcePath As String, destinationPath As String)
    If FileExists(fso, sourcePath) Then
        fso.copyFile sourcePath, destinationPath
    End If
End Sub

' �t�@�C���v���p�e�B�l�̎擾�i��: �쐬���ƍX�V���j
Function GetFileProperties(fso As Object, filePath As String) As String
    Dim file As Object
    If FileExists(fso, filePath) Then
        Set file = fso.GetFile(filePath)
        GetFileProperties = "Created: " & file.DateCreated & ", Last Modified: " & file.DateLastModified
    Else
        GetFileProperties = "File not found"
    End If
End Function

' �t�H���_�擾
Function GetFolder(fso As Object, folderPath As String) As Object
    On Error Resume Next
    Set GetFolder = fso.GetFolder(folderPath)
    On Error GoTo 0
End Function

' �t�@�C���擾
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
