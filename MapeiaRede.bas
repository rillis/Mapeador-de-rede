Attribute VB_Name = "MapeiaRede"
Sub example()
    'Shows all file paths in C:\Windows\Branding
    Dim paths() As String
    paths = getFilesURL("C:\Windows\Branding")
    
    'Folder exists check
    If Length(paths) = 0 Then
        Debug.Print "Folder not found"
        Exit Sub
    End If
    
    Dim path As Variant
    
    'File count
    Debug.Print Length(paths) & " files."
    
    For Each path In paths
        'Print path
        Debug.Print path
        
        'Getting info
        Dim file As Object
        Set file = getFileInfo(path)
        
        'Using info
        Debug.Print file.Type
        Debug.Print file.Size
        'See more uses in https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-object
    Next
End Sub

Public Function getFilesURL(path As String) As String()
    If FolderExists(path) = False Then
        Dim paths() As String
        getFilesURL = paths
        Exit Function
    End If
    
    Dim result() As String
    result() = Split(CreateObject("wscript.shell").Exec("cmd /c dir """ & path & """ /b/s").StdOut.ReadAll, vbCrLf)
    
    Dim final() As String
    
    Dim cont As Integer
    i = 0
    
    For Each r In result
        If InStr(r, ".") > 0 Then
            i = i + 1
        End If
    Next
    
    ReDim final(i - 1)
    
    i = 0
    For Each r In result
        If InStr(r, ".") > 0 Then
            final(i) = r
            i = i + 1
        End If
    Next
    getFilesURL = final
End Function


Function getFileInfo(ByVal path As String) As Object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Set getFileInfo = fso.GetFile(path)
End Function

Public Function Length(arr As Variant) As Long
    Dim le As Long
    On Error Resume Next
    le = UBound(arr) - LBound(arr) + 1
    If Err.Number <> 0 Then
        le = 0
        On Error GoTo 0
    End If
    Length = le
End Function

Public Function FolderExists(strFolderPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strFolderPath) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

