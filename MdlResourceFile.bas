Attribute VB_Name = "MdlResourceFile"
Option Explicit
Private sTempDir As String
Public FullPath As String

Public Function ExtractResources() As Boolean
    Dim FSO As Object, oShell As Object, oItem As Object
    Dim sName As String
    Dim t As Long, nFile As Long
    
    Dim SourceFile As String
    Dim sZipPath As String
    
    SourceFile = ThisWorkbook.FullName
    
    Set FSO = CreateObject("scripting.filesystemobject")
    Set oShell = CreateObject("Shell.Application")
    
    sName = Mid(SourceFile, InStrRev(SourceFile, "\") + 1)
    sName = Left$(sName, InStrRev(sName, ".") - 1)
    sTempDir = Environ("Temp") & "\Res-" & sName
    
    On Error Resume Next
    
    
    If FSO.FolderExists(sTempDir) Then
        FSO.DeleteFolder sTempDir
        t = Timer
        Do While Err.Number <> 0
            Err.Clear
            FSO.DeleteFolder sTempDir
            If t + 5 < Timer Then
                nFile = nFile + 1
                t = Timer
                sTempDir = Environ("Temp") & "\Res-" & sName & nFile
                Err.Clear
                If FSO.FolderExists(sTempDir) Then FSO.DeleteFolder sTempDir
            End If
        Loop
    End If

    FSO.CreateFolder sTempDir
    Do While Err.Number <> 0
        Err.Clear
        FSO.CreateFolder sTempDir
    Loop
    On Error GoTo 0
    
    sZipPath = sTempDir & "\" & sName & ".zip"
    
    FSO.CopyFile SourceFile, sZipPath, True
    
    FullPath = sTempDir & "\resource\"
    
    For Each oItem In oShell.Namespace(sZipPath & "\").Items
        If oItem = "resource" Then
            oShell.Namespace(sTempDir & "\").CopyHere sZipPath & "\resource"
        End If
    Next
    
    On Error Resume Next
    Do
        Err.Clear
            
        Do While oShell.Namespace(sZipPath & "\resource").Items.Count <> oShell.Namespace(sTempDir & "\resource").Items.Count
            DoEvents
        Loop
        If Err.Number = 0 Then Exit Do
    Loop
    
    Kill sZipPath

    ExtractResources = True
End Function

Public Sub RemoveFiles()
    On Error Resume Next
    Dim FSO As Object
    Set FSO = CreateObject("scripting.filesystemobject")
    If FSO.FolderExists(sTempDir) Then
        FSO.DeleteFolder sTempDir
    End If
End Sub


