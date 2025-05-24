Option Explicit

'————— Windows API declarations for VBA7/Win64 compatibility —————
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" ( _
            ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" ( _
            ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
            ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
            ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" ( _
        ByVal hInst As LongPtr, ByVal lpsz As String, ByVal un1 As Long, _
        ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private hwnd As LongPtr
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function ShowWindow Lib "user32" ( _
        ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
        ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
        ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private hwnd As Long
#End If

'————— Window style and icon constants —————
Private Const GWL_STYLE   As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_MINIMIZEBOX  As Long = &H20000
Private Const WS_MAXIMIZEBOX  As Long = &H10000
Private Const WS_THICKFRAME   As Long = &H40000
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_SYSMENU      As Long = &H80000
Private Const WS_POPUP        As Long = &H80000000

Private Const WM_SETICON    As Long = &H80
Private Const ICON_BIG      As Long = 1
Private Const ICON_SMALL    As Long = 0
Private Const IMAGE_ICON = 1
Private Const LR_LOADFROMFILE = &H10

'————— Shell view modes —————
Private Enum FolderViewMode
    FVM_ICON = 1
    FVM_SMALLICON = 2
    FVM_LIST = 3
    FVM_DETAILS = 4
    FVM_THUMBNAIL = 5
    FVM_TILE = 6
    FVM_THUMBSTRIP = 7
End Enum

'————— Late-binding browser view and temp-file tracking —————
Private LocalBrowser As Object
Private sTempDir   As String
Private SourceFile As String
Private sZipPath   As String
'—————————————————————————————————————————
' Apply custom window styles & set the icon
'—————————————————————————————————————————
Private Sub ChangeFormStyle()
    ' add minimize, maximize, thick frame, popup & sysmenu styles
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) _
        Or WS_MINIMIZEBOX _
        Or WS_MAXIMIZEBOX _
        Or WS_THICKFRAME _
        Or WS_POPUP _
        Or WS_SYSMENU

    ' ensure it shows as an application window (taskbar)
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) _
        Or WS_EX_APPWINDOW

    ' redraw the form
    ShowWindow hwnd, 1
    DoEvents

    ' load & set your custom icon (adjust path/module as needed)
    ChangeIcon MdlResourceFile.FullPath & "CubeIcon.ico", hwnd

    ' put focus back into the WebBrowser
    WebBrowser1.SetFocus
End Sub
'————— Update the status bar from the current folder view —————
Private Sub UpdateFolderViewStatus()
    Dim totalItems As Long, selItems As Long
    Dim captionText As String
    On Error Resume Next
    
    totalItems = LocalBrowser.Folder.Items.Count
    If totalItems > 0 Then
        captionText = totalItems & " Item" & IIf(totalItems > 1, "s", "")
        selItems = LocalBrowser.SelectedItems.Count
        If selItems > 0 Then
            captionText = captionText _
                & vbTab & selItems & " Item" & IIf(selItems > 1, "s", "") _
                & " selection" & IIf(selItems > 1, "s", "")
        End If
    End If
    
    LblInfo.Caption = captionText
End Sub

'————— WriteRelationship: Mirror and update .rels for any new resources —————
Private Sub WriteRelationship()
    Dim XDoc As MSXML2.DOMDocument
    Dim listNode As MSXML2.IXMLDOMElement
    Dim Elemnt As MSXML2.IXMLDOMElement
    Dim cListResFile As Collection
    Dim cId As Collection
    Dim bExist As Boolean
    Dim oShell As Object, oItem As Object
    Dim i As Long, ID As Integer

    Set oShell = CreateObject("Shell.Application")
    Set cListResFile = New Collection
    Set cId = New Collection
    
    oShell.Namespace(sTempDir & "\").MoveHere sZipPath & "\_rels", 4
    Do While Len(Dir(sTempDir & "\_rels\.rels")) = 0: Loop

    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load sTempDir & "\_rels\.rels"

    For Each listNode In XDoc.DocumentElement.ChildNodes
        For i = 0 To listNode.Attributes.Length - 1
            With listNode.Attributes(i)
                If .nodeName = "Target" And InStr(.Text, "resource") = 1 Then
                    If Len(Dir(sTempDir & "\" & .Text)) = 0 Then
                        XDoc.DocumentElement.RemoveChild listNode
                    Else
                        cListResFile.Add .Text
                    End If
                ElseIf .nodeName = "Id" Then
                    cId.Add .Text
                End If
            End With
        Next
    Next

    For Each oItem In oShell.Namespace(sTempDir & "\resource\").Items
        bExist = False
        For i = 1 To cListResFile.Count
            If cListResFile(i) = "resource/" & oItem Then bExist = True
        Next
        If Not bExist Then
            ID = 0
            Do
                ID = ID + 1
                bExist = False
                For i = 1 To cId.Count
                    If cId(i) = "rId" & ID Then bExist = True: Exit For
                Next
            Loop While bExist
            cId.Add "rId" & ID

            Set Elemnt = XDoc.createElement("Relationship")
            Elemnt.setAttribute "Target", "resource/" & oItem
            Elemnt.setAttribute "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/resource"
            Elemnt.setAttribute "Id", "rId" & ID
            XDoc.DocumentElement.appendChild Elemnt
        End If
    Next

    XDoc.LoadXML Replace(XDoc.XML, " xmlns=""" & """", "")
    XDoc.Save sTempDir & "\_rels\.rels"
    oShell.Namespace(sZipPath & "\").MoveHere sTempDir & "\_rels", 4
    Do While Len(Dir(sTempDir & "\_rels\.rels")) <> 0: Loop
    Set XDoc = Nothing
End Sub

'————— GetMimeType: Read content type from registry or default —————
Private Function GetMimeType(sExtension As String) As String
    On Error GoTo Defaul
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    GetMimeType = objShell.RegRead("HKCR\." & sExtension & "\Content Type")
    If Len(GetMimeType) = 0 Then GetMimeType = "application/octet-stream"
    Exit Function
Defaul:
    GetMimeType = "application/octet-stream"
End Function

'————— WriteContentTypes: Add any new extensions to [Content_Types].xml —————
Private Sub WriteContentTypes()
    Dim XDoc As MSXML2.DOMDocument
    Dim listNode As MSXML2.IXMLDOMElement
    Dim Elemnt As MSXML2.IXMLDOMElement
    Dim i As Long
    Dim cExtension As Collection
    Dim sFileExt As String
    Dim bExist As Boolean
    Dim oShell As Object, oItem As Object

    Set oShell = CreateObject("Shell.Application")
    oShell.Namespace(sTempDir & "\").MoveHere sZipPath & "\[Content_Types].xml", 4
    Do While Len(Dir(sTempDir & "\[Content_Types].xml")) = 0: Loop

    Set cExtension = New Collection
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load sTempDir & "\[Content_Types].xml"

    For Each listNode In XDoc.DocumentElement.ChildNodes
        If listNode.nodeName = "Default" Then
            For i = 0 To listNode.Attributes.Length - 1
                If listNode.Attributes(i).nodeName = "Extension" Then
                    cExtension.Add listNode.Attributes(i).Text
                End If
            Next
        End If
    Next

    For Each oItem In oShell.Namespace(sTempDir & "\resource\").Items
        sFileExt = Mid$(oItem, InStrRev(oItem, ".") + 1)
        bExist = False
        For i = 1 To cExtension.Count
            If UCase(sFileExt) = UCase(cExtension(i)) Then bExist = True
        Next
        If Not bExist Then
            cExtension.Add sFileExt
            Set Elemnt = XDoc.createElement("Default")
            Elemnt.setAttribute "ContentType", GetMimeType(sFileExt)
            Elemnt.setAttribute "Extension", LCase(sFileExt)
            XDoc.DocumentElement.appendChild Elemnt
        End If
    Next

    XDoc.LoadXML Replace(XDoc.XML, " xmlns=""" & """", "")
    XDoc.Save sTempDir & "\[Content_Types].xml"
    oShell.Namespace(sZipPath & "\").MoveHere sTempDir & "\[Content_Types].xml", 4
    Do While Len(Dir(sTempDir & "\[Content_Types].xml")) <> 0: Loop
    Set XDoc = Nothing
End Sub

'————— CmdAbout: Open hyperlink —————
Private Sub CmdAbout_Click()
    ThisWorkbook.FollowHyperlink "http://www.leandroascierto.com"
End Sub

'————— CmdAdd: Copy picked files into resource folder —————
Private Sub CmdAdd_Click()
    Dim FileNames As Variant, sName As String
    Dim i As Integer

    FileNames = Application.GetOpenFilename(, , , , True)
    If IsArray(FileNames) Then
        For i = LBound(FileNames) To UBound(FileNames)
            sName = Mid(FileNames(i), InStrRev(FileNames(i), "\") + 1)
            FileCopy FileNames(i), sTempDir & "\resource\" & sName
        Next
    End If
    
    WebBrowser1.Refresh2
    WebBrowser1.SetFocus
    UpdateFolderViewStatus
End Sub

'————— CmdOpen: Extract zip and navigate to resource folder —————
Private Sub CmdOpen_Click()
    Dim vFile As Variant
    Dim FF As Integer, lFileHead As Long
    Dim FSO As Object, oShell As Object, oItem As Object
    Dim sName As String
    Dim folderPath As String

    ' 1) Pick source package
    vFile = Application.GetOpenFilename()
    If VarType(vFile) <> vbString Then Exit Sub
    SourceFile = vFile

    ' 2) Validate zip header
    On Error Resume Next
    FF = FreeFile
    Open SourceFile For Binary Access Read Lock Read Write As #FF
        Get #FF, 1, lFileHead
    Close #FF
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation: Exit Sub
    If lFileHead <> &H4034B50 Then
        MsgBox "Tipo de archivo no compatible", vbInformation
        Exit Sub
    End If
    On Error GoTo 0

    ' 3) Prepare temp folder
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set oShell = CreateObject("Shell.Application")
    If Len(sTempDir) > 0 Then If FSO.FolderExists(sTempDir) Then FSO.DeleteFolder sTempDir

    ' 4) Build and create temp
    sName = Mid(SourceFile, InStrRev(SourceFile, "\") + 1)
    sName = Left$(sName, InStrRev(sName, ".") - 1)
    sTempDir = Environ("Temp") & "\Resource-" & sName
    FSO.CreateFolder sTempDir

    ' 5) Copy original into .zip
    sZipPath = sTempDir & "\" & sName & ".zip"
    FSO.CopyFile SourceFile, sZipPath, True

    ' 6) Extract resource folder only
    FSO.CreateFolder sTempDir & "\resource"
    For Each oItem In oShell.Namespace(sZipPath & "\").Items
        If oItem = "resource" Then
            oShell.Namespace(sTempDir & "\resource").MoveHere _
                oShell.Namespace(sZipPath & "\resource\").Items
        End If
    Next

    ' 7) Navigate the WebBrowser to resource
    folderPath = sTempDir & "\resource\"
    If Not FSO.FolderExists(folderPath) Then
        MsgBox "Folder not found: " & folderPath, vbExclamation
        Exit Sub
    End If
    WebBrowser1.Navigate folderPath
    WebBrowser1.Document.CurrentViewMode = FVM_DETAILS

    ' 8) Enable UI and finalize
    CmdSave.Enabled = True
    CmdAdd.Enabled = True
    CmdDelete.Enabled = True
    CmdRename.Enabled = True
    LblInfo.Caption = "Listo"
    Me.Caption = SourceFile
    ChangeFormStyle
    WebBrowser1.SetFocus
End Sub

'————— CmdSave: Validate names, backup, rewrite parts, rebuild zip —————
Private Sub CmdSave_Click()
    Dim oShell As Object, oItem As Object, FSO As Object
    Dim bExist As Boolean

    Set oShell = CreateObject("Shell.Application")
    Set FSO = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    For Each oItem In LocalBrowser.Folder.Items
        If oItem.Name <> EncodeURL(oItem.Name) Then
            MsgBox "The name cannot contain spaces, accents or special characters"
            LocalBrowser.SelectItem oItem, 1 Or 3
            Exit Sub
        End If
    Next
    On Error GoTo 0

    ' Backup
    LblInfo.Caption = "Creando backup..."
    DoEvents
    If Len(Dir(SourceFile & ".backup")) Then Kill SourceFile & ".backup"
    FSO.CopyFile SourceFile, SourceFile & ".backup", True
    Kill SourceFile
    Do While Err.Number = 70
        If MsgBox("Error writing file, if you are using it please close it to continue.", _
                  vbRetryCancel) = vbRetry Then
            Kill SourceFile
        Else
            Exit Sub
        End If
    Loop

    ' Rewrite parts
    LblInfo.Caption = "Save content type..."
    DoEvents
    WriteContentTypes

    LblInfo.Caption = "Save Relationship..."
    DoEvents
    WriteRelationship

    ' Move updated resources back
    For Each oItem In oShell.Namespace(sZipPath & "\").Items
        If oItem = "resource" Then
            LblInfo.Caption = "Deleting old content"
            DoEvents
            oShell.Namespace(ThisWorkbook.Path & "\").MoveHere sZipPath & "\resource"
            FSO.DeleteFolder ThisWorkbook.Path & "\resource"
        End If
    Next
    If oShell.Namespace(sTempDir & "\resource").Items.Count <> 0 Then
        LblInfo.Caption = "Saving Archive..."
        DoEvents
        oShell.Namespace(sZipPath & "\").CopyHere sTempDir & "\resource"
        On Error Resume Next
        Do
            Err.Clear
            Do While oShell.Namespace(sZipPath & "\resource").Items.Count <> _
                      oShell.Namespace(sTempDir & "\resource\").Items.Count
                DoEvents
            Loop
            If Err.Number = 0 Then Exit Do
        Loop
    End If

    ' Finalize save
    LblInfo.Caption = "Finishing..."
    DoEvents
    FSO.CopyFile sZipPath, SourceFile, True
    LblInfo.Caption = "List"
    WebBrowser1.SetFocus
    Beep
End Sub

'————— CmdRename / CmdDelete / CmdDirUp —————
Private Sub CmdRename_Click()
    WebBrowser1.SetFocus
    SendKeys "{F2}"
    UpdateFolderViewStatus
End Sub

Private Sub CmdDelete_Click()
    WebBrowser1.SetFocus
    SendKeys "{DEL}"
    UpdateFolderViewStatus
End Sub

Private Sub CmdDirUp_Click()
    WebBrowser1.GoBack
    WebBrowser1.Document.CurrentViewMode = FVM_DETAILS
End Sub

'————— UserForm lifecycle & styling —————
Private Sub UserForm_Activate()
    ChangeFormStyle
End Sub

Private Sub UserForm_Initialize()
    hwnd = FindWindow(vbNullString, Me.Caption)
    UserForm_Resize
End Sub

Private Sub UserForm_Resize()
    LblInfo.Top = Me.InsideHeight - LblInfo.Height
    LblInfo.Width = Me.InsideWidth
    ImgStatusBar.Width = LblInfo.Width
    ImgStatusBar.Top = LblInfo.Top - 2
    WebBrowser1.Height = Me.InsideHeight - LblInfo.Height - WebBrowser1.Top - 2
    WebBrowser1.Width = Me.InsideWidth
    TextBox1.Width = Me.InsideWidth - CmdDirUp.Width
    ImgToolbar.Width = Me.InsideWidth
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    If Len(sTempDir) Then CreateObject("Scripting.FileSystemObject").DeleteFolder sTempDir
End Sub

'————— Browser events —————
Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
    If Command = CSC_NAVIGATEBACK Then CmdDirUp.Enabled = Enable
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Set LocalBrowser = WebBrowser1.Document
    UpdateFolderViewStatus
    WebBrowser1.SetFocus
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    TextBox1.Text = WebBrowser1.LocationURL
End Sub

'————— ChangeIcon helper —————
Public Sub ChangeIcon(sPath As String, hwndTarget As LongPtr)
    Dim hIcon As LongPtr
    hIcon = LoadImage(0, sPath, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
    SendMessage hwndTarget, WM_SETICON, ICON_SMALL, ByVal hIcon
    hIcon = LoadImage(0, sPath, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
    SendMessage hwndTarget, WM_SETICON, ICON_BIG, ByVal hIcon
End Sub

'————— EncodeURL helper —————
Private Function EncodeURL(ByVal sURL As String) As String
    Dim i As Long, ch As String * 1
    For i = 1 To Len(sURL)
        ch = Mid$(sURL, i, 1)
        Select Case ch
            Case "a" To "z", "A" To "Z", "0" To "9", "-", "_", ".", "~"
                EncodeURL = EncodeURL & ch
            Case Else
                EncodeURL = EncodeURL & "%" & Right$("0" & Hex(Asc(ch)), 2)
        End Select
    Next
End Function


