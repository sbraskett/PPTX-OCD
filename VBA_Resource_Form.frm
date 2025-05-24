VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "VBA Resource  File Editor"
   ClientHeight    =   4320
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   10995
   OleObjectBlob   =   "VBA_Resource_Form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As LongPtr, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
    'Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    'Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#End If

Private Const GWL_STYLE  As Long = (-16)
Private Const GWL_EXSTYLE As Long = -20
Private Const GW_OWNER As Long = 4

Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX  As Long = &H10000
Private Const WS_THICKFRAME  As Long = &H40000
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_POPUP As Long = &H80000000

Private Const WM_SETICON As Long = &H80
Private Const ICON_BIG As Long = 1
Private Const ICON_SMALL As Long = 0

Private Const IMAGE_ICON = 1
Private Const LR_LOADFROMFILE = &H10

Private Enum FolderViewMode
    FVM_ICON = 1
    FVM_SMALLICON = 2
    FVM_LIST = 3
    FVM_DETAILS = 4
    FVM_THUMBNAIL = 5
    FVM_TILE = 6
    FVM_THUMBSTRIP = 7
End Enum

Dim WithEvents LocalBrowser As ShellFolderView
Attribute LocalBrowser.VB_VarHelpID = -1

Private sTempDir As String
Private SourceFile As String
Private sZipPath As String
Private hwnd

Private Sub WriteRelationship()
    Dim XDoc As MSXML2.DOMDocument
    Dim listNode As MSXML2.IXMLDOMElement
    Dim Elemnt As MSXML2.IXMLDOMElement
    Dim cListResFile As Collection
    Dim cId As Collection
    Dim bExist As Boolean
    Dim oShell As Object, oItem As Object
    Dim i As Long
    Dim ID As Integer

    Set oShell = CreateObject("Shell.Application")
    
    Set cListResFile = New Collection
    Set cId = New Collection
    
    
    oShell.Namespace(sTempDir & "\").MoveHere sZipPath & "\_rels", 4
    
    Do While Len(Dir(sTempDir & "\_rels\.rels")) = 0
    
    Loop
    
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (sTempDir & "\_rels\.rels")

  
    For Each listNode In XDoc.DocumentElement.ChildNodes
        For i = 0 To listNode.Attributes.Length - 1
            If listNode.Attributes(i).nodeName = "Target" Then
                
               If InStr(listNode.Attributes(i).Text, "resource") = 1 Then
                    If Len(Dir(sTempDir & "\" & listNode.Attributes(i).Text)) = 0 Then
                        XDoc.DocumentElement.RemoveChild listNode
                    Else
                        cListResFile.Add listNode.Attributes(i).Text
                    End If
               End If
            End If
            If listNode.Attributes(i).nodeName = "Id" Then
                cId.Add listNode.Attributes(i).Text
            End If
        Next
    Next
    
    For Each oItem In oShell.Namespace(sTempDir & "\resource\").Items
        bExist = False
        For i = 1 To cListResFile.Count
            If cListResFile(i) = "resource/" & oItem Then
                bExist = True
            End If
        Next
        
        If bExist = False Then
            ID = 0
            Do
                ID = ID + 1
                bExist = False
                For i = 1 To cId.Count
                    If cId(i) = "rId" & ID Then
                        bExist = True
                        Exit For
                    End If
                Next
                If bExist = False Then Exit Do
            Loop
            cId.Add "rId" & ID
            Set Elemnt = XDoc.createElement("Relationship")
            Elemnt.setAttribute "Target", "resource/" & oItem
            Elemnt.setAttribute "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/resource"
            Elemnt.setAttribute "Id", "rId" & ID
        
            XDoc.DocumentElement.appendChild Elemnt
        End If
    
    Next
    
    XDoc.LoadXML Replace(XDoc.XML, " xmlns=" & Chr(34) & Chr(34), "")
    XDoc.Save sTempDir & "\_rels\.rels"
    
    oShell.Namespace(sZipPath & "\").MoveHere sTempDir & "\_rels", 4
    
    Do While Len(Dir(sTempDir & "\_rels\.rels")) <> 0
    
    Loop
    
    Set XDoc = Nothing
End Sub

Private Function GetMimeType(sExtension As String) As String
    On Error GoTo Defaul
    Dim objShell, strTemp
        
    Set objShell = CreateObject("WScript.Shell")
    
    GetMimeType = objShell.RegRead("HKCR\." & sExtension & "\Content Type")
    If Len(GetMimeType) = 0 Then GetMimeType = "application/octet-stream"
    Exit Function
Defaul:
    GetMimeType = "application/octet-stream"
End Function

Private Sub WriteContentTypes()

    Dim XDoc As MSXML2.DOMDocument
    Dim listNode As MSXML2.IXMLDOMElement
    Dim Elemnt As MSXML2.IXMLDOMElement
    Dim i As Long, j As Long
    Dim cExtension As Collection
    Dim sFileExt As String
    Dim bExist As Boolean
    Dim oShell As Object, oItem As Object
    

    Set oShell = CreateObject("Shell.Application")
    
    oShell.Namespace(sTempDir & "\").MoveHere sZipPath & "\[Content_Types].xml", 4
    
    Do While Len(Dir(sTempDir & "\[Content_Types].xml")) = 0
    
    Loop
    
    'For Each oItem In oShell.Namespace(sZipPath & "\").Items
    '    If oItem = "[Content_Types].xml" Then
    '         oShell.Namespace(sTempDir & "\").MoveHere oItem
    '        Exit For
    '    End If
    'Next
    
    Set cExtension = New Collection
    
    
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (sTempDir & "\[Content_Types].xml")
  
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
            If UCase(sFileExt) = UCase(cExtension(i)) Then
                bExist = True
            End If
        Next
        
        If Not bExist Then
            cExtension.Add sFileExt
            Set Elemnt = XDoc.createElement("Default")
            Elemnt.setAttribute "ContentType", GetMimeType(sFileExt)
            Elemnt.setAttribute "Extension", LCase(sFileExt)
            XDoc.DocumentElement.appendChild Elemnt
        End If
    Next
    
    XDoc.LoadXML Replace(XDoc.XML, " xmlns=" & Chr(34) & Chr(34), "")
    XDoc.Save sTempDir & "\[Content_Types].xml"
    
    oShell.Namespace(sZipPath & "\").MoveHere sTempDir & "\[Content_Types].xml", 4
    
    Do While Len(Dir(sTempDir & "\[Content_Types].xml")) <> 0
    
    Loop
    
    Set XDoc = Nothing

End Sub



Private Sub CmdAbout_Click()
    ThisWorkbook.FollowHyperlink ("http://www.leandroascierto.com")
End Sub

Private Sub CmdAdd_Click()
    Dim FileNames As Variant
    Dim sName As String
    Dim i As Integer

    FileNames = Application.GetOpenFilename(, , , , True)
     
    If IsArray(FileNames) Then
        For i = LBound(FileNames) To UBound(FileNames)
            sName = Mid(FileNames(i), InStrRev(FileNames(i), "\") + 1)
            FileCopy FileNames(i), sTempDir & "\resource\" & sName
        Next i
    End If
    
    WebBrowser1.Refresh2
    WebBrowser1.SetFocus
End Sub



Private Sub CmdOpen_Click()
    Dim FSO As Object, oShell As Object, oItem As Object
    Dim sName As String
    Dim vFile As Variant
    Dim FF As Integer
    Dim lFileHead As Long
    
    On Error Resume Next
    
    vFile = Application.GetOpenFilename()
   
    If VarType(vFile) <> vbString Then Exit Sub
    
    FF = FreeFile
    Open vFile For Binary Access Read Lock Read Write As #FF
        Get #FF, 1, lFileHead
    Close #FF
    
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbInformation, "Error " & Err.Number
        Exit Sub
    End If
    
    If lFileHead <> &H4034B50 Then 'zip head
        MsgBox "Tipo de archivo no compatible", vbInformation
        Exit Sub
    End If
       
    SourceFile = vFile
    
    Set FSO = CreateObject("scripting.filesystemobject")
    Set oShell = CreateObject("Shell.Application")
    
    If Len(sTempDir) Then
        If FSO.FolderExists(sTempDir) Then
            FSO.DeleteFolder sTempDir
        End If
    End If
    On Error GoTo 0
    
    sName = Mid(SourceFile, InStrRev(SourceFile, "\") + 1)
    sName = Left$(sName, InStrRev(sName, ".") - 1)

    LblInfo.Caption = "Abriendo " & Mid(SourceFile, InStrRev(SourceFile, "\") + 1)
    DoEvents

    sTempDir = Environ("Temp") & "\Resource-" & sName
    
    On Error Resume Next
    
    If FSO.FolderExists(sTempDir) Then
        FSO.DeleteFolder sTempDir
        Do While Err.Number <> 0
            Err.Clear
            FSO.DeleteFolder sTempDir
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
    
    
    FSO.CreateFolder sTempDir & "\resource"
    
    For Each oItem In oShell.Namespace(sZipPath & "\").Items
        If oItem = "resource" Then
            oShell.Namespace(sTempDir & "\resource").MoveHere oShell.Namespace(sZipPath & "\resource\").Items 'oItem
        End If
    Next
    
    WebBrowser1.Navigate2 sTempDir & "\resource"
    WebBrowser1.Document.CurrentViewMode = FVM_DETAILS
    
    CmdSave.Enabled = True
    CmdAdd.Enabled = True
    CmdDelete.Enabled = True
    CmdRename.Enabled = True
    LblInfo.Caption = "Listo"
    WebBrowser1.SetFocus
    Me.Caption = SourceFile
    Call ChangeFormStyle
End Sub

Private Sub CmdSave_Click()
    Dim oShell As Object, oItem As Object, FSO As Object
    Dim bExist As Boolean
    
    Set oShell = CreateObject("Shell.Application")
    Set FSO = CreateObject("scripting.filesystemobject")
    
    On Error Resume Next
    

    For Each oItem In LocalBrowser.Folder.Items
        If oItem.Name <> EncodeURL(oItem.Name) Then
            MsgBox "The name cannot contain spaces, accents or special characters"
            LocalBrowser.SelectItem oItem, 1 Or 3
            Exit Sub
        End If
    Next
        
    LblInfo.Caption = "Creando backup..."
    DoEvents
    If Len(Dir(SourceFile & ".backup")) Then Kill SourceFile & ".backup"
    FSO.CopyFile SourceFile, SourceFile & ".backup", True
    
    Kill SourceFile
    
    Do While Err.Number = 70
        If MsgBox("Error writing file, if you are using it please close it to continue.", vbRetryCancel) = vbRetry Then
            Kill SourceFile
        Else
            Exit Sub
        End If
    Loop
    
    On Error GoTo 0
    
    LblInfo.Caption = "Save content type..."
    DoEvents
    WriteContentTypes
    
    LblInfo.Caption = "Save Relationship..."
    DoEvents
    WriteRelationship

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
            Do While oShell.Namespace(sZipPath & "\resource").Items.Count <> oShell.Namespace(sTempDir & "\resource\").Items.Count
                DoEvents
            Loop
            If Err.Number = 0 Then Exit Do
        Loop
    End If
    
    LblInfo.Caption = "Finishing..."
    DoEvents

    FSO.CopyFile sZipPath, SourceFile, True
    
    LblInfo.Caption = "List"
    WebBrowser1.SetFocus
    Beep
End Sub

Private Sub CmdRename_Click()
    WebBrowser1.SetFocus
    SendKeys "{F2}"
End Sub

Private Sub CmdDelete_Click()
    WebBrowser1.SetFocus
    SendKeys "{del}"
End Sub

Private Sub CmdDirUp_Click()
    WebBrowser1.GoBack
    WebBrowser1.Document.CurrentViewMode = FVM_DETAILS
End Sub

Private Sub LocalBrowser_SelectionChanged()
    Dim lCount As Long
    Dim sCaption As String
    
    lCount = LocalBrowser.Folder.Items.Count
    If lCount > 0 Then
        sCaption = lCount & " Item" & IIf(lCount > 1, "s", "")
        lCount = LocalBrowser.SelectedItems.Count
        If lCount > 0 Then
            sCaption = sCaption & vbTab & lCount & " Item" & IIf(lCount > 1, "s", "") & " selection" & IIf(lCount > 1, "s", "")
        End If
    End If
    LblInfo.Caption = sCaption
End Sub

Private Function LocalBrowser_VerbInvoked() As Boolean
    LblInfo.Caption = ""
    LocalBrowser_VerbInvoked = True
End Function

Private Sub UserForm_Activate()
    ChangeFormStyle
End Sub

Private Sub UserForm_Initialize()
    hwnd = FindWindow(vbNullString, Me.Caption)
    UserForm_Resize
End Sub

Private Sub ChangeFormStyle()
    Call SetWindowLong(hwnd, -8, 0)
    ShowWindow hwnd, 0
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX Or WS_THICKFRAME Or WS_POPUP Or WS_SYSMENU
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_APPWINDOW
    ShowWindow hwnd, 1
    DoEvents
    Call ChangeIcon(MdlResourceFile.FullPath & "CubeIcon.ico", hwnd)
    WebBrowser1.SetFocus
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
    
    Dim FSO As Object
    If Len(sTempDir) Then
        Set FSO = CreateObject("scripting.filesystemobject")
        If FSO.FolderExists(sTempDir) Then
            FSO.DeleteFolder sTempDir
        End If
    End If
    
End Sub

Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
    If Command = CSC_NAVIGATEBACK Then
        CmdDirUp.Enabled = Enable
    End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Set LocalBrowser = WebBrowser1.Document
    LocalBrowser_SelectionChanged
    WebBrowser1.SetFocus
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    TextBox1.Text = (WebBrowser1.LocationURL)
End Sub


Public Sub ChangeIcon(sPath As String, hwnd)
    Dim hIcon 'As LongPtr
    hIcon = LoadImage(0, sPath, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
    SendMessage hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon
    
    hIcon = LoadImage(0, sPath, IMAGE_ICON, 32, 32, LR_LOADFROMFILE)
    SendMessage hwnd, WM_SETICON, ICON_BIG, ByVal hIcon
End Sub


Private Function EncodeURL(ByVal sURL As String) As String
    Dim i           As Long
    Dim sChar       As String * 1

    For i = 1 To Len(sURL)
        sChar = Mid$(sURL, i, 1)
        Select Case sChar
            Case "a" To "z", "A" To "Z", "0" To "9", "-", "_", ".", "~"
                EncodeURL = EncodeURL & sChar
            Case Else
                EncodeURL = EncodeURL & "%" & Right$("0" & Hex(Asc(sChar)), 2)
        End Select
    Next
End Function
