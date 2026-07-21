Attribute VB_Name = "ExportMOMToDraftModule"
Option Explicit

Private Const adReadAll As Long = -1
Private Const adTypeText As Long = 2
Private Const OFN_FILEMUSTEXIST As Long = &H1000&
Private Const OFN_PATHMUSTEXIST As Long = &H800&
Private Const OFN_HIDEREADONLY As Long = &H4&
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_NOCHANGEDIR As Long = &H8&
Private Const MAX_FILE_PATH As Long = 32768

#If VBA7 Then
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As LongPtr
        hInstance As LongPtr
        lpstrFilter As LongPtr
        lpstrCustomFilter As LongPtr
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As LongPtr
        nMaxFile As Long
        lpstrFileTitle As LongPtr
        nMaxFileTitle As Long
        lpstrInitialDir As LongPtr
        lpstrTitle As LongPtr
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As LongPtr
        lCustData As LongPtr
        lpfnHook As LongPtr
        lpTemplateName As LongPtr
        pvReserved As LongPtr
        dwReserved As Long
        FlagsEx As Long
    End Type

    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" _
        (ByRef pOpenfilename As OPENFILENAME) As Long

    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
        (ByVal hwnd As LongPtr) As Long
#End If

Public Sub ExportMOMToDraft()
    On Error GoTo ErrorHandler

    Dim folderPath As String
    folderPath = Environ$("USERPROFILE") & "\Downloads\ExportMOM\"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    Dim htmlFilePath As String
    htmlFilePath = PickHtmlFile(folderPath)

    If Len(htmlFilePath) = 0 Then
        MsgBox "No HTML file selected. Draft tidak dibuat.", vbInformation, "Export MOM to Draft"
        Exit Sub
    End If

    Dim projectName As String
    projectName = Trim$(InputBox("Nama Project:", "Export MOM to Draft"))

    If Len(projectName) = 0 Then
        MsgBox "Nama project kosong. Draft tidak dibuat.", vbInformation, "Export MOM to Draft"
        Exit Sub
    End If

    Dim draft As Outlook.MailItem
    Set draft = Application.CreateItem(olMailItem)

    Dim htmlContent As String
    htmlContent = ReadTextFileUtf8(htmlFilePath)

    With draft
        .Subject = "MOM Meeting Persiapan Implementasi " & projectName
        .BodyFormat = olFormatHTML
        .HTMLBody = htmlContent
        .Save
        .Display
    End With

    Exit Sub

ErrorHandler:
    MsgBox "Export MOM to Draft stopped: " & Err.Description, vbExclamation, "Export MOM to Draft"
End Sub

Private Function PickHtmlFile(ByVal folderPath As String) As String
    Dim fileDialog As OPENFILENAME
    Dim fileBuffer As String
    Dim filterText As String
    Dim titleText As String
    Dim defaultExtension As String
    Dim ownerHwnd As LongPtr

    fileBuffer = String$(MAX_FILE_PATH, vbNullChar)
    filterText = "HTML Files (*.html;*.htm)" & vbNullChar & "*.html;*.htm" & vbNullChar & vbNullChar
    titleText = "Pilih file MOM HTML"
    defaultExtension = "html"
    ownerHwnd = GetOutlookWindowHwnd()

    If ownerHwnd <> 0 Then SetForegroundWindow ownerHwnd

    With fileDialog
        .lStructSize = LenB(fileDialog)
        .hwndOwner = ownerHwnd
        .lpstrFilter = StrPtr(filterText)
        .nFilterIndex = 1
        .lpstrFile = StrPtr(fileBuffer)
        .nMaxFile = MAX_FILE_PATH
        .lpstrInitialDir = StrPtr(folderPath)
        .lpstrTitle = StrPtr(titleText)
        .lpstrDefExt = StrPtr(defaultExtension)
        .Flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR
    End With

    If GetOpenFileName(fileDialog) <> 0 Then
        PickHtmlFile = Left$(fileBuffer, InStr(1, fileBuffer, vbNullChar) - 1)
    End If
End Function

Private Function GetOutlookWindowHwnd() As LongPtr
    On Error Resume Next

    GetOutlookWindowHwnd = Application.ActiveWindow.HWND

    If GetOutlookWindowHwnd = 0 Then
        GetOutlookWindowHwnd = Application.ActiveExplorer.HWND
    End If

    On Error GoTo 0
End Function

Private Function ReadTextFileUtf8(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Type = adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
        ReadTextFileUtf8 = .ReadText(adReadAll)
        .Close
    End With
End Function
