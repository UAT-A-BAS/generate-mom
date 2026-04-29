Attribute VB_Name = "ExportMOMToDraftModule"
Option Explicit

Private Const adReadAll As Long = -1
Private Const adTypeText As Long = 2
Private Const wdRowHeightExactly As Long = 2
Private Const wdCellAlignVerticalCenter As Long = 1
Private Const wdAlignParagraphCenter As Long = 1
Private Const wdLineSpaceSingle As Long = 0
Private Const wdPreferredWidthPoints As Long = 3
Private Const wdAdjustNone As Long = 0
Private Const OUTLOOK_TABLE2_TARGET_WIDTH_POINTS As Single = 82.5
Private Const OUTLOOK_TABLE3_DATE_WIDTH_POINTS As Single = 90
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
    htmlContent = FixTable2HeaderForOutlook(htmlContent)

    With draft
        .Subject = "MOM Meeting Persiapan Implementasi " & projectName
        .BodyFormat = olFormatHTML
        .HTMLBody = htmlContent
        .Save
        .Display
    End With

    FixDisplayedTableHeaders draft
    draft.Save

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

Private Function FixTable2HeaderForOutlook(ByVal htmlContent As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .Global = False
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "(<table\b[^>]*class\s*=\s*[""'][^""']*table2[^""']*[""'][^>]*>[\s\S]*?<thead\b[^>]*>\s*)<tr\b[^>]*>[\s\S]*?</tr>"
    End With

    If regex.Test(htmlContent) Then
        FixTable2HeaderForOutlook = regex.Replace(htmlContent, "$1" & GetTable2HeaderRowHtml())
    Else
        FixTable2HeaderForOutlook = htmlContent
    End If
End Function

Private Function GetTable2HeaderRowHtml() As String
    Const headerStyle As String = "background:#9bd255;border:1px solid #111;color:#111;font-weight:700;text-align:center;vertical-align:middle;padding:6px 8px;height:34px;line-height:1.15;mso-line-height-rule:exactly;mso-height-source:userset;"

    GetTable2HeaderRowHtml = _
        "<tr style=""height:34px;mso-height-source:userset;"">" & _
        "<th width=""70"" height=""34"" valign=""middle"" bgcolor=""#9bd255"" style=""width:70px;" & headerStyle & """>No.</th>" & _
        "<th width=""240"" height=""34"" valign=""middle"" bgcolor=""#9bd255"" style=""width:240px;" & headerStyle & """>Aktivitas</th>" & _
        "<th width=""130"" height=""34"" valign=""middle"" bgcolor=""#9bd255"" style=""width:130px;" & headerStyle & """>Status</th>" & _
        "<th width=""125"" height=""34"" valign=""middle"" bgcolor=""#9bd255"" style=""width:125px;" & headerStyle & """>PIC</th>" & _
        "<th width=""125"" height=""34"" valign=""middle"" bgcolor=""#9bd255"" style=""width:125px;" & headerStyle & """>Target</th>" & _
        "<th width=""245"" height=""34"" valign=""middle"" bgcolor=""#9bd255"" style=""width:245px;" & headerStyle & """>Keterangan</th>" & _
        "</tr>"
End Function

Private Sub FixDisplayedTableHeaders(ByVal draft As Outlook.MailItem)
    On Error Resume Next

    Dim wordDoc As Object
    Set wordDoc = draft.GetInspector.WordEditor
    If wordDoc Is Nothing Then Exit Sub

    Dim tbl As Object
    Dim isChecklist As Boolean
    Dim isStrategy As Boolean

    For Each tbl In wordDoc.Tables
        isChecklist = IsChecklistTable(tbl)
        isStrategy = IsStrategyTable(tbl)

        If IsCertificationTable(tbl) Or isChecklist Or isStrategy Then
            FixWordTableHeaderRow tbl
        End If

        If isChecklist Then
            SetWordTableColumnWidth tbl, 5, OUTLOOK_TABLE2_TARGET_WIDTH_POINTS
        End If

        If isStrategy Then
            SetWordTableColumnWidth tbl, 1, OUTLOOK_TABLE3_DATE_WIDTH_POINTS
        End If
    Next tbl
End Sub

Private Function IsCertificationTable(ByVal tbl As Object) As Boolean
    On Error GoTo SafeExit

    If tbl.Rows.Count = 0 Then Exit Function
    If tbl.Rows(1).Cells.Count < 6 Then Exit Function

    Dim headerText As String
    Dim cell As Object

    For Each cell In tbl.Rows(1).Cells
        headerText = headerText & " " & CleanWordCellText(cell.Range.Text)
    Next cell

    headerText = UCase$(headerText)

    IsCertificationTable = _
        InStr(headerText, "NOMOR") > 0 And _
        InStr(headerText, "BPRO") > 0 And _
        InStr(headerText, "CHANGES") > 0 And _
        InStr(headerText, "RELEASE") > 0 And _
        InStr(headerText, "BLUEPRINT") > 0

SafeExit:
End Function

Private Function IsChecklistTable(ByVal tbl As Object) As Boolean
    On Error GoTo SafeExit

    If tbl.Rows.Count = 0 Then Exit Function
    If tbl.Rows(1).Cells.Count < 6 Then Exit Function

    Dim headerText As String
    Dim cell As Object

    For Each cell In tbl.Rows(1).Cells
        headerText = headerText & " " & CleanWordCellText(cell.Range.Text)
    Next cell

    headerText = UCase$(headerText)

    IsChecklistTable = _
        InStr(headerText, "NO") > 0 And _
        InStr(headerText, "AKTIVITAS") > 0 And _
        InStr(headerText, "STATUS") > 0 And _
        InStr(headerText, "PIC") > 0 And _
        InStr(headerText, "TARGET") > 0 And _
        InStr(headerText, "KETERANGAN") > 0

SafeExit:
End Function

Private Function IsStrategyTable(ByVal tbl As Object) As Boolean
    On Error GoTo SafeExit

    If tbl.Rows.Count = 0 Then Exit Function
    If tbl.Rows(1).Cells.Count < 6 Then Exit Function

    Dim headerText As String
    Dim cell As Object

    For Each cell In tbl.Rows(1).Cells
        headerText = headerText & " " & CleanWordCellText(cell.Range.Text)
    Next cell

    headerText = UCase$(headerText)

    IsStrategyTable = _
        InStr(headerText, "TANGGAL") > 0 And _
        InStr(headerText, "JAM") > 0 And _
        InStr(headerText, "AKTIVITAS") > 0 And _
        InStr(headerText, "PIC") > 0 And _
        InStr(headerText, "STATUS") > 0 And _
        InStr(headerText, "KETERANGAN") > 0

SafeExit:
End Function

Private Sub FixWordTableHeaderRow(ByVal tbl As Object)
    On Error Resume Next

    tbl.AllowAutoFit = False
    tbl.AutoFitBehavior 0

    With tbl.Rows(1)
        .HeightRule = wdRowHeightExactly
        .Height = 26
        .AllowBreakAcrossPages = False
        .Range.Font.Bold = True
        .Range.Font.Size = 12
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .Range.ParagraphFormat.SpaceBefore = 0
        .Range.ParagraphFormat.SpaceAfter = 0
        .Cells.VerticalAlignment = wdCellAlignVerticalCenter
    End With
End Sub

Private Sub SetWordTableColumnWidth(ByVal tbl As Object, ByVal columnIndex As Long, ByVal widthPoints As Single)
    On Error Resume Next

    With tbl.Columns(columnIndex)
        .PreferredWidthType = wdPreferredWidthPoints
        .PreferredWidth = widthPoints
        .SetWidth widthPoints, wdAdjustNone
    End With
End Sub

Private Function CleanWordCellText(ByVal value As String) As String
    value = Replace(value, Chr$(13), "")
    value = Replace(value, Chr$(7), "")
    CleanWordCellText = Trim$(value)
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
