Attribute VB_Name = "modDocuments"
'==============================================================================
' modDocuments - fills Word templates and exports PDF (late bound, no
' references needed). One Word instance per batch, one bad row never stops it.
'==============================================================================
Option Explicit

' Word constants (late binding, so no reference to the Word library is needed)
Private Const wdAlertsNone As Long = 0
Private Const wdFindStop As Long = 0
Private Const wdCollapseEnd As Long = 0
Private Const wdDoNotSaveChanges As Long = 0
Private Const wdExportFormatPDF As Long = 17
Private Const wdExportOptimizeForPrint As Long = 0
Private Const wdFormatDocumentDefault As Long = 16

' Placeholder syntax {{FieldName}} as a Word wildcard pattern
Private Const PLACEHOLDER_PATTERN As String = "\{\{[A-Za-z0-9_]@\}\}"

Public Function StartWord() As Object
    Dim wd As Object
    Set wd = CreateObject("Word.Application")
    wd.Visible = False
    wd.Options.SaveNormalPrompt = False
    wd.DisplayAlerts = wdAlertsNone
    wd.ScreenUpdating = False
    Set StartWord = wd
End Function

Public Sub QuitWord(ByRef wd As Object)
    If wd Is Nothing Then Exit Sub
    On Error Resume Next
    Do While wd.Documents.Count > 0
        wd.Documents(1).Close SaveChanges:=wdDoNotSaveChanges
    Loop
    wd.Quit SaveChanges:=wdDoNotSaveChanges
    Set wd = Nothing
End Sub

' Builds the placeholder -> text dictionary for one row:
' Settings keys + every column of tblAgreements + computed fields.
Public Function RowValues(ByVal lo As ListObject, ByVal rowIx As Long, ByVal s As Object, _
                          ByVal templateVersion As String) As Object
    Dim d As Object, k As Variant, c As ListColumn, df As String, nf As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    df = SettingOr(s, "DateFormat", "dd.mm.yyyy")
    nf = SettingOr(s, "NumberFormat", "#,##0")

    For Each k In s.Keys
        d(k) = s(k)
    Next k
    For Each c In lo.ListColumns
        d(c.Name) = FormatValue(lo.DataBodyRange.Cells(rowIx, c.Index).Value, df, nf)
    Next c
    d("DocumentDate") = Format$(Date, df)
    d("TemplateVersion") = templateVersion
    Set RowValues = d
End Function

Private Function FormatValue(ByVal v As Variant, ByVal df As String, ByVal nf As String) As String
    If IsError(v) Or IsEmpty(v) Then
        FormatValue = ""
    ElseIf VarType(v) = vbDate Then
        FormatValue = Format$(v, df)
    ElseIf VarType(v) = vbString Then
        ' Alt+Enter line breaks in a cell become line breaks in Word
        FormatValue = Replace(Replace(Trim$(v), vbCrLf, Chr$(11)), vbLf, Chr$(11))
    ElseIf IsNumeric(v) Then
        FormatValue = Format$(v, nf)
    Else
        FormatValue = CStr(v)
    End If
End Function

' Creates one agreement. Returns "" on success, otherwise the reason it failed.
' pdfPath returns the path of the PDF that was written.
Public Function GenerateOne(ByVal wd As Object, ByVal lo As ListObject, ByVal rowIx As Long, _
                            ByVal s As Object, ByVal templates As Object, ByRef pdfPath As String) As String
    Dim doc As Object, tpl As Variant, tplPath As String, folder As String, baseName As String
    Dim problems As String, stage As String

    On Error GoTo Fail
    stage = "preparing paths"
    tpl = templates(FldText(lo, rowIx, "AgreementType"))
    tplPath = TemplateFolder(s) & tpl(0)
    folder = OutputFolder(s, "Generated")
    EnsureFolder folder

    baseName = SafeFileName(tpl(1) & "_" & FldText(lo, rowIx, "AgreementID") & "_" & _
                            FldText(lo, rowIx, "CounterpartyName"))
    pdfPath = folder & baseName & ".pdf"
    If FileExists(pdfPath) And Not SettingYes(s, "OverwriteExisting") Then
        GenerateOne = "PDF already exists (set OverwriteExisting = Yes to regenerate): " & pdfPath
        Exit Function
    End If

    ' Open the template READ-ONLY and never save it: the filled copy is only
    ' exported (PDF) or saved under a new name (.docx), so the template can't
    ' be overwritten. (Documents.Add was less reliable with .docx templates.)
    stage = "opening template " & tplPath
    Set doc = wd.Documents.Open(FileName:=tplPath, ConfirmConversions:=False, ReadOnly:=True, _
                                AddToRecentFiles:=False, Visible:=False)

    stage = "filling placeholders"
    problems = FillDocument(doc, RowValues(lo, rowIx, s, CStr(tpl(2))))
    If Len(problems) > 0 Then
        doc.Close SaveChanges:=wdDoNotSaveChanges
        GenerateOne = "Not exported. " & problems
        Exit Function
    End If

    stage = "exporting PDF to " & pdfPath
    If SettingYes(s, "PdfA") Then
        doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF, _
            OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, UseISO19005_1:=True
    Else
        doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF, _
            OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint
    End If

    If SettingYes(s, "KeepDocx") Then
        stage = "saving .docx"
        doc.SaveAs2 FileName:=folder & baseName & ".docx", FileFormat:=wdFormatDocumentDefault
    End If
    stage = "closing document"
    doc.Close SaveChanges:=wdDoNotSaveChanges
    Exit Function

Fail:
    GenerateOne = "Error " & Err.Number & " while " & stage & ": " & Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=wdDoNotSaveChanges
End Function

' Replaces every {{Field}} in every part of the document: body, headers,
' footers, footnotes and text boxes. Returns "" or a list of problems.
' The template decides which fields are needed - there is no mapping list
' to maintain in code.
Private Function FillDocument(ByVal doc As Object, ByVal values As Object) As String
    Dim story As Object, problems As Object, k As Variant, out As String
    Set problems = CreateObject("Scripting.Dictionary")

    For Each story In doc.StoryRanges
        Do
            FillRange story, values, problems
            Set story = story.NextStoryRange
        Loop Until story Is Nothing
    Next story

    For Each k In problems.Keys
        out = out & k & "; "
    Next k
    If Len(out) > 0 Then out = Left$(out, Len(out) - 2)
    FillDocument = out
End Function

Private Sub FillRange(ByVal storyRange As Object, ByVal values As Object, ByVal problems As Object)
    Dim r As Object, key As String
    Set r = storyRange.Duplicate
    With r.Find
        .ClearFormatting
        .Text = PLACEHOLDER_PATTERN
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
    End With

    Do While r.Find.Execute
        key = Mid$(r.Text, 3, Len(r.Text) - 4)
        If Not values.Exists(key) Then
            problems("Unknown placeholder {{" & key & "}}") = True
        ElseIf Len(values(key)) = 0 Then
            problems("No value for {{" & key & "}}") = True
        Else
            ' Setting .Text (instead of Find.Replacement) keeps the formatting
            ' and has no 255-character limit.
            r.Text = values(key)
        End If
        r.Collapse wdCollapseEnd
    Loop
End Sub

'--------------------------------------------------------------- Outlook ------
' Connects to a running classic Outlook, or starts it. errMsg explains a failure.
' Only classic Outlook for Windows can be automated; the new Outlook and
' Outlook on the web cannot.
Public Function StartOutlook(Optional ByRef errMsg As String) As Object
    On Error Resume Next
    Set StartOutlook = GetObject(, "Outlook.Application")
    If StartOutlook Is Nothing Then
        Err.Clear
        Set StartOutlook = CreateObject("Outlook.Application")
        If StartOutlook Is Nothing Then errMsg = "Error " & Err.Number & ": " & Err.Description
    End If
End Function

' Subject and body shared by Outlook drafts and .eml files.
Private Function SigningSubject(ByVal lo As ListObject, ByVal rowIx As Long, ByVal s As Object) As String
    SigningSubject = FldText(lo, rowIx, "AgreementType") & " for signing - " & _
                     FldText(lo, rowIx, "AgreementID") & " - " & SettingOr(s, "OurCompanyName", "")
End Function

Private Function SigningBody(ByVal lo As ListObject, ByVal rowIx As Long, ByVal s As Object) As String
    SigningBody = "Dear " & FldText(lo, rowIx, "ContactName") & "," & vbCrLf & vbCrLf & _
                  "Please find attached the " & LCase$(FldText(lo, rowIx, "AgreementType")) & _
                  " (" & FldText(lo, rowIx, "AgreementID") & ") for review and signature." & vbCrLf & vbCrLf & _
                  "Kind regards," & vbCrLf & FldText(lo, rowIx, "AccountOwner") & vbCrLf & _
                  SettingOr(s, "OurCompanyName", "")
End Function

' Saves an Outlook draft (does not send) with the agreement attached.
Public Function CreateSigningDraft(ByVal ol As Object, ByVal lo As ListObject, ByVal rowIx As Long, _
                                   ByVal s As Object) As String
    Dim m As Object
    On Error GoTo Fail
    Set m = ol.CreateItem(0)
    With m
        .To = FldText(lo, rowIx, "ContactEmail")
        .Subject = SigningSubject(lo, rowIx, s)
        .Body = SigningBody(lo, rowIx, s)
        .Attachments.Add FldText(lo, rowIx, "DocumentPath")
        .Save
    End With
    Exit Function
Fail:
    CreateSigningDraft = "Outlook error " & Err.Number & ": " & Err.Description
End Function

'------------------------------------------------------------ .eml files ------
' Writes a ready-to-send e-mail file (RFC 822 / MIME) with the PDF attached,
' for PCs without classic Outlook. "X-Unsent: 1" makes mail apps that support
' it (classic Outlook, Windows Mail and others) open it as an editable draft.
' Everything non-ASCII is Base64-encoded, so Norwegian letters survive.
Public Function CreateEmlDraft(ByVal lo As ListObject, ByVal rowIx As Long, ByVal s As Object, _
                               ByRef emlPath As String) As String
    Dim pdf As String, folder As String, boundary As String, t As String, fname As String
    On Error GoTo Fail
    pdf = FldText(lo, rowIx, "DocumentPath")
    fname = FileNameOf(pdf)
    folder = OutputFolder(s, "Outbox")
    EnsureFolder folder
    emlPath = folder & Replace(fname, ".pdf", ".eml", , , vbTextCompare)
    boundary = "----=_AgreementGenerator_" & Format$(Now, "yyyymmddhhnnss") & "_" & rowIx

    t = "To: " & EncodeWord(FldText(lo, rowIx, "ContactName")) & " <" & FldText(lo, rowIx, "ContactEmail") & ">" & vbCrLf
    t = t & "Subject: " & EncodeWord(SigningSubject(lo, rowIx, s)) & vbCrLf
    t = t & "X-Unsent: 1" & vbCrLf
    t = t & "MIME-Version: 1.0" & vbCrLf
    t = t & "Content-Type: multipart/mixed; boundary=""" & boundary & """" & vbCrLf & vbCrLf
    t = t & "This is a multi-part message in MIME format." & vbCrLf & vbCrLf
    t = t & "--" & boundary & vbCrLf
    t = t & "Content-Type: text/plain; charset=""utf-8""" & vbCrLf
    t = t & "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf
    t = t & Base64(Utf8Bytes(SigningBody(lo, rowIx, s)), True) & vbCrLf
    t = t & "--" & boundary & vbCrLf
    t = t & "Content-Type: application/pdf; name=""" & EncodeWord(fname) & """" & vbCrLf
    t = t & "Content-Disposition: attachment; filename=""" & EncodeWord(fname) & """" & vbCrLf
    t = t & "Content-Transfer-Encoding: base64" & vbCrLf & vbCrLf
    t = t & Base64(FileBytes(pdf), True) & vbCrLf
    t = t & "--" & boundary & "--" & vbCrLf

    WriteTextFile emlPath, t
    Exit Function
Fail:
    CreateEmlDraft = "E-mail file error " & Err.Number & ": " & Err.Description
End Function

' Header text: plain if ASCII, otherwise a MIME encoded-word (=?UTF-8?B?...?=).
Private Function EncodeWord(ByVal text As String) As String
    Dim i As Long
    For i = 1 To Len(text)
        If AscW(Mid$(text, i, 1)) > 126 Or AscW(Mid$(text, i, 1)) < 32 Then
            EncodeWord = "=?UTF-8?B?" & Base64(Utf8Bytes(text), False) & "?="
            Exit Function
        End If
    Next i
    EncodeWord = text
End Function

Private Function Utf8Bytes(ByVal text As String) As Variant
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' text
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.Position = 0
    st.Type = 1                 ' binary
    st.Position = 3             ' skip the UTF-8 byte order mark
    Utf8Bytes = st.Read
    st.Close
End Function

Private Function FileBytes(ByVal path As String) As Variant
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.LoadFromFile path
    FileBytes = st.Read
    st.Close
End Function

' Base64 via MSXML. lineBreaks = True gives the 76-character lines MIME expects.
Private Function Base64(ByVal bytes As Variant, ByVal lineBreaks As Boolean) As String
    Dim node As Object, out As String
    Set node = CreateObject("MSXML2.DOMDocument").createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    out = Replace(node.Text, vbCr, "")
    If lineBreaks Then
        Base64 = Replace(out, vbLf, vbCrLf)
    Else
        Base64 = Replace(out, vbLf, "")
    End If
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, text;
    Close #f
End Sub
