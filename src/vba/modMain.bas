Attribute VB_Name = "modMain"
'==============================================================================
' modMain - entry points for the buttons on the Agreements sheet
'
' Flow:  Not created -> Created -> Sent for signing -> Signed -> Processed
' Each button acts on the selected rows, or on every eligible row when no
' cell inside the table is selected.
'==============================================================================
Option Explicit

'------------------------------------------------------------------ setup ------
' Palette ("Autumn Harvest", same as the workbook and the LinkedIn carousel)
Private Const C_NAVY As Long = &H392C20      ' #202C39  (VBA colours are BGR)
Private Const C_NAVY2 As Long = &H453828     ' #283845
Private Const C_SAND As Long = &H8DB0B8      ' #B8B08D
Private Const C_GOLD As Long = &H92D4F2      ' #F2D492
Private Const C_ORANGE As Long = &H5995F2    ' #F29559
Private Const C_WHITE As Long = &HFFFFFF

Public Sub Setup()
    Dim s As Object, ws As Worksheet, i As Long, x As Double
    On Error GoTo Fail
    Set s = LoadSettings()
    EnsureFolder OutputFolder(s, "Generated")
    EnsureFolder OutputFolder(s, "Signed")
    EnsureFolder OutputFolder(s, "Archive")
    EnsureFolder OutputFolder(s, "Outbox")

    Set ws = ThisWorkbook.Worksheets(SHEET_AGREEMENTS)
    For i = ws.Shapes.Count To 1 Step -1   ' backwards: deleting while looping forwards skips items
        If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i
    If ws.Rows(3).RowHeight < 36 Then ws.Rows(3).RowHeight = 40

    ' Buttons in workflow order, grouped with a wider gap between groups.
    x = ws.Range("A3").Left + 6
    x = AddButton(ws, x, "Validate", "ValidateAgreements", C_SAND, C_NAVY)
    x = AddButton(ws, x, "Generate documents", "GenerateDocuments", C_ORANGE, C_NAVY) + 14
    x = AddButton(ws, x, "Send for signing", "SendForSigning", C_NAVY, C_GOLD)
    x = AddButton(ws, x, "Import signed", "ImportSignedDocuments", C_NAVY, C_GOLD)
    x = AddButton(ws, x, "Mark signed", "MarkSigned", C_NAVY2, C_WHITE)
    x = AddButton(ws, x, "Process", "ProcessSigned", C_NAVY, C_GOLD) + 14
    x = AddButton(ws, x, "Reset", "ResetRows", C_WHITE, C_NAVY, C_SAND)
    x = AddButton(ws, x, "Demo: simulate signing", "DemoSimulateSigning", C_WHITE, C_NAVY, C_ORANGE)

    ws.Activate
    ws.Range("A6").Select
    MsgBox "Setup done." & vbLf & "Output folder: " & ResolveFolder(SettingOr(s, "OutputFolder", ""), "output"), _
           vbInformation, "Agreement Generator"
    Exit Sub
Fail:
    MsgBox "Setup failed: " & Err.Description, vbExclamation, "Agreement Generator"
End Sub

' Draws one rounded, flat button in row 3 and returns the x position for the next one.
' A shape with OnAction works like a button but can be styled, unlike Form buttons.
Private Function AddButton(ByVal ws As Worksheet, ByVal x As Double, ByVal caption As String, _
                           ByVal macroName As String, ByVal fillColor As Long, ByVal textColor As Long, _
                           Optional ByVal borderColor As Long = -1) As Double
    Const msoShapeRoundedRectangle As Long = 5
    Const BTN_HEIGHT As Double = 26
    Dim shp As Shape, w As Double, y As Double
    w = 26 + Len(caption) * 6.3
    y = ws.Range("A3").Top + (ws.Rows(3).RowHeight - BTN_HEIGHT) / 2
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, BTN_HEIGHT)
    With shp
        .Name = "btn_" & macroName
        .OnAction = macroName
        .Placement = xlFreeFloating
        .Adjustments.Item(1) = 0.3
        .Shadow.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = fillColor
        If borderColor = -1 Then
            .Line.Visible = msoFalse
        Else
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = borderColor
            .Line.Weight = 1.25
        End If
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse
            .MarginLeft = 4: .MarginRight = 4: .MarginTop = 0: .MarginBottom = 0
            With .TextRange
                .Text = caption
                .ParagraphFormat.Alignment = msoAlignCenter
                .Font.Name = "Segoe UI Semibold"
                .Font.Size = 10
                .Font.Fill.ForeColor.RGB = textColor
            End With
        End With
    End With
    AddButton = x + w + 8
End Function

'--------------------------------------------------------------- validate -----
Public Sub ValidateAgreements()
    Dim lo As ListObject, s As Object, tpls As Object, ids As Object
    Dim i As Long, bad As Long, msg As String, st As String
    On Error GoTo Fail
    Set lo = Agreements(): Set s = LoadSettings(): Set tpls = LoadTemplates(): Set ids = CountIds(lo)
    BeginBatch "Validating..."
    For i = 1 To RowCount(lo)
        st = FldText(lo, i, "Status")
        If st = ST_NOT_CREATED Or st = ST_CREATED Then
            msg = ValidateRow(lo, i, s, tpls, ids)
            Fld(lo, i, "LastError").Value = msg
            If Len(msg) > 0 Then bad = bad + 1
        End If
    Next i
    EndBatch
    MsgBox IIf(bad = 0, "All rows passed validation.", bad & " row(s) have problems. See the LastError column."), _
           IIf(bad = 0, vbInformation, vbExclamation), "Validate"
    Exit Sub
Fail:
    EndBatch
    MsgBox "Validation stopped: " & Err.Description, vbExclamation, "Validate"
End Sub

'--------------------------------------------------------------- generate -----
Public Sub GenerateDocuments()
    Dim lo As ListObject, s As Object, tpls As Object, ids As Object, wd As Object
    Dim targets As Collection, usedSel As Boolean, ix As Variant, n As Long
    Dim ok As Long, failed As Long, skipped As Long, msg As String, pdf As String

    On Error GoTo Fail
    Set lo = Agreements(): Set s = LoadSettings(): Set tpls = LoadTemplates(): Set ids = CountIds(lo)
    Set targets = ChooseRowsToGenerate(lo, usedSel)
    If targets Is Nothing Then Exit Sub          ' cancelled
    If targets.Count = 0 Then MsgBox "No rows with status '" & ST_NOT_CREATED & "'.", vbInformation: Exit Sub

    BeginBatch "Starting Word..."
    Set wd = StartWord()
    For Each ix In targets
        n = n + 1
        Progress "Generating", n, targets.Count
        If FldText(lo, ix, "Status") <> ST_NOT_CREATED Then
            skipped = skipped + 1
        Else
            msg = ValidateRow(lo, ix, s, tpls, ids)
            If Len(msg) = 0 Then msg = GenerateOne(wd, lo, ix, s, tpls, pdf)
            If Len(msg) = 0 Then
                SetPathCell Fld(lo, ix, "DocumentPath"), pdf
                MoveStatus lo, ix, ST_CREATED, "Generate", FileNameOf(pdf)
                ok = ok + 1
            Else
                SetRowError lo, ix, "Generate", msg
                failed = failed + 1
            End If
        End If
    Next ix
    QuitWord wd
    EndBatch
    Summary "Generate documents", usedSel, ok, failed, skipped, "created"
    Exit Sub
Fail:
    QuitWord wd
    EndBatch
    MsgBox "Generation stopped: " & Err.Description, vbExclamation, "Generate documents"
End Sub

' Asks which rows to generate and confirms before anything is created.
'  - Whole list: every row with status Not created
'  - Select rows: the user picks cells in the table
'      one row  -> confirm customer and agreement
'      several  -> yes/no with the count
' Returns Nothing when the user cancels.
Private Function ChooseRowsToGenerate(ByVal lo As ListObject, ByRef usedSel As Boolean) As Collection
    Dim picked As Collection, pick As Range, defaultAddr As String, wholeList As New Collection
    Dim answer As VbMsgBoxResult, ix As Variant, ready As Long, msg As String, i As Long

    For i = 1 To RowCount(lo)
        If FldText(lo, i, "Status") = ST_NOT_CREATED Then wholeList.Add i
    Next i

    answer = MsgBox("Which agreements do you want to generate?" & vbLf & vbLf & _
                    "Yes  =  the whole list (" & wholeList.Count & " rows with status '" & ST_NOT_CREATED & "')" & vbLf & _
                    "No   =  let me select rows" & vbLf & _
                    "Cancel  =  stop", vbYesNoCancel + vbQuestion, "Generate documents")
    If answer = vbCancel Then Exit Function
    If answer = vbYes Then
        usedSel = False
        Set ChooseRowsToGenerate = wholeList
        Exit Function
    End If

    ' Let the user point at the rows. Pre-fill with the current selection if it is in the table.
    lo.Parent.Activate
    If TypeName(Selection) = "Range" Then
        If RowsFromRange(lo, Selection).Count > 0 Then defaultAddr = Selection.Address
    End If
    On Error Resume Next
    Set pick = Application.InputBox("Select the row(s) to generate. Click a cell in each row " & _
                                    "(Ctrl+click for several, or drag over a range).", _
                                    "Generate documents", defaultAddr, Type:=8)
    On Error GoTo 0
    If pick Is Nothing Then Exit Function

    Set picked = RowsFromRange(lo, pick)
    If picked.Count = 0 Then
        MsgBox "The selection is not inside the Agreements table.", vbExclamation, "Generate documents"
        Exit Function
    End If
    usedSel = True

    For Each ix In picked
        If FldText(lo, ix, "Status") = ST_NOT_CREATED Then ready = ready + 1
    Next ix

    If picked.Count = 1 Then
        ix = picked(1)
        If ready = 0 Then
            MsgBox FldText(lo, ix, "AgreementID") & " has status '" & FldText(lo, ix, "Status") & _
                   "', so there is nothing to generate.", vbInformation, "Generate documents"
            Exit Function
        End If
        msg = "Create this agreement?" & vbLf & vbLf & _
              "Customer:" & vbTab & FldText(lo, ix, "CounterpartyName") & vbLf & _
              "Org. no.:" & vbTab & FldText(lo, ix, "CounterpartyOrgNo") & vbLf & _
              "Agreement:" & vbTab & FldText(lo, ix, "AgreementType") & vbLf & _
              "ID:" & vbTab & vbTab & FldText(lo, ix, "AgreementID") & vbLf & _
              "Contact:" & vbTab & FldText(lo, ix, "ContactName") & " (" & FldText(lo, ix, "ContactEmail") & ")"
    Else
        msg = "Create documents for " & ready & " selected row(s)?"
        If ready < picked.Count Then msg = msg & vbLf & vbLf & (picked.Count - ready) & _
            " of the " & picked.Count & " selected rows are not '" & ST_NOT_CREATED & "' and will be skipped."
        If ready = 0 Then
            MsgBox "None of the selected rows has status '" & ST_NOT_CREATED & "'.", vbInformation, "Generate documents"
            Exit Function
        End If
    End If

    If MsgBox(msg, vbYesNo + vbQuestion, "Generate documents") = vbYes Then Set ChooseRowsToGenerate = picked
End Function

'--------------------------------------------------------- send / signing -----
Public Sub SendForSigning()
    Dim lo As ListObject, s As Object, ol As Object, targets As Collection, usedSel As Boolean
    Dim ix As Variant, ok As Long, failed As Long, skipped As Long, msg As String
    Dim drafts As Boolean, eml As Boolean, emlPath As String, note As String, verb As String

    On Error GoTo Fail
    Set lo = Agreements(): Set s = LoadSettings()
    Set targets = TargetRows(lo, ST_CREATED, usedSel)
    If targets.Count = 0 Then MsgBox "No rows with status '" & ST_CREATED & "'.", vbInformation: Exit Sub

    drafts = SettingYes(s, "CreateOutlookDrafts")
    eml = SettingYes(s, "CreateEmlDrafts")
    If drafts Then
        Set ol = StartOutlook(msg)
        If ol Is Nothing Then
            If eml Then
                drafts = False   ' fall back to .eml files
            Else
                MsgBox "Could not connect to Outlook." & vbLf & msg & vbLf & vbLf & _
                       "Drafts need classic Outlook for Windows (not the new Outlook). Set CreateEmlDrafts = Yes " & _
                       "to save e-mail files instead, or CreateOutlookDrafts = No.", vbExclamation, "Send for signing"
                Exit Sub
            End If
        End If
    End If

    BeginBatch "Sending for signing..."
    For Each ix In targets
        If FldText(lo, ix, "Status") <> ST_CREATED Then
            skipped = skipped + 1
        ElseIf Not FileExists(FldText(lo, ix, "DocumentPath")) Then
            SetRowError lo, ix, "Send", "Document not found: " & FldText(lo, ix, "DocumentPath")
            failed = failed + 1
        Else
            msg = "": note = "Marked as sent"
            If drafts Then
                msg = CreateSigningDraft(ol, lo, ix, s)
                note = "Outlook draft to " & FldText(lo, ix, "ContactEmail")
            ElseIf eml Then
                msg = CreateEmlDraft(lo, ix, s, emlPath)
                note = "E-mail file " & FileNameOf(emlPath)
            End If
            If Len(msg) = 0 Then
                MoveStatus lo, ix, ST_SENT, "Send for signing", note
                ok = ok + 1
            Else
                SetRowError lo, ix, "Send", msg
                failed = failed + 1
            End If
        End If
    Next ix
    EndBatch

    verb = "sent for signing"
    If drafts Then verb = verb & " (drafts saved in Outlook > Drafts)"
    If eml And Not drafts Then verb = verb & " (e-mails saved in " & OutputFolder(s, "Outbox") & ")"
    Summary "Send for signing", usedSel, ok, failed, skipped, verb
    Exit Sub
Fail:
    EndBatch
    MsgBox "Stopped: " & Err.Description, vbExclamation, "Send for signing"
End Sub

' Picks up signed files dropped in output\Signed. A file matches a row when its
' name contains the AgreementID, e.g. SA_AGR-2026-001_Company_signed.pdf.
Public Sub ImportSignedDocuments()
    Dim lo As ListObject, s As Object, folder As String, f As Object, files As New Collection
    Dim i As Long, id As String, fileVar As Variant, ok As Long, waiting As Long

    On Error GoTo Fail
    Set lo = Agreements(): Set s = LoadSettings()
    folder = OutputFolder(s, "Signed")
    EnsureFolder folder
    For Each f In Fso().GetFolder(folder).Files
        If LCase$(Fso().GetExtensionName(f.Name)) = "pdf" Then files.Add f.Path
    Next f

    BeginBatch "Importing signed documents..."
    For i = 1 To RowCount(lo)
        If FldText(lo, i, "Status") = ST_SENT Then
            id = FldText(lo, i, "AgreementID")
            waiting = waiting + 1
            For Each fileVar In files
                If NameHasId(FileNameOf(CStr(fileVar)), id) Then
                    SetPathCell Fld(lo, i, "SignedDocumentPath"), CStr(fileVar)
                    MoveStatus lo, i, ST_SIGNED, "Import signed", FileNameOf(CStr(fileVar))
                    ok = ok + 1
                    Exit For
                End If
            Next fileVar
        End If
    Next i
    EndBatch
    MsgBox ok & " signed document(s) imported." & vbLf & (waiting - ok) & " agreement(s) still waiting for signature.", _
           vbInformation, "Import signed"
    Exit Sub
Fail:
    EndBatch
    MsgBox "Import stopped: " & Err.Description, vbExclamation, "Import signed"
End Sub

' Manual alternative, e.g. for agreements signed on paper. Selected rows only.
Public Sub MarkSigned()
    Dim lo As ListObject, targets As Collection, usedSel As Boolean, ix As Variant
    Dim ok As Long, skipped As Long
    On Error GoTo Fail
    Set lo = Agreements()
    Set targets = TargetRows(lo, ST_SENT, usedSel, requireSelection:=True)
    If targets.Count = 0 Then MsgBox "Select the rows to mark as signed first.", vbInformation: Exit Sub
    BeginBatch "Marking signed..."
    For Each ix In targets
        If FldText(lo, ix, "Status") = ST_SENT Then
            MoveStatus lo, ix, ST_SIGNED, "Mark signed", "Marked manually"
            ok = ok + 1
        Else
            skipped = skipped + 1
        End If
    Next ix
    EndBatch
    Summary "Mark signed", usedSel, ok, 0, skipped, "marked signed"
    Exit Sub
Fail:
    EndBatch
    MsgBox "Stopped: " & Err.Description, vbExclamation, "Mark signed"
End Sub

'---------------------------------------------------------------- process -----
' Signed -> Processed. Moves the signed file to output\Archive\<year>.
Public Sub ProcessSigned()
    Dim lo As ListObject, s As Object, targets As Collection, usedSel As Boolean, ix As Variant
    Dim src As String, dest As String, folder As String, ok As Long, failed As Long, skipped As Long

    On Error GoTo Fail
    Set lo = Agreements(): Set s = LoadSettings()
    Set targets = TargetRows(lo, ST_SIGNED, usedSel)
    If targets.Count = 0 Then MsgBox "No rows with status '" & ST_SIGNED & "'.", vbInformation: Exit Sub

    BeginBatch "Processing..."
    For Each ix In targets
        If FldText(lo, ix, "Status") <> ST_SIGNED Then
            skipped = skipped + 1
        Else
            src = FldText(lo, ix, "SignedDocumentPath")
            If Len(src) > 0 And FileExists(src) Then
                folder = OutputFolder(s, "Archive") & Year(Date) & Sep()
                EnsureFolder folder
                dest = folder & FileNameOf(src)
                If FileExists(dest) Then Fso().DeleteFile dest
                Fso().MoveFile src, dest
                SetPathCell Fld(lo, ix, "SignedDocumentPath"), dest
                MoveStatus lo, ix, ST_PROCESSED, "Process", "Archived to " & dest
            Else
                MoveStatus lo, ix, ST_PROCESSED, "Process", "No signed file on record (marked manually)"
            End If
            ok = ok + 1
        End If
    Next ix
    EndBatch
    Summary "Process", usedSel, ok, failed, skipped, "processed"
    Exit Sub
Fail:
    EndBatch
    MsgBox "Stopped: " & Err.Description, vbExclamation, "Process"
End Sub

'------------------------------------------------------------------ reset -----
Public Sub ResetRows()
    Dim lo As ListObject, targets As Collection, usedSel As Boolean, ix As Variant, reason As String
    On Error GoTo Fail
    Set lo = Agreements()
    Set targets = TargetRows(lo, "", usedSel, requireSelection:=True)
    If targets.Count = 0 Then MsgBox "Select the rows to reset first.", vbInformation: Exit Sub
    reason = InputBox("Reset " & targets.Count & " row(s) to '" & ST_NOT_CREATED & "'." & vbLf & _
                      "Old files are renamed *_VOID_. Reason (for the log):", "Reset", "Data corrected")
    If Len(reason) = 0 Then Exit Sub
    BeginBatch "Resetting..."
    For Each ix In targets
        ResetStatus lo, ix, reason
    Next ix
    EndBatch
    Exit Sub
Fail:
    EndBatch
    MsgBox "Stopped: " & Err.Description, vbExclamation, "Reset"
End Sub

'------------------------------------------------------------------- demo -----
' Simulates the e-signing service returning signed PDFs: copies the generated
' PDF of rows in 'Sent for signing' into output\Signed as *_signed.pdf.
' With no selection, about 70% of the rows "get signed", to make the demo realistic.
Public Sub DemoSimulateSigning()
    Dim lo As ListObject, s As Object, targets As Collection, usedSel As Boolean, ix As Variant
    Dim src As String, dest As String, n As Long
    On Error GoTo Fail
    Set lo = Agreements(): Set s = LoadSettings()
    Set targets = TargetRows(lo, ST_SENT, usedSel)
    EnsureFolder OutputFolder(s, "Signed")
    Randomize
    For Each ix In targets
        src = FldText(lo, ix, "DocumentPath")
        If FldText(lo, ix, "Status") = ST_SENT And FileExists(src) And (usedSel Or Rnd() < 0.7) Then
            dest = OutputFolder(s, "Signed") & Replace(FileNameOf(src), ".pdf", "_signed.pdf")
            Fso().CopyFile src, dest, True
            n = n + 1
        End If
    Next ix
    MsgBox n & " signed PDF(s) dropped in " & OutputFolder(s, "Signed") & vbLf & _
           "Now click 'Import signed'.", vbInformation, "Demo"
    Exit Sub
Fail:
    MsgBox "Demo stopped: " & Err.Description, vbExclamation, "Demo"
End Sub

'---------------------------------------------------------------- helpers -----
' True when the file name contains the ID and the ID isn't the start of a
' longer one (AGR-2026-001 must not match AGR-2026-0011).
Private Function NameHasId(ByVal fileName As String, ByVal id As String) As Boolean
    Dim p As Long, nextCh As String
    If Len(id) = 0 Then Exit Function
    p = InStr(1, fileName, id, vbTextCompare)
    If p = 0 Then Exit Function
    nextCh = Mid$(fileName, p + Len(id), 1)
    NameHasId = Not (nextCh Like "#")
End Function

Private Sub Summary(ByVal title As String, ByVal usedSel As Boolean, ByVal ok As Long, _
                    ByVal failed As Long, ByVal skipped As Long, ByVal verb As String)
    Dim m As String
    m = IIf(usedSel, "Selected rows", "All eligible rows") & vbLf & vbLf & _
        ok & " " & verb & vbLf
    If failed > 0 Then m = m & failed & " failed (see LastError and the Log sheet)" & vbLf
    If skipped > 0 Then m = m & skipped & " skipped (wrong status for this step)" & vbLf
    MsgBox m, IIf(failed > 0, vbExclamation, vbInformation), title
End Sub
