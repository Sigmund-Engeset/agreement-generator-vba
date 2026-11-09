Attribute VB_Name = "modStatus"
'==============================================================================
' modStatus - status state machine, row targeting and data validation
'==============================================================================
Option Explicit

Public Const ST_NOT_CREATED As String = "Not created"
Public Const ST_CREATED As String = "Created"
Public Const ST_SENT As String = "Sent for signing"
Public Const ST_SIGNED As String = "Signed"
Public Const ST_PROCESSED As String = "Processed"

Public Function StatusFlow() As Variant
    StatusFlow = Array(ST_NOT_CREATED, ST_CREATED, ST_SENT, ST_SIGNED, ST_PROCESSED)
End Function

Public Function StatusIndex(ByVal status As String) As Long
    Dim f As Variant, i As Long
    f = StatusFlow()
    StatusIndex = -1
    For i = LBound(f) To UBound(f)
        If StrComp(f(i), status, vbTextCompare) = 0 Then StatusIndex = i: Exit Function
    Next i
End Function

' Which timestamp column a status writes to.
Private Function StampColumn(ByVal status As String) As String
    Select Case status
        Case ST_CREATED: StampColumn = "CreatedAt"
        Case ST_SENT: StampColumn = "SentAt"
        Case ST_SIGNED: StampColumn = "SignedAt"
        Case ST_PROCESSED: StampColumn = "ProcessedAt"
    End Select
End Function

' The only way status changes. Allows exactly one step forward, stamps the
' time and writes an audit log line. Raises an error on an illegal move.
Public Sub MoveStatus(ByVal lo As ListObject, ByVal rowIx As Long, ByVal newStatus As String, _
                      ByVal action As String, Optional ByVal message As String = "")
    Dim oldStatus As String, id As String
    oldStatus = FldText(lo, rowIx, "Status")
    id = FldText(lo, rowIx, "AgreementID")
    If StatusIndex(newStatus) <> StatusIndex(oldStatus) + 1 Then
        Err.Raise vbObjectError + 200, "MoveStatus", _
            id & ": cannot move from '" & oldStatus & "' to '" & newStatus & "'."
    End If
    Fld(lo, rowIx, "Status").Value = newStatus
    With Fld(lo, rowIx, StampColumn(newStatus))
        .Value = Now
        .NumberFormat = "yyyy-mm-dd hh:mm"
    End With
    Fld(lo, rowIx, "LastError").Value = ""
    AddLog id, action, oldStatus, newStatus, message
End Sub

' Back to Not created: clears paths and timestamps. The old PDF/.docx are kept
' but renamed *_VOID_<timestamp>, so the next Generate can't be confused with them.
Public Sub ResetStatus(ByVal lo As ListObject, ByVal rowIx As Long, ByVal reason As String)
    Dim oldStatus As String, h As Variant, pdf As String, stamp As String, ext As Variant
    oldStatus = FldText(lo, rowIx, "Status")
    pdf = FldText(lo, rowIx, "DocumentPath")
    If Len(pdf) > 0 Then
        stamp = "_VOID_" & Format$(Now, "yyyymmdd-hhnnss")
        For Each ext In Array(".pdf", ".docx")
            h = Left$(pdf, Len(pdf) - 4) & ext
            If FileExists(CStr(h)) Then Fso().MoveFile CStr(h), Left$(pdf, Len(pdf) - 4) & stamp & ext
        Next ext
        reason = reason & " (old document renamed *" & stamp & ")"
    End If
    Fld(lo, rowIx, "Status").Value = ST_NOT_CREATED
    For Each h In Array("DocumentPath", "SignedDocumentPath")
        SetPathCell Fld(lo, rowIx, CStr(h)), ""
    Next h
    For Each h In Array("CreatedAt", "SentAt", "SignedAt", "ProcessedAt", "LastError")
        Fld(lo, rowIx, CStr(h)).ClearContents
    Next h
    AddLog FldText(lo, rowIx, "AgreementID"), "Reset", oldStatus, ST_NOT_CREATED, reason
End Sub

Public Sub SetRowError(ByVal lo As ListObject, ByVal rowIx As Long, ByVal action As String, ByVal message As String)
    Fld(lo, rowIx, "LastError").Value = message
    AddLog FldText(lo, rowIx, "AgreementID"), action & " failed", FldText(lo, rowIx, "Status"), "", message
End Sub

'------------------------------------------------------------- targeting ------
' Rows a button should act on:
'  - if the user has selected cells inside the table: those (visible) rows
'  - otherwise: every row whose Status = eligibleStatus
' usedSelection tells the caller which mode was used (for the summary message).
Public Function TargetRows(ByVal lo As ListObject, ByVal eligibleStatus As String, _
                           ByRef usedSelection As Boolean, _
                           Optional ByVal requireSelection As Boolean = False) As Collection
    Dim result As New Collection, sel As Range, i As Long

    If RowCount(lo) > 0 And TypeName(Selection) = "Range" Then
        If ActiveSheet Is lo.Parent Then Set sel = Intersect(Selection, lo.DataBodyRange)
    End If

    If Not sel Is Nothing Then
        usedSelection = True
        Set result = RowsFromRange(lo, sel)
    ElseIf Not requireSelection Then
        usedSelection = False
        For i = 1 To RowCount(lo)
            If StrComp(FldText(lo, i, "Status"), eligibleStatus, vbTextCompare) = 0 Then result.Add i
        Next i
    End If
    Set TargetRows = result
End Function

' Table row numbers (1-based, unique, visible rows only) covered by a range.
' Returns an empty collection when the range is not inside the table.
Public Function RowsFromRange(ByVal lo As ListObject, ByVal rng As Range) As Collection
    Dim result As New Collection, seen As Object, sel As Range, area As Range, rw As Range, ix As Long
    Set seen = CreateObject("Scripting.Dictionary")
    If RowCount(lo) > 0 And Not rng Is Nothing Then
        If rng.Worksheet Is lo.Parent Then Set sel = Intersect(rng, lo.DataBodyRange)
    End If
    If Not sel Is Nothing Then
        For Each area In sel.Areas
            For Each rw In area.Rows
                If Not rw.EntireRow.Hidden Then
                    ix = rw.Row - lo.DataBodyRange.Row + 1
                    If Not seen.Exists(ix) Then seen.Add ix, True: result.Add ix
                End If
            Next rw
        Next area
    End If
    Set RowsFromRange = result
End Function

'------------------------------------------------------------ validation ------
' Returns "" when the row is valid, otherwise a readable list of problems.
' Placeholder-level checks (blank values the template needs) happen during
' generation, because only the template knows which fields it uses.
Public Function ValidateRow(ByVal lo As ListObject, ByVal rowIx As Long, ByVal s As Object, _
                            ByVal templates As Object, ByVal idCount As Object) As String
    Dim p As String, id As String, t As String, email As String, tplPath As String

    id = FldText(lo, rowIx, "AgreementID")
    If Len(id) = 0 Then
        p = p & "AgreementID is blank. "
    ElseIf idCount(id) > 1 Then
        p = p & "AgreementID is used " & idCount(id) & " times. "
    End If

    t = FldText(lo, rowIx, "AgreementType")
    If Not templates.Exists(t) Then
        p = p & "AgreementType '" & t & "' has no row on the Templates sheet. "
    Else
        tplPath = TemplateFolder(s) & templates(t)(0)
        If Not FileExists(tplPath) Then p = p & "Template file not found: " & tplPath & ". "
    End If

    If Len(FldText(lo, rowIx, "CounterpartyName")) = 0 Then p = p & "CounterpartyName is blank. "

    email = FldText(lo, rowIx, "ContactEmail")
    If Not IsValidEmail(email) Then p = p & "ContactEmail '" & email & "' is not a valid address. "

    If Not IsDate(Fld(lo, rowIx, "StartDate").Value) Then p = p & "StartDate is not a date. "

    If SettingYes(s, "ValidateNorwegianOrgNo") And _
       StrComp(FldText(lo, rowIx, "CounterpartyCountry"), "Norway", vbTextCompare) = 0 Then
        If Not IsValidNorwegianOrgNo(FldText(lo, rowIx, "CounterpartyOrgNo")) Then
            p = p & "CounterpartyOrgNo fails the MOD11 check. "
        End If
    End If

    ValidateRow = Trim$(p)
End Function

' Count of each AgreementID, for the duplicate check.
Public Function CountIds(ByVal lo As ListObject) As Object
    Dim d As Object, i As Long, id As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For i = 1 To RowCount(lo)
        id = FldText(lo, i, "AgreementID")
        If Len(id) > 0 Then d(id) = d(id) + 1
    Next i
    Set CountIds = d
End Function

Public Function IsValidEmail(ByVal e As String) As Boolean
    Dim atPos As Long
    If Len(e) < 6 Or InStr(e, " ") > 0 Then Exit Function
    atPos = InStr(e, "@")
    If atPos < 2 Or atPos <> InStrRev(e, "@") Then Exit Function
    IsValidEmail = (Mid$(e, atPos + 1) Like "?*.?*")
End Function

' Norwegian organisation number: 9 digits, weights 3 2 7 6 5 4 3 2, MOD11.
Public Function IsValidNorwegianOrgNo(ByVal orgNo As String) As Boolean
    Dim d As String, w As Variant, i As Long, total As Long, check As Long
    d = Replace(Replace(orgNo, " ", ""), ".", "")
    If Len(d) <> 9 Or Not d Like "#########" Then Exit Function
    w = Array(3, 2, 7, 6, 5, 4, 3, 2)
    For i = 0 To 7
        total = total + CLng(Mid$(d, i + 1, 1)) * w(i)
    Next i
    check = 11 - (total Mod 11)
    If check = 11 Then check = 0
    If check = 10 Then Exit Function
    IsValidNorwegianOrgNo = (check = CLng(Right$(d, 1)))
End Function
