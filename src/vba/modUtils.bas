Attribute VB_Name = "modUtils"
'==============================================================================
' modUtils - settings, table access, logging and file helpers
' Agreement Generator (portfolio project) - github.com/Sigmund-Engeset
'==============================================================================
Option Explicit

Public Const SHEET_AGREEMENTS As String = "Agreements"
Public Const TBL_AGREEMENTS As String = "tblAgreements"
Public Const TBL_SETTINGS As String = "tblSettings"
Public Const TBL_TEMPLATES As String = "tblTemplates"
Public Const TBL_LOG As String = "tblLog"

Private mOldCalc As Long

'---------------------------------------------------------------- tables ------
Public Function Agreements() As ListObject
    Set Agreements = ThisWorkbook.Worksheets(SHEET_AGREEMENTS).ListObjects(TBL_AGREEMENTS)
End Function

Public Function FindTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set FindTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTable Is Nothing Then Exit Function
    Next ws
    Err.Raise vbObjectError + 100, "FindTable", "Table '" & tableName & "' was not found."
End Function

' Column index by header name, so the code never depends on column letters.
Public Function ColIx(ByVal lo As ListObject, ByVal header As String) As Long
    On Error Resume Next
    ColIx = lo.ListColumns(header).Index
    On Error GoTo 0
    If ColIx = 0 Then Err.Raise vbObjectError + 101, "ColIx", _
        "Column '" & header & "' is missing in table " & lo.Name & "."
End Function

Public Function Fld(ByVal lo As ListObject, ByVal rowIx As Long, ByVal header As String) As Range
    Set Fld = lo.DataBodyRange.Cells(rowIx, ColIx(lo, header))
End Function

Public Function FldText(ByVal lo As ListObject, ByVal rowIx As Long, ByVal header As String) As String
    Dim v As Variant
    v = Fld(lo, rowIx, header).Value
    If IsError(v) Or IsEmpty(v) Then FldText = "" Else FldText = Trim$(CStr(v))
End Function

Public Function RowCount(ByVal lo As ListObject) As Long
    If lo.DataBodyRange Is Nothing Then RowCount = 0 Else RowCount = lo.ListRows.Count
End Function

'-------------------------------------------------------------- settings ------
' Key/value pairs from tblSettings. Case-insensitive keys.
Public Function LoadSettings() As Object
    Dim d As Object, lo As ListObject, r As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set lo = FindTable(TBL_SETTINGS)
    For r = 1 To RowCount(lo)
        d(FldText(lo, r, "Key")) = FldText(lo, r, "Value")
    Next r
    Set LoadSettings = d
End Function

Public Function SettingYes(ByVal s As Object, ByVal key As String) As Boolean
    If s.Exists(key) Then SettingYes = (LCase$(s(key)) = "yes" Or LCase$(s(key)) = "true")
End Function

Public Function SettingOr(ByVal s As Object, ByVal key As String, ByVal fallback As String) As String
    If s.Exists(key) Then
        If Len(s(key)) > 0 Then SettingOr = s(key): Exit Function
    End If
    SettingOr = fallback
End Function

' AgreementType -> Array(TemplateFile, FilePrefix, TemplateVersion)
Public Function LoadTemplates() As Object
    Dim d As Object, lo As ListObject, r As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set lo = FindTable(TBL_TEMPLATES)
    For r = 1 To RowCount(lo)
        If Len(FldText(lo, r, "AgreementType")) > 0 Then
            d(FldText(lo, r, "AgreementType")) = Array(FldText(lo, r, "TemplateFile"), _
                FldText(lo, r, "FilePrefix"), FldText(lo, r, "TemplateVersion"))
        End If
    Next r
    Set LoadTemplates = d
End Function

'----------------------------------------------------------------- paths ------
Public Function Sep() As String
    Sep = Application.PathSeparator
End Function

Public Function WorkbookFolder() As String
    WorkbookFolder = ThisWorkbook.Path
    If LCase$(Left$(WorkbookFolder, 4)) = "http" Then
        Err.Raise vbObjectError + 110, "WorkbookFolder", _
            "This workbook is opened from OneDrive/SharePoint, so its path is a URL." & vbLf & _
            "Set TemplateFolder and OutputFolder on the Settings sheet to local folder paths."
    End If
End Function

' Blank setting -> <workbook folder>\<defaultSub>. Relative (.\x) -> under workbook folder.
Public Function ResolveFolder(ByVal settingValue As String, ByVal defaultSub As String) As String
    Dim p As String
    p = Trim$(settingValue)
    If Len(p) = 0 Then
        p = WorkbookFolder() & Sep() & defaultSub
    ElseIf Left$(p, 2) = "." & Sep() Then
        p = WorkbookFolder() & Mid$(p, 2)
    End If
    If Right$(p, 1) <> Sep() Then p = p & Sep()
    ResolveFolder = p
End Function

Public Function TemplateFolder(ByVal s As Object) As String
    TemplateFolder = ResolveFolder(SettingOr(s, "TemplateFolder", ""), "templates")
End Function

Public Function OutputFolder(ByVal s As Object, ByVal subFolder As String) As String
    OutputFolder = ResolveFolder(SettingOr(s, "OutputFolder", ""), "output") & subFolder & Sep()
End Function

' FileSystemObject instead of Dir$: Dir$ keeps global state, so calling it
' inside another Dir$ loop silently breaks that loop.
Public Function Fso() As Object
    Static f As Object
    If f Is Nothing Then Set f = CreateObject("Scripting.FileSystemObject")
    Set Fso = f
End Function

Public Function FileExists(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    FileExists = Fso().FileExists(path)
End Function

' Creates the folder and any missing parent folders (works for UNC paths too).
Public Sub EnsureFolder(ByVal path As String)
    If Right$(path, 1) = Sep() Then path = Left$(path, Len(path) - 1)
    If Len(path) = 0 Then Exit Sub
    If Fso().FolderExists(path) Then Exit Sub
    EnsureFolder Fso().GetParentFolderName(path)
    Fso().CreateFolder path
End Sub

Public Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    For Each ch In bad
        s = Replace(s, ch, "")
    Next ch
    s = Application.WorksheetFunction.Trim(s)
    SafeFileName = Replace(s, " ", "_")
End Function

Public Function FileNameOf(ByVal path As String) As String
    FileNameOf = Mid$(path, InStrRev(path, Sep()) + 1)
End Function

'--------------------------------------------------------------- logging ------
Public Sub AddLog(ByVal agreementId As String, ByVal action As String, _
                  ByVal fromStatus As String, ByVal toStatus As String, ByVal message As String)
    Dim lo As ListObject, lr As ListRow
    Set lo = FindTable(TBL_LOG)
    ' The table ships with one empty row; reuse it before adding new ones.
    If RowCount(lo) = 1 And IsEmpty(lo.DataBodyRange.Cells(1, 1).Value) Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    lr.Range.Value = Array(Now, agreementId, action, fromStatus, toStatus, CurrentUser(), message)
    lr.Range.Cells(1, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Public Function CurrentUser() As String
    CurrentUser = Environ$("USERNAME")
    If Len(CurrentUser) = 0 Then CurrentUser = Application.UserName
End Function

'------------------------------------------------------------ hyperlinks ------
Public Sub SetPathCell(ByVal cell As Range, ByVal path As String)
    cell.Hyperlinks.Delete
    cell.Value = path
    If Len(path) > 0 Then
        cell.Parent.Hyperlinks.Add Anchor:=cell, Address:=path, TextToDisplay:=path
    End If
End Sub

'----------------------------------------------------------- performance ------
Public Sub BeginBatch(ByVal statusText As String)
    mOldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = statusText
End Sub

Public Sub EndBatch()
    If mOldCalc <> 0 Then Application.Calculation = mOldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub Progress(ByVal text As String, ByVal i As Long, ByVal n As Long)
    Application.StatusBar = text & " " & i & " of " & n & "..."
    DoEvents
End Sub
