# Improvements to the document-generation mechanism

The original tool (2025, banking agreements) worked like this: read a row, open Word, run a fixed set of find/replace operations, export a PDF, and track status in a column. That is the right idea and it saved a lot of manual work. These are the changes that make it more robust, easier to maintain and easier to audit. Sections 1–9 are implemented in this repo; section 10 lists what I would do next.

---

## 1. Let the template drive the fields

**v1 pattern:** one replace call per field, written in the code.

```vba
ReplaceText doc, "<<CustomerName>>", ws.Cells(r, 4).Value
ReplaceText doc, "<<OrgNo>>", ws.Cells(r, 5).Value
' ...one line per field, per template
```

Adding a field meant changing the code, the template and the sheet, and keeping column numbers in sync.

**v2:** search the document for anything that looks like a placeholder, and look its name up in a dictionary built from the row's column headers and the Settings sheet.

```vba
Private Const PLACEHOLDER_PATTERN As String = "\{\{[A-Za-z0-9_]@\}\}"

Do While r.Find.Execute
    key = Mid$(r.Text, 3, Len(r.Text) - 4)
    If Not values.Exists(key) Then
        problems("Unknown placeholder {{" & key & "}}") = True
    ElseIf Len(values(key)) = 0 Then
        problems("No value for {{" & key & "}}") = True
    Else
        r.Text = values(key)
    End If
    r.Collapse wdCollapseEnd
Loop
```

A new field is now a new column plus `{{NewColumn}}` in the template. The same code serves every template, and each template only "requires" the fields it uses (an NDA doesn't need a contract value).

## 2. Never open the template itself

**v1:** `Documents.Open(templatePath)` → replace → `SaveAs`. If the macro failed between open and save-as, or someone pressed Ctrl+S, the template could be overwritten with one customer's data.

**v2:** `Documents.Add(Template:=templatePath)` creates a new, unsaved document based on the template. The template file is never opened for editing.

## 3. Replace in every part of the document

`ActiveDocument.Content.Find` only searches the main body. Headers, footers, footnotes and text boxes are separate *story ranges*. v2 loops through all of them:

```vba
For Each story In doc.StoryRanges
    Do
        FillRange story, values, problems
        Set story = story.NextStoryRange   ' linked headers/footers per section
    Loop Until story Is Nothing
Next story
```

The templates use `{{AgreementID}}` in the header and `{{TemplateVersion}}` in the footer to prove it.

## 4. No 255-character limit, formatting kept

`Find.Replacement.Text` fails on values longer than 255 characters (a long scope description, for example). v2 finds the token and sets `Range.Text` directly, which has no limit and keeps the token's formatting (bold stays bold). Line breaks typed with Alt+Enter in Excel are converted to Word line breaks.

## 5. Block incomplete documents

Before exporting, v2 knows if any placeholder was unknown or blank. In that case the PDF is **not** exported and the row gets a readable error such as `Not exported. No value for {{HourlyRate}}`. A PDF with `{{HourlyRate}}` in it can never be sent to a customer.

Validation also runs before Word is touched: blank/duplicate IDs, missing template files, e-mail format, start date, and MOD11 on Norwegian organisation numbers.

## 6. Structured data instead of cell addresses

The data is an Excel table (`tblAgreements`), and code reads columns by header name:

```vba
FldText(lo, rowIx, "CounterpartyName")   ' instead of ws.Cells(r, 4)
```

Inserting or moving a column no longer breaks the macro. Settings (company details, folders, date and number formats) are in `tblSettings`, and agreement types are mapped to templates in `tblTemplates`, so non-developers can maintain them.

## 7. A real status flow with an audit trail

v1 had a status column that anyone could type into. v2 has one function that changes status, `MoveStatus`, which:

- allows exactly one step forward (Not created → Created → Sent for signing → Signed → Processed),
- stamps the time in the matching column (`CreatedAt`, `SentAt`, `SignedAt`, `ProcessedAt`),
- writes a log line with Windows user, old and new status, and a message.

`Reset` is the only way back: it asks for a reason, logs it and renames the old PDF `*_VOID_<timestamp>` so a voided document is never mistaken for the current one. This makes it possible to answer "who sent this agreement, and when?" months later.

## 8. Robust batch runs

- **One Word instance** per batch instead of one per document, invisible, with alerts off.
- **Per-row error handling:** a bad row is logged and skipped; the rest of the batch continues.
- **Guaranteed cleanup:** Word is closed in both the normal and error paths, so no orphan `WINWORD.EXE` processes.
- Screen updating, events and calculation are paused during the run, with progress in the status bar.
- Late binding (`CreateObject("Word.Application")`), so the workbook runs without setting library references on each PC.

## 9. Closing the loop on signing

Signed documents are picked up by convention: any PDF in `output\Signed` whose name contains the AgreementID moves that row to Signed. `Process` then archives the file to `output\Archive\<year>`. PDF export can use PDF/A (`UseISO19005_1`) for long-term archiving. Optional Outlook drafts (saved, not sent) give a human a final check before anything leaves the building.

---

## 10. Next steps (not implemented)

| Idea | Why |
|---|---|
| **E-signing API** (Signicat/Posten for BankID, DocuSign, Adobe Sign) via `MSXML2.ServerXMLHTTP` | Status updates straight from the signing provider instead of a folder drop |
| **Brreg lookup** on org number | Auto-fill name and address, and stop agreements with deleted or bankrupt companies. I've built this lookup in VBA before (30,000 org numbers) |
| **Conditional clauses** with bookmarks or Building Blocks | E.g. include a data-processing clause only when `ProcessesPersonalData = Yes` |
| **Content controls** instead of text tokens | Typed fields in the template that can't be broken by partial formatting edits |
| **Template hash on the row** | Prove exactly which template version produced a document |
| **Sheet protection** with `UserInterfaceOnly:=True` | Status and audit columns editable only through the macros |
| **Power Platform** (SharePoint list + Power Automate *Populate a Microsoft Word template* + approvals) | Multi-user, no dependency on one desktop with Word |
| **Python version** (`docxtpl` + LibreOffice) | Server-side batches without Office, easier testing |
