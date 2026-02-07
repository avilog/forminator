Option Explicit

' =======================
' CONFIG
' =======================
Private Const API_BASE As String = "http://127.0.0.1:8000"
Private Const AUDIT_URL As String = API_BASE & "/audit"
Private Const FILL_URL As String = API_BASE & "/fill"
Private Const REVISE_URL As String = API_BASE & "/revise"

Private Const DEEP_CONTEXT As Boolean = True            ' send temp text snapshot path
Private Const APPLY_REASONS_COMMENTS As Boolean = True  ' add "Reason:" comments after fills
Private Const FIX_FROM_COMMENTS As Boolean = True       ' apply user comments as revision instructions after fill

' =======================
' ENTRY POINT
' =======================
Public Sub FillForm_Tracked_ByPython()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim prevTrack As Boolean
    prevTrack = doc.TrackRevisions
    doc.TrackRevisions = True

    On Error GoTo CleanFail

    ' 1) temp text snapshot for deep context
    Dim tempTextPath As String
    tempTextPath = SaveTempTextSnapshot(doc)

    ' 2) build payload
    Dim payload As String
    payload = BuildPayload(doc, tempTextPath, DEEP_CONTEXT)

    ' 3) audit
    Dim auditResp As String
    auditResp = HttpPostJson(AUDIT_URL, payload)

    Dim auditSummary As String
    If Not ParseAuditResponse(auditResp, auditSummary) Then
        MsgBox "Audit failed. Fix extraction or template before filling." & vbCrLf & vbCrLf & auditSummary, vbExclamation, "Form filler"
        GoTo CleanExit
    End If

    If Len(auditSummary) > 0 Then
        MsgBox auditSummary, vbInformation, "Form filler audit"
    End If

    ' 4) fill
    Dim fillResp As String
    fillResp = HttpPostJson(FILL_URL, payload)

    ApplyFillPlan doc, fillResp, APPLY_REASONS_COMMENTS

    ' 5) fix stage from instruction comments
    If FIX_FROM_COMMENTS Then
        FixTextByInstructionComments doc
    End If

    ' cleanup snapshot
    On Error Resume Next
    Kill tempTextPath
    On Error GoTo 0

CleanExit:
    doc.TrackRevisions = prevTrack
    Exit Sub

CleanFail:
    doc.TrackRevisions = prevTrack
    MsgBox "Fill failed: " & Err.Description, vbExclamation, "Form filler"
End Sub

' =======================
' ENTRY POINT: FIX-ONLY
' =======================
' Standalone entry point so users can run ONLY the fix stage
' (FIX: / INSTR: / EDIT: comments) without a full fill pass.
' Assign to a Quick Access button or run via Alt+F8.
Public Sub FixByComments()
    Dim doc As Document
    Set doc = ActiveDocument

    Dim prevTrack As Boolean
    prevTrack = doc.TrackRevisions
    doc.TrackRevisions = True

    On Error GoTo FixOnlyFail

    FixTextByInstructionComments doc

    doc.TrackRevisions = prevTrack
    Exit Sub

FixOnlyFail:
    doc.TrackRevisions = prevTrack
    MsgBox "Fix-by-comments failed: " & Err.Description, vbExclamation, "Form filler"
End Sub

' ==========================================================
' BUILD PAYLOAD
' ==========================================================
Private Function BuildPayload(ByVal doc As Document, ByVal tempTextPath As String, ByVal deep As Boolean) As String
    Dim sb As String
    sb = "{"
    sb = sb & """doc_name"":""" & JsonEsc(doc.Name) & ""","
    sb = sb & """doc_folder"":""" & JsonEsc(doc.Path) & ""","
    sb = sb & """temp_text_path"":""" & JsonEsc(tempTextPath) & ""","
    sb = sb & """deep_context"":" & LCase$(CStr(deep)) & ","
    sb = sb & """content_controls"":" & CollectContentControlsJson(doc) & ","
    sb = sb & """placeholders"":" & CollectPlaceholdersJson(doc) & ","
    sb = sb & """underscore_runs"":" & CollectUnderscoreRunsJson(doc) & ","
    sb = sb & """checkbox_groups"":" & CollectCheckboxGroupsJson(doc)
    sb = sb & "}"
    BuildPayload = sb
End Function

Private Function CollectContentControlsJson(ByVal doc As Document) As String
    Dim parts As String: parts = "["
    Dim first As Boolean: first = True

    Dim cc As ContentControl
    For Each cc In doc.ContentControls
        If Not first Then parts = parts & ","
        first = False

        Dim ctx As String
        ctx = GetContextAroundRange(cc.Range, 200, doc)

        ' Current text / placeholder text
        Dim curText As String
        curText = ""
        On Error Resume Next
        curText = cc.Range.Text
        On Error GoTo 0
        If cc.ShowingPlaceholderText Then curText = "[placeholder] " & curText

        ' Dropdown / combobox options
        Dim opts As String: opts = ""
        If cc.Type = wdContentControlDropdownList Or cc.Type = wdContentControlComboBox Then
            Dim de As ContentControlListEntry
            Dim optFirst As Boolean: optFirst = True
            For Each de In cc.DropdownListEntries
                If Not optFirst Then opts = opts & ", "
                optFirst = False
                opts = opts & de.Text
            Next de
        End If

        parts = parts & "{"
        parts = parts & """id"":" & CStr(cc.ID) & ","
        parts = parts & """tag"":""" & JsonEsc(cc.Tag) & ""","
        parts = parts & """title"":""" & JsonEsc(cc.Title) & ""","
        parts = parts & """cc_type"":""" & JsonEsc(ContentControlTypeName(cc.Type)) & ""","
        parts = parts & """current_text"":""" & JsonEsc(Clip(curText, 200)) & ""","
        parts = parts & """options"":""" & JsonEsc(Clip(opts, 500)) & ""","
        parts = parts & """context"":""" & JsonEsc(Clip(ctx, 600)) & """"
        parts = parts & "}"
    Next cc

    parts = parts & "]"
    CollectContentControlsJson = parts
End Function

Private Function CollectPlaceholdersJson(ByVal doc As Document) As String
    Dim r As Range
    Set r = doc.StoryRanges(wdMainTextStory)

    Dim parts As String: parts = "["
    Dim first As Boolean: first = True
    Dim occ As Long: occ = 0

    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "[<]{2}[!<>]@[>]{2}"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
    End With

    Do While r.Find.Execute
        occ = occ + 1

        Dim token As String: token = r.Text
        Dim key As String: key = ExtractPlaceholderKey(token)
        Dim ctx As String: ctx = GetContextAroundRange(r, 200, doc)

        If Not first Then parts = parts & ","
        first = False

        parts = parts & "{"
        parts = parts & """token"":""" & JsonEsc(token) & ""","
        parts = parts & """key"":""" & JsonEsc(key) & ""","
        parts = parts & """occurrence"":" & CStr(occ) & ","
        parts = parts & """context"":""" & JsonEsc(Clip(ctx, 600)) & """"
        parts = parts & "}"

        r.Collapse wdCollapseEnd
    Loop

    parts = parts & "]"
    CollectPlaceholdersJson = parts
End Function

' ==========================================================
' UNDERSCORES (group contiguous underscore runs)
' ==========================================================
Private Function CollectUnderscoreRunsJson(ByVal doc As Document) As String
    Dim scan As Range
    Set scan = doc.StoryRanges(wdMainTextStory)

    Dim parts As String: parts = "["
    Dim first As Boolean: first = True
    Dim occ As Long: occ = 0

    Dim inGroup As Boolean: inGroup = False
    Dim groupStart As Long: groupStart = -1
    Dim groupEnd As Long: groupEnd = -1

    With scan.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "_{4,}"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
    End With

    Do While scan.Find.Execute
        If Not inGroup Then
            inGroup = True
            groupStart = scan.Start
            groupEnd = scan.End
        Else
            If scan.Start = groupEnd Then
                groupEnd = scan.End
            Else
                occ = occ + 1
                If Not first Then parts = parts & ","
                first = False

                Dim grp As Range
                Set grp = doc.Range(Start:=groupStart, End:=groupEnd)
                Dim ctx As String
                ctx = GetContextAroundRange(grp, 200, doc)

                parts = parts & "{""occurrence"":" & CStr(occ) & ",""context"":""" & JsonEsc(Clip(ctx, 600)) & """}"

                groupStart = scan.Start
                groupEnd = scan.End
            End If
        End If

        scan.Collapse wdCollapseEnd
    Loop

    If inGroup Then
        occ = occ + 1
        If Not first Then parts = parts & ","
        first = False

        Dim grpLast As Range
        Set grpLast = doc.Range(Start:=groupStart, End:=groupEnd)
        Dim ctxLast As String
        ctxLast = GetContextAroundRange(grpLast, 200, doc)

        parts = parts & "{""occurrence"":" & CStr(occ) & ",""context"":""" & JsonEsc(Clip(ctxLast, 600)) & """}"
    End If

    parts = parts & "]"
    CollectUnderscoreRunsJson = parts
End Function

' ==========================================================
' CHECKBOX GROUPS: group per paragraph containing "[ ]"
' ==========================================================
Private Function CollectCheckboxGroupsJson(ByVal doc As Document) As String
    Dim story As Range
    Set story = doc.StoryRanges(wdMainTextStory)

    Dim scan As Range
    Set scan = story.Duplicate

    Dim parts As String: parts = "["
    Dim first As Boolean: first = True

    Dim groupOcc As Long: groupOcc = 0

    Dim curParaStart As Long: curParaStart = -1
    Dim curParaEnd As Long
    Dim curParaText As String
    Dim boxesJson As String
    Dim boxCount As Long

    With scan.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "[ ]"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With

    Do While scan.Find.Execute
        Dim p As Range
        Set p = scan.Paragraphs(1).Range

        Dim pStart As Long, pEnd As Long
        pStart = p.Start
        pEnd = p.End

        If curParaStart = -1 Then
            curParaStart = pStart
            curParaEnd = pEnd
            curParaText = p.Text
            boxesJson = "["
            boxCount = 0
        ElseIf pStart <> curParaStart Then
            If boxCount > 0 Then
                boxesJson = boxesJson & "]"
                groupOcc = groupOcc + 1

                If Not first Then parts = parts & ","
                first = False

                Dim ctx As String
                ctx = GetContextAroundRange(doc.Range(curParaStart, curParaEnd), 200, doc)

                parts = parts & "{""occurrence"":" & CStr(groupOcc) & _
                                ",""text"":""" & JsonEsc(Clip(curParaText, 800)) & """" & _
                                ",""boxes"":" & boxesJson & _
                                ",""context"":""" & JsonEsc(Clip(ctx, 600)) & """}"
            End If

            curParaStart = pStart
            curParaEnd = pEnd
            curParaText = p.Text
            boxesJson = "["
            boxCount = 0
        End If

        boxCount = boxCount + 1
        If boxCount > 1 Then boxesJson = boxesJson & ","

        Dim off As Long
        off = scan.Start - curParaStart

        ' Extract label: text from end of "[ ]" to next "[ ]" or end of paragraph
        Dim lblStart As Long: lblStart = scan.End
        Dim lblEnd As Long: lblEnd = pEnd - 1   ' exclude paragraph mark
        Dim lblRange As Range
        Set lblRange = doc.Range(lblStart, lblEnd)
        Dim lblText As String: lblText = ""
        ' Look for next "[ ]" in remaining para text to trim label
        Dim nextBox As Long
        nextBox = InStr(lblRange.Text, "[ ]")
        If nextBox > 0 Then
            lblText = Trim$(Left$(lblRange.Text, nextBox - 1))
        Else
            lblText = Trim$(lblRange.Text)
        End If
        ' Clean up common separators
        If Right$(lblText, 1) = "," Or Right$(lblText, 1) = ";" Then lblText = Trim$(Left$(lblText, Len(lblText) - 1))

        boxesJson = boxesJson & "{""index"":" & CStr(boxCount) & ",""offset"":" & CStr(off) & ",""label"":""" & JsonEsc(Clip(lblText, 100)) & """}"

        scan.Collapse wdCollapseEnd
    Loop

    If curParaStart <> -1 And boxCount > 0 Then
        boxesJson = boxesJson & "]"
        groupOcc = groupOcc + 1

        If Not first Then parts = parts & ","
        first = False

        Dim ctxLast As String
        ctxLast = GetContextAroundRange(doc.Range(curParaStart, curParaEnd), 200, doc)

        parts = parts & "{""occurrence"":" & CStr(groupOcc) & _
                        ",""text"":""" & JsonEsc(Clip(curParaText, 800)) & """" & _
                        ",""boxes"":" & boxesJson & _
                        ",""context"":""" & JsonEsc(Clip(ctxLast, 600)) & """}"
    End If

    parts = parts & "]"
    CollectCheckboxGroupsJson = parts
End Function

' ==========================================================
' APPLY FILL PLAN
' Lines:  CC|id|val|reason   PH|token|val|reason
'         US|occ|val|reason  CB|group|indicesCsv|reason
' ==========================================================
Private Sub ApplyFillPlan(ByVal doc As Document, ByVal plan As String, ByVal addReasons As Boolean)
    Dim lines() As String
    lines = Split(NormLF(plan), vbLf)

    Dim dictCC As Object: Set dictCC = CreateObject("Scripting.Dictionary")
    Dim dictCCWhy As Object: Set dictCCWhy = CreateObject("Scripting.Dictionary")

    Dim dictPH As Object: Set dictPH = CreateObject("Scripting.Dictionary")
    Dim dictPHWhy As Object: Set dictPHWhy = CreateObject("Scripting.Dictionary")

    Dim dictUS As Object: Set dictUS = CreateObject("Scripting.Dictionary")
    Dim dictUSWhy As Object: Set dictUSWhy = CreateObject("Scripting.Dictionary")

    Dim dictCB As Object: Set dictCB = CreateObject("Scripting.Dictionary")
    Dim dictCBWhy As Object: Set dictCBWhy = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim$(lines(i))
        If Len(ln) = 0 Then GoTo NextLine

        Dim parts() As String
        parts = Split(ln, "|")
        If UBound(parts) < 2 Then GoTo NextLine

        Dim kind As String
        kind = parts(0)

        If kind = "CC" Then
            dictCC(parts(1)) = UrlDec(parts(2))
            If UBound(parts) >= 3 Then dictCCWhy(parts(1)) = UrlDec(parts(3))

        ElseIf kind = "PH" Then
            Dim tok As String
            tok = UrlDec(parts(1))
            dictPH(tok) = UrlDec(parts(2))
            If UBound(parts) >= 3 Then dictPHWhy(tok) = UrlDec(parts(3))

        ElseIf kind = "US" Then
            dictUS(parts(1)) = UrlDec(parts(2))
            If UBound(parts) >= 3 Then dictUSWhy(parts(1)) = UrlDec(parts(3))

        ElseIf kind = "CB" Then
            dictCB(parts(1)) = UrlDec(parts(2))
            If UBound(parts) >= 3 Then dictCBWhy(parts(1)) = UrlDec(parts(3))
        End If

NextLine:
    Next i

    ' ── content controls ──
    Dim cc As ContentControl
    For Each cc In doc.ContentControls
        Dim idKey As String
        idKey = CStr(cc.ID)

        If dictCC.Exists(idKey) Then
            Dim ccStart As Long
            ccStart = cc.Range.Start

            Dim ccVal As String
            ccVal = CStr(dictCC(idKey))
            cc.Range.Text = ccVal

            If addReasons And dictCCWhy.Exists(idKey) Then
                AddCommentSafe doc, doc.Range(ccStart, ccStart + Len(ccVal)), _
                    "Reason: " & Clip(CStr(dictCCWhy(idKey)), 220)
            End If
        End If
    Next cc

    ' ── placeholders ──
    Dim phTok As Variant
    For Each phTok In dictPH.Keys
        ReplaceAll doc, CStr(phTok), CStr(dictPH(phTok))

        If addReasons And dictPHWhy.Exists(CStr(phTok)) Then
            Dim rrPH As Range
            Set rrPH = FindFirst(doc.StoryRanges(wdMainTextStory), CStr(dictPH(phTok)))
            If Not rrPH Is Nothing Then
                AddCommentSafe doc, rrPH, "Reason: " & Clip(CStr(dictPHWhy(CStr(phTok))), 220)
            End If
        End If
    Next phTok

    ' ── underscores ──
    If dictUS.Count > 0 Then ReplaceUnderscoreRuns doc, dictUS, dictUSWhy, addReasons

    ' ── checkboxes ──
    If dictCB.Count > 0 Then ApplyCheckboxGroups doc, dictCB, dictCBWhy, addReasons
End Sub

Private Sub ReplaceUnderscoreRuns(ByVal doc As Document, ByVal dictUS As Object, ByVal dictUSWhy As Object, ByVal addReasons As Boolean)
    Dim scan As Range
    Set scan = doc.StoryRanges(wdMainTextStory)

    Dim occ As Long: occ = 0
    Dim inGroup As Boolean: inGroup = False
    Dim groupStart As Long: groupStart = -1
    Dim groupEnd As Long: groupEnd = -1

    With scan.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "_{4,}"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
    End With

    Do While scan.Find.Execute
        If Not inGroup Then
            inGroup = True
            groupStart = scan.Start
            groupEnd = scan.End
        Else
            If scan.Start = groupEnd Then
                groupEnd = scan.End
            Else
                occ = occ + 1
                ApplyOneUnderscore doc, dictUS, dictUSWhy, addReasons, occ, groupStart, groupEnd
                groupStart = scan.Start
                groupEnd = scan.End
            End If
        End If

        scan.Collapse wdCollapseEnd
    Loop

    If inGroup Then
        occ = occ + 1
        ApplyOneUnderscore doc, dictUS, dictUSWhy, addReasons, occ, groupStart, groupEnd
    End If
End Sub

Private Sub ApplyOneUnderscore(ByVal doc As Document, ByVal dictUS As Object, ByVal dictUSWhy As Object, ByVal addReasons As Boolean, ByVal occ As Long, ByVal gStart As Long, ByVal gEnd As Long)
    Dim key As String: key = CStr(occ)
    If Not dictUS.Exists(key) Then Exit Sub

    Dim grp As Range
    Set grp = doc.Range(Start:=gStart, End:=gEnd)

    Dim startPos As Long: startPos = grp.Start

    Dim val As String
    val = CStr(dictUS(key))
    grp.Text = val

    If addReasons And dictUSWhy.Exists(key) Then
        AddCommentSafe doc, doc.Range(startPos, startPos + Len(val)), _
            "Reason: " & Clip(CStr(dictUSWhy(key)), 220)
    End If
End Sub

Private Sub ApplyCheckboxGroups(ByVal doc As Document, ByVal dictCB As Object, ByVal dictCBWhy As Object, ByVal addReasons As Boolean)
    Dim story As Range
    Set story = doc.StoryRanges(wdMainTextStory)

    Dim scan As Range
    Set scan = story.Duplicate

    Dim groupOcc As Long: groupOcc = 0
    Dim curParaStart As Long: curParaStart = -1
    Dim boxOffsets() As Long
    Dim boxCount As Long

    With scan.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "[ ]"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With

    Do While scan.Find.Execute
        Dim p As Range
        Set p = scan.Paragraphs(1).Range

        Dim pStart As Long
        pStart = p.Start

        If curParaStart = -1 Then
            curParaStart = pStart
            boxCount = 0
            ReDim boxOffsets(1 To 1)
        ElseIf pStart <> curParaStart Then
            If boxCount > 0 Then
                groupOcc = groupOcc + 1
                ApplyOneCheckboxGroup doc, dictCB, dictCBWhy, addReasons, groupOcc, curParaStart, boxOffsets, boxCount
            End If

            curParaStart = pStart
            boxCount = 0
            ReDim boxOffsets(1 To 1)
        End If

        boxCount = boxCount + 1
        If boxCount > UBound(boxOffsets) Then ReDim Preserve boxOffsets(1 To boxCount)

        boxOffsets(boxCount) = scan.Start - curParaStart
        scan.Collapse wdCollapseEnd
    Loop

    If curParaStart <> -1 And boxCount > 0 Then
        groupOcc = groupOcc + 1
        ApplyOneCheckboxGroup doc, dictCB, dictCBWhy, addReasons, groupOcc, curParaStart, boxOffsets, boxCount
    End If
End Sub

Private Sub ApplyOneCheckboxGroup(ByVal doc As Document, ByVal dictCB As Object, ByVal dictCBWhy As Object, ByVal addReasons As Boolean, ByVal groupOcc As Long, ByVal paraStart As Long, ByRef boxOffsets() As Long, ByVal boxCount As Long)
    Dim key As String: key = CStr(groupOcc)
    If Not dictCB.Exists(key) Then Exit Sub

    Dim indicesCsv As String
    indicesCsv = CStr(dictCB(key))

    Dim idxParts() As String
    idxParts = Split(indicesCsv, ",")

    Dim firstMarkedStart As Long: firstMarkedStart = -1

    Dim i As Long
    For i = LBound(idxParts) To UBound(idxParts)
        Dim idx As Long
        idx = CLng(Trim$(idxParts(i)))

        If idx >= 1 And idx <= boxCount Then
            Dim off As Long
            off = boxOffsets(idx)

            Dim r As Range
            Set r = doc.Range(Start:=paraStart + off, End:=paraStart + off + 3)

            If r.Text = "[ ]" Then
                r.Text = "[X]"
                If firstMarkedStart = -1 Then firstMarkedStart = r.Start
            End If
        End If
    Next i

    If addReasons And firstMarkedStart <> -1 And dictCBWhy.Exists(key) Then
        AddCommentSafe doc, doc.Range(firstMarkedStart, firstMarkedStart + 3), _
            "Reason: " & Clip(CStr(dictCBWhy(key)), 220)
    End If
End Sub

' ==========================================================
' AUDIT RESPONSE PARSER
' ==========================================================
Private Function ParseAuditResponse(ByVal resp As String, ByRef summary As String) As Boolean
    Dim ln As String
    ln = Trim$(NormLF(resp))
    summary = ""

    Dim parts() As String
    parts = Split(ln, "|")
    If UBound(parts) < 2 Then
        ParseAuditResponse = False
        summary = "Invalid audit response format."
        Exit Function
    End If

    If parts(0) <> "AUDIT" Then
        ParseAuditResponse = False
        summary = "Unexpected audit response."
        Exit Function
    End If

    summary = UrlDec(parts(2))
    ParseAuditResponse = (parts(1) = "OK")
End Function

' ==========================================================
' FIX TEXT BY INSTRUCTION COMMENTS  (tracked)
'
' Any user comment is treated as a revision instruction.
' AI-generated comments (Reason: / Applied fix.) are skipped.
'
' Comments are collected first, then processed BACK-TO-FRONT
' so that character-position shifts from earlier replacements
' never corrupt later ones.
' ==========================================================
Public Sub FixTextByInstructionComments(ByVal doc As Document)
    Dim prevTrack As Boolean
    prevTrack = doc.TrackRevisions
    doc.TrackRevisions = True

    On Error GoTo CleanFail

    Dim commentCount As Long
    commentCount = doc.Comments.Count
    If commentCount = 0 Then GoTo CleanExit

    ' ── 1. Collect instruction comments ──
    Dim instrTexts() As String
    Dim scopeStarts() As Long
    Dim scopeEnds() As Long
    Dim selTexts() As String
    Dim instrCount As Long: instrCount = 0

    ReDim instrTexts(1 To commentCount)
    ReDim scopeStarts(1 To commentCount)
    ReDim scopeEnds(1 To commentCount)
    ReDim selTexts(1 To commentCount)

    Dim i As Long
    For i = 1 To commentCount
        Dim rawComment As String
        rawComment = doc.Comments(i).Range.Text

        Dim instrText As String
        instrText = ExtractInstruction(rawComment)

        If Len(instrText) > 0 Then
            instrCount = instrCount + 1
            instrTexts(instrCount) = instrText
            scopeStarts(instrCount) = doc.Comments(i).Scope.Start
            scopeEnds(instrCount) = doc.Comments(i).Scope.End
            selTexts(instrCount) = Clip(doc.Comments(i).Scope.Text, 5000)
        End If
    Next i

    If instrCount = 0 Then GoTo CleanExit

    ' ── 2. Process back-to-front ──
    Dim j As Long
    For j = instrCount To 1 Step -1
        Dim payload As String
        payload = BuildRevisePayload(doc.Name, doc.Path, instrTexts(j), selTexts(j))

        Dim resp As String
        resp = HttpPostJson(REVISE_URL, payload)

        Dim revised As String
        Dim reason As String

        If ParseReviseResponse(resp, revised, reason) Then
            Dim target As Range
            Set target = doc.Range(scopeStarts(j), scopeEnds(j))

            ' Trim trailing paragraph marks and cell markers that Word
            ' refuses to delete (causes "the range cannot be deleted").
            Do While target.End > target.Start
                Dim lastCh As Range
                Set lastCh = doc.Range(target.End - 1, target.End)
                If lastCh.Text = vbCr Or lastCh.Text = Chr$(7) Then
                    target.End = target.End - 1
                Else
                    Exit Do
                End If
            Loop

            If target.End = target.Start Then GoTo NextComment

            Dim startPos As Long
            startPos = target.Start

            target.Text = revised

            Dim applied As Range
            Set applied = doc.Range(startPos, startPos + Len(revised))
            AddCommentSafe doc, applied, "Applied fix. " & Clip(reason, 220)
        End If
NextComment:
    Next j

CleanExit:
    doc.TrackRevisions = prevTrack
    Exit Sub

CleanFail:
    doc.TrackRevisions = prevTrack
    MsgBox "Fix-by-comments failed: " & Err.Description, vbExclamation, "Form filler"
End Sub

Private Function ExtractInstruction(ByVal commentText As String) As String
    ' Return the comment text as an instruction UNLESS it was generated
    ' by this macro (Reason: / Applied fix.). Any normal user comment
    ' is treated as a revision instruction.
    Dim t As String
    t = Trim$(commentText)

    If Len(t) = 0 Then
        ExtractInstruction = vbNullString
        Exit Function
    End If

    ' Skip AI-generated comments
    Dim u As String: u = UCase$(t)
    If Left$(u, 7) = "REASON:" Then ExtractInstruction = vbNullString: Exit Function
    If Left$(u, 12) = "APPLIED FIX." Then ExtractInstruction = vbNullString: Exit Function

    ' Strip optional FIX:/INSTR:/EDIT: prefix if user still uses them
    If Left$(u, 4) = "FIX:" Then t = Trim$(Mid$(t, 5))
    If Left$(u, 6) = "INSTR:" Then t = Trim$(Mid$(t, 7))
    If Left$(u, 5) = "EDIT:" Then t = Trim$(Mid$(t, 6))

    ExtractInstruction = t
End Function

Private Function BuildRevisePayload(ByVal docName As String, ByVal docFolder As String, ByVal instruction As String, ByVal selectedText As String) As String
    BuildRevisePayload = "{" & _
        """doc_name"":""" & JsonEsc(docName) & """," & _
        """doc_folder"":""" & JsonEsc(docFolder) & """," & _
        """instruction"":""" & JsonEsc(instruction) & """," & _
        """selected_text"":""" & JsonEsc(selectedText) & """" & _
    "}"
End Function

Private Function ParseReviseResponse(ByVal resp As String, ByRef revised As String, ByRef reason As String) As Boolean
    Dim ln As String
    ln = Trim$(NormLF(resp))
    If Len(ln) = 0 Then
        ParseReviseResponse = False
        Exit Function
    End If

    Dim parts() As String
    parts = Split(ln, "|")
    If UBound(parts) < 2 Then
        ParseReviseResponse = False
        Exit Function
    End If
    If parts(0) <> "REV" Then
        ParseReviseResponse = False
        Exit Function
    End If

    revised = UrlDec(parts(1))
    reason = UrlDec(parts(2))
    ParseReviseResponse = True
End Function

' ==========================================================
' SNAPSHOT (UTF-8 text)
' ==========================================================
Private Function SaveTempTextSnapshot(ByVal doc As Document) As String
    Dim tempFolder As String
    tempFolder = Environ$("TEMP")
    If Right$(tempFolder, 1) <> "\" Then tempFolder = tempFolder & "\"

    Dim fileName As String
    fileName = "word_text_snapshot_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim fullPath As String
    fullPath = tempFolder & fileName

    Dim snapshot As String
    snapshot = Clip(doc.Content.Text, 50000)

    WriteUtf8Text fullPath, snapshot
    SaveTempTextSnapshot = fullPath
End Function

Private Sub WriteUtf8Text(ByVal path As String, ByVal text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2
    stm.Close
End Sub

' ==========================================================
' HTTP  (timeout raised for LLM round-trips)
' ==========================================================
Private Function HttpPostJson(ByVal url As String, ByVal body As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' resolve, connect, send, receive (ms)
    http.SetTimeouts 10000, 10000, 30000, 120000

    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.Send body

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 513, "HttpPostJson", "HTTP " & http.Status & ": " & http.ResponseText
    End If

    HttpPostJson = http.ResponseText
End Function

' ==========================================================
' COMMENTS
' ==========================================================
Private Sub AddCommentSafe(ByVal doc As Document, ByVal r As Range, ByVal text As String)
    On Error Resume Next

    ' Temporarily set author to Forminator Bot
    Dim savedName As String: savedName = Application.UserName
    Dim savedInit As String: savedInit = Application.UserInitials
    Application.UserName = "Forminator Bot"
    Application.UserInitials = "FB"

    doc.Comments.Add Range:=r, text:=text

    Application.UserName = savedName
    Application.UserInitials = savedInit

    On Error GoTo 0
End Sub

' ==========================================================
' REPLACE HELPERS
' ==========================================================
Private Sub ReplaceAll(ByVal doc As Document, ByVal findText As String, ByVal replaceText As String)
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Function FindFirst(ByVal withinRange As Range, ByVal needle As String) As Range
    Dim r As Range
    Set r = withinRange.Duplicate

    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = needle
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With

    If r.Find.Execute Then
        Set FindFirst = r.Duplicate
    Else
        Set FindFirst = Nothing
    End If
End Function

' ==========================================================
' UTILITIES
' ==========================================================
Private Function ContentControlTypeName(ByVal t As Long) As String
    Select Case t
        Case wdContentControlRichText: ContentControlTypeName = "rich_text"
        Case wdContentControlText: ContentControlTypeName = "text"
        Case wdContentControlDate: ContentControlTypeName = "date"
        Case wdContentControlDropdownList: ContentControlTypeName = "dropdown"
        Case wdContentControlComboBox: ContentControlTypeName = "combobox"
        Case 8: ContentControlTypeName = "checkbox"  ' wdContentControlCheckBox (Word 2010+)
        Case Else: ContentControlTypeName = "other"
    End Select
End Function

Private Function ExtractPlaceholderKey(ByVal token As String) As String
    Dim s As String
    s = Trim$(token)
    If Left$(s, 2) = "<<" And Right$(s, 2) = ">>" Then
        ExtractPlaceholderKey = Trim$(Mid$(s, 3, Len(s) - 4))
    Else
        ExtractPlaceholderKey = s
    End If
End Function

Private Function GetContextAroundRange(ByVal r As Range, ByVal charsEachSide As Long, ByVal doc As Document) As String
    Dim s As Long, e As Long
    Dim docEnd As Long: docEnd = doc.Content.End

    s = r.Start - charsEachSide
    If s < 0 Then s = 0

    e = r.End + charsEachSide
    If e > docEnd Then e = docEnd

    ' ── snap start forward to first space so we don't clip a word ──
    If s > 0 Then
        Dim probe As Range
        Set probe = doc.Range(Start:=s, End:=s + 1)
        ' walk forward until we hit a space (max 30 chars to avoid runaway)
        Dim guard As Long: guard = 0
        Do While s < r.Start And guard < 30
            Set probe = doc.Range(Start:=s, End:=s + 1)
            If probe.Text = " " Or probe.Text = vbCr Or probe.Text = vbLf Then
                s = s + 1   ' skip the space itself
                Exit Do
            End If
            s = s + 1
            guard = guard + 1
        Loop
    End If

    ' ── snap end backward to last space so we don't clip a word ──
    If e < docEnd Then
        Dim guard2 As Long: guard2 = 0
        Do While e > r.End And guard2 < 30
            Dim ch As Range
            Set ch = doc.Range(Start:=e - 1, End:=e)
            If ch.Text = " " Or ch.Text = vbCr Or ch.Text = vbLf Then
                e = e - 1   ' exclude the trailing space
                Exit Do
            End If
            e = e - 1
            guard2 = guard2 + 1
        Loop
    End If

    Dim ctx As Range
    Set ctx = doc.Range(Start:=s, End:=e)
    GetContextAroundRange = ctx.Text
End Function

Private Function Clip(ByVal s As String, ByVal maxLen As Long) As String
    If Len(s) > maxLen Then Clip = Left$(s, maxLen) Else Clip = s
End Function

Private Function NormLF(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    NormLF = s
End Function

Private Function UrlDec(ByVal s As String) As String
    Dim i As Long, out As String, c As String
    out = ""
    i = 1
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If c = "%" And i + 2 <= Len(s) Then
            out = out & ChrW(CLng("&H" & Mid$(s, i + 1, 2)))
            i = i + 3
        Else
            out = out & c
            i = i + 1
        End If
    Loop
    UrlDec = out
End Function

Private Function JsonEsc(ByVal s As String) As String
    Dim i As Long, ch As Integer, out As String
    out = ""
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        Select Case ch
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 8: out = out & "\b"
            Case 9: out = out & "\t"
            Case 10: out = out & "\n"
            Case 12: out = out & "\f"
            Case 13: out = out & "\r"
            Case Is < 32
                out = out & "\u00" & Right$("0" & Hex$(ch), 2)
            Case Else
                out = out & ChrW(ch)
        End Select
    Next i
    JsonEsc = out
End Function
