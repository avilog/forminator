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
Private Const FIX_FROM_COMMENTS As Boolean = True       ' apply FIX:/INSTR:/EDIT: comments after fill

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

    ApplyFillPlanPlain_WithReasons doc, fillResp, APPLY_REASONS_COMMENTS

    ' 5) fix stage from instruction comments
    If FIX_FROM_COMMENTS Then
        FixTextByInstructionComments_Tracked doc
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

' ==========================================================
' BUILD PAYLOAD
' ==========================================================
Private Function BuildPayload(ByVal doc As Document, ByVal tempTextPath As String, ByVal deep As Boolean) As String
    Dim sb As String
    sb = "{"
    sb = sb & """doc_name"":""" & JsonEscapeStrict(doc.Name) & ""","
    sb = sb & """temp_text_path"":""" & JsonEscapeStrict(tempTextPath) & ""","
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
        ctx = GetContextAroundRange(cc.Range, 80, doc)

        parts = parts & "{"
        parts = parts & """id"":" & CStr(cc.ID) & ","
        parts = parts & """tag"":""" & JsonEscapeStrict(cc.Tag) & ""","
        parts = parts & """title"":""" & JsonEscapeStrict(cc.Title) & ""","
        parts = parts & """cc_type"":""" & JsonEscapeStrict(ContentControlTypeName(cc.Type)) & ""","
        parts = parts & """context"":""" & JsonEscapeStrict(LimitLen(ctx, 400)) & """"
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
        Dim ctx As String: ctx = GetContextAroundRange(r, 80, doc)

        If Not first Then parts = parts & ","
        first = False

        parts = parts & "{"
        parts = parts & """token"":""" & JsonEscapeStrict(token) & ""","
        parts = parts & """key"":""" & JsonEscapeStrict(key) & ""","
        parts = parts & """occurrence"":" & CStr(occ) & ","
        parts = parts & """context"":""" & JsonEscapeStrict(LimitLen(ctx, 400)) & """"
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
                ctx = GetContextAroundRange(grp, 80, doc)

                parts = parts & "{""occurrence"":" & CStr(occ) & ",""context"":""" & JsonEscapeStrict(LimitLen(ctx, 400)) & """}"

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
        ctxLast = GetContextAroundRange(grpLast, 80, doc)

        parts = parts & "{""occurrence"":" & CStr(occ) & ",""context"":""" & JsonEscapeStrict(LimitLen(ctxLast, 400)) & """}"
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
                ctx = GetContextAroundRange(doc.Range(curParaStart, curParaEnd), 80, doc)

                parts = parts & "{""occurrence"":" & CStr(groupOcc) & _
                                ",""text"":""" & JsonEscapeStrict(LimitLen(curParaText, 800)) & """" & _
                                ",""boxes"":" & boxesJson & _
                                ",""context"":""" & JsonEscapeStrict(LimitLen(ctx, 400)) & """}"
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

        boxesJson = boxesJson & "{""index"":" & CStr(boxCount) & ",""offset"":" & CStr(off) & ",""label"":""""}"

        scan.Collapse wdCollapseEnd
    Loop

    If curParaStart <> -1 And boxCount > 0 Then
        boxesJson = boxesJson & "]"
        groupOcc = groupOcc + 1

        If Not first Then parts = parts & ","
        first = False

        Dim ctxLast As String
        ctxLast = GetContextAroundRange(doc.Range(curParaStart, curParaEnd), 80, doc)

        parts = parts & "{""occurrence"":" & CStr(groupOcc) & _
                        ",""text"":""" & JsonEscapeStrict(LimitLen(curParaText, 800)) & """" & _
                        ",""boxes"":" & boxesJson & _
                        ",""context"":""" & JsonEscapeStrict(LimitLen(ctxLast, 400)) & """}"
    End If

    parts = parts & "]"
    CollectCheckboxGroupsJson = parts
End Function

' ==========================================================
' APPLY FILL PLAN (with reasons)
' Lines:
'   CC|id|val|reason
'   PH|token|val|reason
'   US|occ|val|reason
'   CB|group|indicesCsv|reason
' ==========================================================
Private Sub ApplyFillPlanPlain_WithReasons(ByVal doc As Document, ByVal plan As String, ByVal addReasons As Boolean)
    Dim lines() As String
    lines = Split(NormalizeNewlines(plan), vbLf)

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
        Dim line As String
        line = Trim$(lines(i))
        If Len(line) = 0 Then GoTo NextLine

        Dim parts() As String
        parts = Split(line, "|")
        If UBound(parts) < 2 Then GoTo NextLine

        Dim kind As String
        kind = parts(0)

        If kind = "CC" Then
            dictCC(parts(1)) = UrlDecode(parts(2))
            If UBound(parts) >= 3 Then dictCCWhy(parts(1)) = UrlDecode(parts(3))

        ElseIf kind = "PH" Then
            Dim tok As String
            tok = UrlDecode(parts(1))
            dictPH(tok) = UrlDecode(parts(2))
            If UBound(parts) >= 3 Then dictPHWhy(tok) = UrlDecode(parts(3))

        ElseIf kind = "US" Then
            dictUS(parts(1)) = UrlDecode(parts(2))
            If UBound(parts) >= 3 Then dictUSWhy(parts(1)) = UrlDecode(parts(3))

        ElseIf kind = "CB" Then
            dictCB(parts(1)) = UrlDecode(parts(2))
            If UBound(parts) >= 3 Then dictCBWhy(parts(1)) = UrlDecode(parts(3))
        End If

NextLine:
    Next i

    ' content controls
    Dim cc As ContentControl
    For Each cc In doc.ContentControls
        Dim idKey As String
        idKey = CStr(cc.ID)

        If dictCC.Exists(idKey) Then
            Dim startPos As Long
            startPos = cc.Range.Start

            Dim val As String
            val = CStr(dictCC(idKey))
            cc.Range.Text = val

            If addReasons And dictCCWhy.Exists(idKey) Then
                Dim rr As Range
                Set rr = doc.Range(startPos, startPos + Len(val))
                AddCommentSafe doc, rr, "Reason: " & LimitLen(CStr(dictCCWhy(idKey)), 220)
            End If
        End If
    Next cc

    ' placeholders
    Dim phTok As Variant
    For Each phTok In dictPH.Keys
        ReplaceAll doc, CStr(phTok), CStr(dictPH(phTok))

        If addReasons And dictPHWhy.Exists(CStr(phTok)) Then
            Dim rrPH As Range
            Set rrPH = FindFirst(doc.StoryRanges(wdMainTextStory), CStr(dictPH(phTok)))
            If Not rrPH Is Nothing Then
                AddCommentSafe doc, rrPH, "Reason: " & LimitLen(CStr(dictPHWhy(CStr(phTok))), 220)
            End If
        End If
    Next phTok

    ' underscores
    If dictUS.Count > 0 Then ReplaceUnderscoreRunsByOccurrence doc, dictUS, dictUSWhy, addReasons

    ' checkboxes
    If dictCB.Count > 0 Then ApplyCheckboxGroups doc, dictCB, dictCBWhy, addReasons
End Sub

Private Sub ReplaceUnderscoreRunsByOccurrence(ByVal doc As Document, ByVal dictUS As Object, ByVal dictUSWhy As Object, ByVal addReasons As Boolean)
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
                ApplyOneUnderscoreGroup doc, dictUS, dictUSWhy, addReasons, occ, groupStart, groupEnd
                groupStart = scan.Start
                groupEnd = scan.End
            End If
        End If

        scan.Collapse wdCollapseEnd
    Loop

    If inGroup Then
        occ = occ + 1
        ApplyOneUnderscoreGroup doc, dictUS, dictUSWhy, addReasons, occ, groupStart, groupEnd
    End If
End Sub

Private Sub ApplyOneUnderscoreGroup(ByVal doc As Document, ByVal dictUS As Object, ByVal dictUSWhy As Object, ByVal addReasons As Boolean, ByVal occ As Long, ByVal groupStart As Long, ByVal groupEnd As Long)
    Dim key As String
    key = CStr(occ)
    If Not dictUS.Exists(key) Then Exit Sub

    Dim grp As Range
    Set grp = doc.Range(Start:=groupStart, End:=groupEnd)

    Dim startPos As Long
    startPos = grp.Start

    Dim val As String
    val = CStr(dictUS(key))
    grp.Text = val

    If addReasons And dictUSWhy.Exists(key) Then
        Dim rr As Range
        Set rr = doc.Range(startPos, startPos + Len(val))
        AddCommentSafe doc, rr, "Reason: " & LimitLen(CStr(dictUSWhy(key)), 220)
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
    Dim key As String
    key = CStr(groupOcc)
    If Not dictCB.Exists(key) Then Exit Sub

    Dim indicesCsv As String
    indicesCsv = CStr(dictCB(key))

    Dim idxParts() As String
    idxParts = Split(indicesCsv, ",")

    Dim firstMarkedStart As Long
    firstMarkedStart = -1

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
        Dim rr As Range
        Set rr = doc.Range(firstMarkedStart, firstMarkedStart + 3)
        AddCommentSafe doc, rr, "Reason: " & LimitLen(CStr(dictCBWhy(key)), 220)
    End If
End Sub

' ==========================================================
' AUDIT RESPONSE PARSER
' ==========================================================
Private Function ParseAuditResponse(ByVal resp As String, ByRef summary As String) As Boolean
    Dim line As String
    line = Trim$(NormalizeNewlines(resp))
    summary = ""

    Dim parts() As String
    parts = Split(line, "|")
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

    summary = UrlDecode(parts(2))
    ParseAuditResponse = (parts(1) = "OK")
End Function

' ==========================================================
' FIX TEXT BY INSTRUCTION COMMENTS (tracked)
' Comment text must start with FIX: / INSTR: / EDIT:
' Uses Len(instr) > 0 to avoid quote-related paste issues
' ==========================================================
Public Sub FixTextByInstructionComments_Tracked(ByVal doc As Document)
    Dim prevTrack As Boolean
    prevTrack = doc.TrackRevisions
    doc.TrackRevisions = True

    On Error GoTo CleanFail

    Dim i As Long
    Dim instr As String

    For i = 1 To doc.Comments.Count
        instr = NormalizeInstructionPrefix(doc.Comments(i).Range.Text)

        If Len(instr) > 0 Then
            Dim target As Range
            Set target = doc.Comments(i).Scope.Duplicate

            Dim selectedText As String
            selectedText = LimitLen(target.Text, 5000)

            Dim payload As String
            payload = BuildRevisePayload(doc.Name, instr, selectedText)

            Dim resp As String
            resp = HttpPostJson(REVISE_URL, payload)

            Dim revised As String
            Dim reason As String

            If ParseReviseResponse(resp, revised, reason) Then
                Dim startPos As Long
                startPos = target.Start

                target.Text = revised

                Dim applied As Range
                Set applied = doc.Range(startPos, startPos + Len(revised))

                AddCommentSafe doc, applied, "Applied fix. " & LimitLen(reason, 220)
            End If
        End If
    Next i

CleanExit:
    doc.TrackRevisions = prevTrack
    Exit Sub

CleanFail:
    doc.TrackRevisions = prevTrack
    MsgBox "Fix-by-comments failed: " & Err.Description, vbExclamation, "Form filler"
End Sub

Private Function NormalizeInstructionPrefix(ByVal commentText As String) As String
    Dim t As String
    t = Trim$(commentText)

    If Len(t) >= 4 And UCase$(Left$(t, 4)) = "FIX:" Then
        NormalizeInstructionPrefix = Trim$(Mid$(t, 5))
        Exit Function
    End If
    If Len(t) >= 6 And UCase$(Left$(t, 6)) = "INSTR:" Then
        NormalizeInstructionPrefix = Trim$(Mid$(t, 7))
        Exit Function
    End If
    If Len(t) >= 5 And UCase$(Left$(t, 5)) = "EDIT:" Then
        NormalizeInstructionPrefix = Trim$(Mid$(t, 6))
        Exit Function
    End If

    NormalizeInstructionPrefix = vbNullString
End Function

Private Function BuildRevisePayload(ByVal docName As String, ByVal instruction As String, ByVal selectedText As String) As String
    BuildRevisePayload = "{" & _
        """doc_name"":""" & JsonEscapeStrict(docName) & """," & _
        """instruction"":""" & JsonEscapeStrict(instruction) & """," & _
        """selected_text"":""" & JsonEscapeStrict(selectedText) & """" & _
    "}"
End Function

Private Function ParseReviseResponse(ByVal resp As String, ByRef revised As String, ByRef reason As String) As Boolean
    Dim line As String
    line = Trim$(NormalizeNewlines(resp))
    If Len(line) = 0 Then
        ParseReviseResponse = False
        Exit Function
    End If

    Dim parts() As String
    parts = Split(line, "|")
    If UBound(parts) < 2 Then
        ParseReviseResponse = False
        Exit Function
    End If
    If parts(0) <> "REV" Then
        ParseReviseResponse = False
        Exit Function
    End If

    revised = UrlDecode(parts(1))
    reason = UrlDecode(parts(2))
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
    snapshot = LimitLen(doc.Content.Text, 50000)

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
' HTTP
' ==========================================================
Private Function HttpPostJson(ByVal url As String, ByVal body As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

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
    doc.Comments.Add Range:=r, Text:=text
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
    Dim startPos As Long, endPos As Long
    startPos = r.Start - charsEachSide
    If startPos < 0 Then startPos = 0

    endPos = r.End + charsEachSide
    If endPos > doc.Content.End Then endPos = doc.Content.End

    Dim ctx As Range
    Set ctx = doc.Range(Start:=startPos, End:=endPos)
    GetContextAroundRange = ctx.Text
End Function

Private Function LimitLen(ByVal s As String, ByVal maxLen As Long) As String
    If Len(s) > maxLen Then LimitLen = Left$(s, maxLen) Else LimitLen = s
End Function

Private Function NormalizeNewlines(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    NormalizeNewlines = s
End Function

Private Function UrlDecode(ByVal s As String) As String
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
    UrlDecode = out
End Function

Private Function JsonEscapeStrict(ByVal s As String) As String
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
    JsonEscapeStrict = out
End Function
