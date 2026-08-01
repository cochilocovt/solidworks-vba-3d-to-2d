Option Explicit

' R23 Phase 9. Content-envelope-aware final layout.
'
' SAFETY BOUNDARY. Exactly one procedure changes a drawing:
' ApplyPlacementPlan, which refuses without an explicit allowMutation
' argument. R23_ProbeContentEnvelope never passes it and contains no
' IView.Position assignment and no EditRebuild3 call.
'
' R23-900. An envelope is everything that travels with a view: model
' outline, display-dimension primitives and text, note extents, leader
' points, callouts, section segments, arrow geometry and J-label positions
' with their text heights. Layout that reasons about the model outline alone
' moves a view into a place its annotations do not fit - which is exactly
' how the old J-J label ended up in the zone region.
'
' R23-902. The fixed P-0251 upward bias is replaced, not adjusted.
' Module9_LayoutEngine.bas lines 442-446 pin the source row to the top
' boundary - "Bias the P-0251 source row upward" - and a row pinned to a
' boundary has nowhere to put the annotations that hang above it. Placement
' here is derived from the envelopes' own sizes and the usable rectangle.
'
' FRAMES. Every coordinate in a CContentEnvelope is PAGE frame. Three of the
' sources are not:
'
'   IView.GetOutline          page frame, documented
'   IAnnotation.GetPosition   sheet frame in drawings, documented
'   INote.GetExtent           sheet space, documented
'   IView.GetSectionLineInfo2 VIEW-SKETCH frame, proved by Phase 0
'                             (payloadSegmentFrame=ViewSketchProved)
'   IDisplayData points       frame not stated by the Remarks
'
' Section geometry is converted once through ViewSketchToPage, whose inverse
' is round-trip checked against Module17_SectionPath.PageToViewSketch before
' anything is contributed. Display-data points are tested against the view's
' own documented outline and reported as consistent or not, because a
' contract the Help does not state is not a contract this project asserts.

' R23-905. Minimum clearance between any two view envelopes, and between a
' view envelope and a protected region. Matches Module9_LayoutEngine's
' existing VIEW_GAP_M so the two engines do not disagree about what "clear"
' means while both exist.
Public Const VIEW_CLEARANCE_M As Double = 0.012

' R23-906. J-J arrows and labels need at least 2 mm from the content border
' and the part-identification band.
Public Const SECTION_CLEARANCE_M As Double = 0.002

' R23-904. One correction pass, then fail. A second pass that still needs
' correcting is a placement that does not converge, and iterating it hides
' that.
Public Const MAX_CORRECTION_PASSES As Long = 1

' A round trip through the forward and inverse view transform must return
' the original page point to within this. It is a numerical check on two
' inverse functions, not a geometric tolerance.
Public Const TRANSFORM_ROUNDTRIP_TOLERANCE_M As Double = 0.000000001

Public Const ENVELOPE_KIND_VIEW As String = "ViewEnvelope"
Public Const ENVELOPE_KIND_PROTECTED As String = "ProtectedRegion"

Private Const SOURCE_OUTLINE As String = "Outline"
Private Const SOURCE_DIMENSION As String = "Dimension"
Private Const SOURCE_TEXT As String = "Text"
Private Const SOURCE_NOTE As String = "Note"
Private Const SOURCE_LEADER As String = "Leader"
Private Const SOURCE_ARROW As String = "Arrow"
Private Const SOURCE_SECTION As String = "Section"

' MCP corpus value for swDrawingViewTypes_e.swDrawingSectionView.
Private Const VIEW_TYPE_SECTION As Long = 2

' R23-909 seal. Set when layout completes; compared afterwards.
Private mLayoutSealed As Boolean
Private mSealedMutationSequence As Long
Private mSealedLastMutation As String

Private mEmitDiagnostics As Boolean

Private Sub EmitInfo( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddInfo message
    If mEmitDiagnostics Then Debug.Print message
End Sub

Private Sub EmitWarning( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddWarning message
    If mEmitDiagnostics Then Debug.Print message
End Sub

Private Sub EmitFailure( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddFailure message
    If mEmitDiagnostics Then Debug.Print message
End Sub

Public Function EnvelopeToken(ByVal value As String) As String
    Dim result As String
    result = value
    result = Replace$(result, "|", "/")
    result = Replace$(result, "=", ":")
    result = Replace$(result, vbCr, " ")
    result = Replace$(result, vbLf, " ")
    If Len(Trim$(result)) = 0 Then result = "Empty"
    EnvelopeToken = result
End Function

Private Function AppendFailure( _
    ByVal existing As String, _
    ByVal reason As String) As String

    If Len(existing) = 0 Then
        AppendFailure = reason
    Else
        AppendFailure = existing & ";" & reason
    End If
End Function

Private Function FormatMetres(ByVal value As Double) As String
    FormatMetres = Format$(value, "0.000000")
End Function

Private Function SafeViewName( _
    ByRef swView As SldWorks.View) As String

    On Error Resume Next
    If swView Is Nothing Then
        SafeViewName = "Nothing"
        Exit Function
    End If

    SafeViewName = "Unnamed"
    SafeViewName = swView.GetName2
End Function

Public Function NewEnvelope( _
    ByVal kind As String, _
    ByVal name As String) As CContentEnvelope

    Dim result As CContentEnvelope
    Set result = New CContentEnvelope
    result.Kind = kind
    result.Name = name
    Set NewEnvelope = result
End Function

' Contributes one page point, rejecting anything off the sheet. A page-frame
' point belonging to a view on the sheet lies on the sheet, so a rejection
' is a frame smell rather than a rounding artefact - and it is counted
' instead of being silently dropped.
Private Sub ContributePoint( _
    ByRef envelope As CContentEnvelope, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal source As String, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double)

    If x < 0# Or y < 0# Or x > sheetWidth Or y > sheetHeight Then
        envelope.RejectedOffSheet = envelope.RejectedOffSheet + 1
        Exit Sub
    End If

    envelope.AddPoint x, y, source
End Sub

' IView.GetOutline returns [Xmin, Ymin, Xmax, Ymax] in metres on the drawing
' page, per its Remarks. This is the one source whose page frame the Help
' states outright.
Private Function AddOutline( _
    ByRef envelope As CContentEnvelope, _
    ByRef swView As SldWorks.View, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double) As Boolean

    On Error GoTo Failed

    Dim outline As Variant
    outline = swView.GetOutline

    If IsEmpty(outline) Or Not IsArray(outline) Then
        envelope.SourceFailures = AppendFailure( _
            envelope.SourceFailures, "NoOutline")
        Exit Function
    End If

    Dim baseIndex As Long
    baseIndex = LBound(outline)

    If (UBound(outline) - baseIndex + 1) < 4 Then
        envelope.SourceFailures = AppendFailure( _
            envelope.SourceFailures, "OutlineTooShort")
        Exit Function
    End If

    ContributePoint envelope, CDbl(outline(baseIndex)), _
        CDbl(outline(baseIndex + 1)), SOURCE_OUTLINE, _
        sheetWidth, sheetHeight
    ContributePoint envelope, CDbl(outline(baseIndex + 2)), _
        CDbl(outline(baseIndex + 3)), SOURCE_OUTLINE, _
        sheetWidth, sheetHeight

    AddOutline = True
    Exit Function

Failed:
    envelope.SourceFailures = AppendFailure( _
        envelope.SourceFailures, "OutlineError:" & CStr(Err.Number))
End Function

' IAnnotation.GetPosition is documented as sheet-relative in a drawing, and
' leader points are computed from the text and attachment points, so they
' share that frame.
Private Sub AddAnnotationGeometry( _
    ByRef envelope As CContentEnvelope, _
    ByRef annotation As SldWorks.Annotation, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double)

    On Error GoTo Failed

    If annotation Is Nothing Then Exit Sub

    Dim position As Variant
    position = annotation.GetPosition

    If IsArray(position) Then
        If (UBound(position) - LBound(position) + 1) >= 2 Then
            ContributePoint envelope, _
                CDbl(position(LBound(position))), _
                CDbl(position(LBound(position) + 1)), _
                SOURCE_DIMENSION, sheetWidth, sheetHeight
        End If
    End If

    Dim leaderCount As Long
    leaderCount = CLng(annotation.GetLeaderCount)

    Dim i As Long
    For i = 0 To leaderCount - 1
        Dim points As Variant
        points = annotation.GetLeaderPointsAtIndex(i)

        If Not IsArray(points) Then GoTo ContinueLeader

        Dim first As Long
        Dim total As Long
        first = LBound(points)
        total = UBound(points) - first + 1

        ' Consume whole XYZ triples from the array actually returned. The
        ' point count could be derived from GetLeaderStyle, but that value
        ' is OR-ed with attachment bitmask flags, and the corpus returns
        ' those flag members with mangled values - reading the array length
        ' needs no such reconstruction.
        Dim p As Long
        For p = 0 To (total \ 3) - 1
            ContributePoint envelope, _
                CDbl(points(first + p * 3)), _
                CDbl(points(first + p * 3 + 1)), _
                SOURCE_LEADER, sheetWidth, sheetHeight
        Next p

ContinueLeader:
    Next i

    Exit Sub

Failed:
    envelope.SourceFailures = AppendFailure( _
        envelope.SourceFailures, "AnnotationError:" & CStr(Err.Number))
End Sub

' Display-dimension primitives: the drawn lines and the text boxes.
'
' The Remarks do NOT state the frame of IDisplayData coordinates, so this
' contributes them and reports whether they agree with the view's own
' documented outline. GetTextPositionAtIndex is explicitly an OFFSET from
' the display data origin, not an absolute point, so it is added to the
' annotation position rather than used on its own - treating an offset as a
' coordinate would drag every envelope towards the sheet origin.
Private Sub AddDisplayDimensionGeometry( _
    ByRef envelope As CContentEnvelope, _
    ByRef displayDimension As SldWorks.DisplayDimension, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef consistentPoints As Long, _
    ByRef inconsistentPoints As Long, _
    ByVal outlineMinX As Double, _
    ByVal outlineMinY As Double, _
    ByVal outlineMaxX As Double, _
    ByVal outlineMaxY As Double)

    On Error GoTo Failed

    Dim annotation As SldWorks.Annotation
    Set annotation = displayDimension.GetAnnotation

    Dim originX As Double
    Dim originY As Double
    Dim originKnown As Boolean

    If Not annotation Is Nothing Then
        Dim position As Variant
        position = annotation.GetPosition
        If IsArray(position) Then
            If (UBound(position) - LBound(position) + 1) >= 2 Then
                originX = CDbl(position(LBound(position)))
                originY = CDbl(position(LBound(position) + 1))
                originKnown = True
            End If
        End If
    End If

    Dim displayData As SldWorks.DisplayData
    Set displayData = displayDimension.GetDisplayData

    If displayData Is Nothing Then Exit Sub

    Dim lineCount As Long
    lineCount = CLng(displayData.GetLineCount)

    Dim i As Long
    For i = 0 To lineCount - 1
        Dim lineInfo As Variant
        lineInfo = displayData.GetLineAtIndex3(i)

        If Not IsArray(lineInfo) Then GoTo ContinueLine

        Dim lineBase As Long
        lineBase = LBound(lineInfo)

        ' [color, lineType, lineStyle, lineWeight, startPt[3], endPt[3]]
        If (UBound(lineInfo) - lineBase + 1) < 10 Then GoTo ContinueLine

        RecordFrameAgreement CDbl(lineInfo(lineBase + 4)), _
            CDbl(lineInfo(lineBase + 5)), _
            outlineMinX, outlineMinY, outlineMaxX, outlineMaxY, _
            consistentPoints, inconsistentPoints

        ContributePoint envelope, CDbl(lineInfo(lineBase + 4)), _
            CDbl(lineInfo(lineBase + 5)), SOURCE_DIMENSION, _
            sheetWidth, sheetHeight
        ContributePoint envelope, CDbl(lineInfo(lineBase + 7)), _
            CDbl(lineInfo(lineBase + 8)), SOURCE_DIMENSION, _
            sheetWidth, sheetHeight

ContinueLine:
    Next i

    If Not originKnown Then Exit Sub

    Dim textCount As Long
    textCount = CLng(displayData.GetTextCount)

    For i = 0 To textCount - 1
        Dim offset As Variant
        offset = displayData.GetTextPositionAtIndex(i)

        If Not IsArray(offset) Then GoTo ContinueText
        If (UBound(offset) - LBound(offset) + 1) < 2 Then GoTo ContinueText

        Dim textX As Double
        Dim textY As Double
        textX = originX + CDbl(offset(LBound(offset)))
        textY = originY + CDbl(offset(LBound(offset) + 1))

        Dim boxWidth As Double
        Dim boxHeight As Double
        boxWidth = 0#
        boxHeight = 0#

        On Error Resume Next
        boxWidth = CDbl(displayData.GetTextInBoxWidthAtIndex(i))
        boxHeight = CDbl(displayData.GetTextInBoxHeightAtIndex(i))
        If boxHeight <= 0# Then
            boxHeight = CDbl(displayData.GetTextHeightAtIndex(i))
        End If
        On Error GoTo Failed

        ContributePoint envelope, textX, textY, SOURCE_TEXT, _
            sheetWidth, sheetHeight
        ContributePoint envelope, textX + boxWidth, textY + boxHeight, _
            SOURCE_TEXT, sheetWidth, sheetHeight

ContinueText:
    Next i

    Exit Sub

Failed:
    envelope.SourceFailures = AppendFailure( _
        envelope.SourceFailures, "DisplayDataError:" & CStr(Err.Number))
End Sub

' A page-frame point belonging to this view lies at or near the view's own
' documented outline. A view-frame point would sit near the sheet origin
' instead. This counts the agreement rather than asserting the frame.
Private Sub RecordFrameAgreement( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal outlineMinX As Double, _
    ByVal outlineMinY As Double, _
    ByVal outlineMaxX As Double, _
    ByVal outlineMaxY As Double, _
    ByRef consistentPoints As Long, _
    ByRef inconsistentPoints As Long)

    Dim slack As Double
    slack = VIEW_CLEARANCE_M * 10#

    If x >= (outlineMinX - slack) And x <= (outlineMaxX + slack) And _
       y >= (outlineMinY - slack) And y <= (outlineMaxY + slack) Then

        consistentPoints = consistentPoints + 1
    Else
        inconsistentPoints = inconsistentPoints + 1
    End If
End Sub

' INote.GetExtent returns six doubles - lower-left and upper-right - in
' sheet space, per its Remarks.
Private Sub AddNoteExtents( _
    ByRef envelope As CContentEnvelope, _
    ByRef swView As SldWorks.View, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double)

    On Error GoTo Failed

    Dim notes As Variant
    notes = swView.GetNotes

    If IsEmpty(notes) Or Not IsArray(notes) Then Exit Sub

    Dim i As Long
    For i = LBound(notes) To UBound(notes)
        Dim note As SldWorks.Note
        Set note = notes(i)
        If note Is Nothing Then GoTo ContinueNote

        Dim extent As Variant
        extent = note.GetExtent

        If IsArray(extent) Then
            Dim baseIndex As Long
            baseIndex = LBound(extent)

            If (UBound(extent) - baseIndex + 1) >= 5 Then
                ContributePoint envelope, CDbl(extent(baseIndex)), _
                    CDbl(extent(baseIndex + 1)), SOURCE_NOTE, _
                    sheetWidth, sheetHeight
                ContributePoint envelope, CDbl(extent(baseIndex + 3)), _
                    CDbl(extent(baseIndex + 4)), SOURCE_NOTE, _
                    sheetWidth, sheetHeight
            End If
        End If

        AddAnnotationGeometry envelope, note.GetAnnotation, _
            sheetWidth, sheetHeight

ContinueNote:
    Next i

    Exit Sub

Failed:
    envelope.SourceFailures = AppendFailure( _
        envelope.SourceFailures, "NoteError:" & CStr(Err.Number))
End Sub

' R23-705's transform, inverted. Page = origin + scale * R(angle) * view.
'
' Phase 0 proved the section-line payload is in the source view's sketch
' frame (payloadSegmentFrame=ViewSketchProved), so section geometry has to
' come back through here before it can join a page-frame envelope.
Public Function ViewSketchToPage( _
    ByRef swView As SldWorks.View, _
    ByVal viewX As Double, _
    ByVal viewY As Double, _
    ByRef pageX As Double, _
    ByRef pageY As Double, _
    ByRef frameProof As String) As Boolean

    On Error GoTo Failed

    frameProof = "frame=Unproven"

    If swView Is Nothing Then Exit Function

    Dim xform As Variant
    xform = swView.GetXform

    If IsEmpty(xform) Or Not IsArray(xform) Then
        frameProof = "frame=Reject|reason=NoXform"
        Exit Function
    End If

    If (UBound(xform) - LBound(xform) + 1) < 3 Then
        frameProof = "frame=Reject|reason=XformTooShort"
        Exit Function
    End If

    Dim baseIndex As Long
    baseIndex = LBound(xform)

    Dim originX As Double
    Dim originY As Double
    Dim viewScale As Double
    originX = CDbl(xform(baseIndex))
    originY = CDbl(xform(baseIndex + 1))
    viewScale = CDbl(xform(baseIndex + 2))

    If viewScale <= 0# Then
        frameProof = "frame=Reject|reason=NonPositiveScale"
        Exit Function
    End If

    Dim viewAngle As Double
    viewAngle = swView.Angle

    Dim deltaX As Double
    Dim deltaY As Double
    deltaX = viewX * Cos(viewAngle) - viewY * Sin(viewAngle)
    deltaY = viewX * Sin(viewAngle) + viewY * Cos(viewAngle)

    pageX = originX + deltaX * viewScale
    pageY = originY + deltaY * viewScale

    frameProof = "frame=Page" & _
        "|source=Inverse(IView.GetXform+IView.Angle)" & _
        "|conversions=1"

    ViewSketchToPage = True
    Exit Function

Failed:
    frameProof = "frame=Reject|reason=Error:" & CStr(Err.Number)
    ViewSketchToPage = False
End Function

' Round-trips one page point through Module17's forward transform and this
' module's inverse. Two functions that claim to be inverses either agree to
' floating-point noise or one of them is wrong, and asserting the pairing is
' cheaper than discovering it after a view has been moved.
Public Function ProveInverseTransform( _
    ByRef swView As SldWorks.View) As String

    On Error GoTo Failed

    Dim outline As Variant
    outline = swView.GetOutline

    If Not IsArray(outline) Then
        ProveInverseTransform = "inverse=NoOutline"
        Exit Function
    End If

    Dim baseIndex As Long
    baseIndex = LBound(outline)

    Dim samplePageX As Double
    Dim samplePageY As Double
    samplePageX = CDbl(outline(baseIndex))
    samplePageY = CDbl(outline(baseIndex + 1))

    Dim viewX As Double
    Dim viewY As Double
    Dim forwardProof As String

    If Not Module17_SectionPath.PageToViewSketch( _
        swView, samplePageX, samplePageY, viewX, viewY, forwardProof) Then

        ProveInverseTransform = "inverse=ForwardFailed"
        Exit Function
    End If

    Dim backX As Double
    Dim backY As Double
    Dim inverseProof As String

    If Not ViewSketchToPage(swView, viewX, viewY, backX, backY, _
        inverseProof) Then

        ProveInverseTransform = "inverse=InverseFailed"
        Exit Function
    End If

    Dim deltaM As Double
    deltaM = Sqr((backX - samplePageX) * (backX - samplePageX) + _
        (backY - samplePageY) * (backY - samplePageY))

    ProveInverseTransform = "inverse=" & _
        CStr(deltaM <= TRANSFORM_ROUNDTRIP_TOLERANCE_M) & _
        "|roundTripDeltaM=" & Format$(deltaM, "0.000000000000")
    Exit Function

Failed:
    ProveInverseTransform = "inverse=Error:" & CStr(Err.Number)
End Function

' R23-900, section half. Parses IView.GetSectionLineInfo2 and contributes
' every segment endpoint, both arrow heads and both label points with the
' label text height.
'
' The documented grammar is
' [numSectionLines, layer, {numSegments, {lineType, startPt[3], endPt[3]},
'  arrow1[11], arrow2[11], textPt1[3], textPt2[3], textHeight}].
' GetSectionLineCount2's own Remarks say its Size includes a layer double
' for EACH section line, which the grammar above does not show. Rather than
' pick one reading, both are tried and the one whose consumption matches the
' array length exactly is reported. A parse that consumes the wrong number
' of doubles produces plausible coordinates, which is the worst kind of
' wrong.
Private Function AddSectionGeometry( _
    ByRef envelope As CContentEnvelope, _
    ByRef swView As SldWorks.View, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef grammarProof As String) As Boolean

    On Error GoTo Failed

    grammarProof = "sectionGrammar=NotAttempted"

    Dim info As Variant
    info = swView.GetSectionLineInfo2

    If IsEmpty(info) Or Not IsArray(info) Then
        grammarProof = "sectionGrammar=NoInfo"
        Exit Function
    End If

    Dim total As Long
    total = UBound(info) - LBound(info) + 1

    If ConsumeSectionInfo(envelope, swView, info, False, _
        sheetWidth, sheetHeight, True) = total Then

        grammarProof = "sectionGrammar=LayerOnce|items=" & CStr(total)
        AddSectionGeometry = _
            (ConsumeSectionInfo(envelope, swView, info, False, _
                sheetWidth, sheetHeight, False) = total)
        Exit Function
    End If

    If ConsumeSectionInfo(envelope, swView, info, True, _
        sheetWidth, sheetHeight, True) = total Then

        grammarProof = "sectionGrammar=LayerPerLine|items=" & CStr(total)
        AddSectionGeometry = _
            (ConsumeSectionInfo(envelope, swView, info, True, _
                sheetWidth, sheetHeight, False) = total)
        Exit Function
    End If

    grammarProof = "sectionGrammar=Unmatched|items=" & CStr(total)
    envelope.SourceFailures = AppendFailure( _
        envelope.SourceFailures, "SectionGrammarUnmatched")
    Exit Function

Failed:
    grammarProof = "sectionGrammar=Error:" & CStr(Err.Number)
End Function

' Walks the section-line array under one grammar. When dryRun is True it
' contributes nothing and only reports how many doubles the grammar
' consumed, so the grammar is chosen before any point is added.
Private Function ConsumeSectionInfo( _
    ByRef envelope As CContentEnvelope, _
    ByRef swView As SldWorks.View, _
    ByRef info As Variant, _
    ByVal layerPerLine As Boolean, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByVal dryRun As Boolean) As Long

    On Error GoTo Failed

    Dim baseIndex As Long
    baseIndex = LBound(info)

    Dim cursor As Long
    cursor = 0

    Dim lineCount As Long
    lineCount = CLng(info(baseIndex + cursor))
    cursor = cursor + 1

    If Not layerPerLine Then cursor = cursor + 1

    If lineCount < 1 Then
        ConsumeSectionInfo = -1
        Exit Function
    End If

    Dim lineIndex As Long
    For lineIndex = 1 To lineCount
        If layerPerLine Then cursor = cursor + 1

        Dim segmentCount As Long
        segmentCount = CLng(info(baseIndex + cursor))
        cursor = cursor + 1

        If segmentCount < 1 Then
            ConsumeSectionInfo = -1
            Exit Function
        End If

        Dim segment As Long
        For segment = 1 To segmentCount
            ' lineType, startPt[3], endPt[3]
            ContributeSectionPoint envelope, swView, info, _
                baseIndex + cursor + 1, SOURCE_SECTION, _
                sheetWidth, sheetHeight, dryRun
            ContributeSectionPoint envelope, swView, info, _
                baseIndex + cursor + 4, SOURCE_SECTION, _
                sheetWidth, sheetHeight, dryRun
            cursor = cursor + 7
        Next segment

        ' arrow1 and arrow2: start[3], end[3], width, height, style
        Dim arrow As Long
        For arrow = 1 To 2
            ContributeSectionPoint envelope, swView, info, _
                baseIndex + cursor, SOURCE_ARROW, _
                sheetWidth, sheetHeight, dryRun
            ContributeSectionPoint envelope, swView, info, _
                baseIndex + cursor + 3, SOURCE_ARROW, _
                sheetWidth, sheetHeight, dryRun
            cursor = cursor + 11
        Next arrow

        ' textPt1[3], textPt2[3], textHeight. The label occupies the text
        ' height above its point, so the height is added as a second point
        ' rather than recorded and ignored.
        Dim textHeight As Double
        textHeight = CDbl(info(baseIndex + cursor + 6))

        ContributeSectionLabel envelope, swView, info, _
            baseIndex + cursor, textHeight, sheetWidth, sheetHeight, dryRun
        ContributeSectionLabel envelope, swView, info, _
            baseIndex + cursor + 3, textHeight, sheetWidth, sheetHeight, _
            dryRun

        cursor = cursor + 7
    Next lineIndex

    ConsumeSectionInfo = cursor
    Exit Function

Failed:
    ConsumeSectionInfo = -1
End Function

Private Sub ContributeSectionPoint( _
    ByRef envelope As CContentEnvelope, _
    ByRef swView As SldWorks.View, _
    ByRef info As Variant, _
    ByVal index As Long, _
    ByVal source As String, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByVal dryRun As Boolean)

    If dryRun Then Exit Sub
    On Error GoTo Failed

    Dim pageX As Double
    Dim pageY As Double
    Dim frameProof As String

    If Not ViewSketchToPage(swView, CDbl(info(index)), _
        CDbl(info(index + 1)), pageX, pageY, frameProof) Then

        envelope.SourceFailures = AppendFailure( _
            envelope.SourceFailures, "SectionFrameConversionFailed")
        Exit Sub
    End If

    ContributePoint envelope, pageX, pageY, source, _
        sheetWidth, sheetHeight
    Exit Sub

Failed:
    envelope.SourceFailures = AppendFailure( _
        envelope.SourceFailures, "SectionPointError:" & CStr(Err.Number))
End Sub

Private Sub ContributeSectionLabel( _
    ByRef envelope As CContentEnvelope, _
    ByRef swView As SldWorks.View, _
    ByRef info As Variant, _
    ByVal index As Long, _
    ByVal textHeight As Double, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByVal dryRun As Boolean)

    If dryRun Then Exit Sub
    On Error GoTo Failed

    Dim pageX As Double
    Dim pageY As Double
    Dim frameProof As String

    If Not ViewSketchToPage(swView, CDbl(info(index)), _
        CDbl(info(index + 1)), pageX, pageY, frameProof) Then

        Exit Sub
    End If

    ContributePoint envelope, pageX, pageY, SOURCE_TEXT, _
        sheetWidth, sheetHeight

    If textHeight > 0# Then
        ContributePoint envelope, pageX, pageY + textHeight, _
            SOURCE_TEXT, sheetWidth, sheetHeight
    End If

    Exit Sub

Failed:
    envelope.SourceFailures = AppendFailure( _
        envelope.SourceFailures, "SectionLabelError:" & CStr(Err.Number))
End Sub

' R23-900. One complete content envelope for one view.
Public Function BuildViewEnvelope( _
    ByRef swView As SldWorks.View, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByRef evidence As CRunEvidence) As CContentEnvelope

    Dim envelope As CContentEnvelope
    Set envelope = NewEnvelope(ENVELOPE_KIND_VIEW, SafeViewName(swView))
    Set BuildViewEnvelope = envelope

    On Error GoTo Failed

    If swView Is Nothing Then
        envelope.SourceFailures = "NoView"
        Exit Function
    End If

    If Not AddOutline(envelope, swView, sheetWidth, sheetHeight) Then
        Exit Function
    End If

    Dim outlineMinX As Double
    Dim outlineMinY As Double
    Dim outlineMaxX As Double
    Dim outlineMaxY As Double
    outlineMinX = envelope.MinX
    outlineMinY = envelope.MinY
    outlineMaxX = envelope.MaxX
    outlineMaxY = envelope.MaxY

    Dim consistentPoints As Long
    Dim inconsistentPoints As Long

    Dim dimensions As Variant
    dimensions = swView.GetDisplayDimensions

    If IsArray(dimensions) Then
        Dim i As Long
        For i = LBound(dimensions) To UBound(dimensions)
            Dim displayDimension As SldWorks.DisplayDimension
            Set displayDimension = dimensions(i)
            If displayDimension Is Nothing Then GoTo ContinueDimension

            AddAnnotationGeometry envelope, _
                displayDimension.GetAnnotation, sheetWidth, sheetHeight

            AddDisplayDimensionGeometry envelope, displayDimension, _
                sheetWidth, sheetHeight, _
                consistentPoints, inconsistentPoints, _
                outlineMinX, outlineMinY, outlineMaxX, outlineMaxY

ContinueDimension:
        Next i
    End If

    AddNoteExtents envelope, swView, sheetWidth, sheetHeight

    Dim grammarProof As String
    AddSectionGeometry envelope, swView, sheetWidth, sheetHeight, _
        grammarProof

    EmitInfo evidence, "ENVELOPE|" & envelope.Summary() & _
        "|" & grammarProof & _
        "|displayDataFramePageConsistent=" & CStr(consistentPoints) & _
        "|displayDataFrameInconsistent=" & CStr(inconsistentPoints) & _
        "|" & ProveInverseTransform(swView)
    Exit Function

Failed:
    envelope.SourceFailures = AppendFailure( _
        envelope.SourceFailures, "EnvelopeError:" & CStr(Err.Number))
End Function

Public Function RectangleEnvelope( _
    ByVal name As String, _
    ByVal leftX As Double, _
    ByVal bottomY As Double, _
    ByVal rightX As Double, _
    ByVal topY As Double) As CContentEnvelope

    Dim result As CContentEnvelope
    Set result = NewEnvelope(ENVELOPE_KIND_PROTECTED, name)
    result.AddRectangle leftX, bottomY, rightX, topY, "Protected"
    Set RectangleEnvelope = result
End Function

' R23-901. The protected sheet rectangles, taken from the measurements
' Module8_RuntimeSupport.MeasureControlledSheetRegions already recorded on
' the evidence ledger. Re-deriving them here would create a second set of
' numbers that can drift from the first.
Public Function BuildProtectedRegions( _
    ByRef evidence As CRunEvidence) As Collection

    Dim result As Collection
    Set result = New Collection
    Set BuildProtectedRegions = result

    On Error GoTo Failed

    If evidence Is Nothing Then Exit Function

    ' The content border is protected as four strips rather than one
    ' rectangle: the drawable area is INSIDE it, so a single rectangle would
    ' declare every view a violation.
    result.Add RectangleEnvelope("ContentBorderLeftStrip", _
        0#, 0#, evidence.ContentBorderLeft, evidence.SheetHeight)
    result.Add RectangleEnvelope("ContentBorderRightStrip", _
        evidence.ContentBorderRight, 0#, _
        evidence.SheetWidth, evidence.SheetHeight)
    result.Add RectangleEnvelope("ContentBorderBottomStrip", _
        0#, 0#, evidence.SheetWidth, evidence.ContentBorderBottom)
    result.Add RectangleEnvelope("ContentBorderTopStrip", _
        0#, evidence.ContentBorderTop, _
        evidence.SheetWidth, evidence.SheetHeight)

    result.Add RectangleEnvelope("TitleBlock", _
        evidence.TitleBlockLeft, evidence.TitleBlockBottom, _
        evidence.TitleBlockRight, evidence.TitleBlockTop)

    If evidence.PartIdentificationBoundsProven Then
        result.Add RectangleEnvelope("PartIdentificationBand", _
            evidence.PartIdentificationLeft, _
            evidence.PartIdentificationBottom, _
            evidence.PartIdentificationRight, _
            evidence.PartIdentificationTop)
    End If

    Exit Function

Failed:
    Set BuildProtectedRegions = result
End Function

' Separating-axis clearance. Positive is a gap; zero or negative means the
' rectangles touch or overlap.
Public Function ClearanceBetween( _
    ByRef first As CContentEnvelope, _
    ByRef second As CContentEnvelope) As Double

    Dim gapX As Double
    Dim gapY As Double

    gapX = first.MinX - second.MaxX
    If (second.MinX - first.MaxX) > gapX Then
        gapX = second.MinX - first.MaxX
    End If

    gapY = first.MinY - second.MaxY
    If (second.MinY - first.MaxY) > gapY Then
        gapY = second.MinY - first.MaxY
    End If

    If gapX > gapY Then
        ClearanceBetween = gapX
    Else
        ClearanceBetween = gapY
    End If
End Function

' R23-906. A section view's arrows and label need 2 mm from the protected
' regions; everything else needs the ordinary view clearance.
Private Function RequiredClearance( _
    ByRef swView As SldWorks.View, _
    ByRef other As CContentEnvelope) As Double

    RequiredClearance = VIEW_CLEARANCE_M

    On Error Resume Next
    If swView Is Nothing Then Exit Function
    If other Is Nothing Then Exit Function

    If StrComp(other.Kind, ENVELOPE_KIND_PROTECTED, _
        vbBinaryCompare) <> 0 Then Exit Function

    If swView.Type = VIEW_TYPE_SECTION Then
        RequiredClearance = SECTION_CLEARANCE_M
    End If
End Function

' R23-905 and R23-906. Every view envelope against every other view
' envelope and every protected rectangle, with the required clearance
' stated for each pair rather than assumed.
Public Function VerifyClearances( _
    ByRef viewEnvelopes As Collection, _
    ByRef views As Collection, _
    ByRef protectedRegions As Collection) As String

    On Error GoTo Failed

    Dim failures As String
    Dim checks As Long

    Dim i As Long
    For i = 1 To viewEnvelopes.Count
        Dim envelope As CContentEnvelope
        Set envelope = viewEnvelopes(i)

        Dim swView As SldWorks.View
        Set swView = views(i)

        Dim j As Long
        For j = i + 1 To viewEnvelopes.Count
            Dim other As CContentEnvelope
            Set other = viewEnvelopes(j)

            checks = checks + 1
            If ClearanceBetween(envelope, other) < VIEW_CLEARANCE_M Then
                failures = AppendFailure(failures, _
                    "ViewOverlap:" & envelope.Name & "/" & other.Name)
            End If
        Next j

        Dim p As Long
        For p = 1 To protectedRegions.Count
            Dim region As CContentEnvelope
            Set region = protectedRegions(p)

            Dim required As Double
            required = RequiredClearance(swView, region)

            checks = checks + 1
            If ClearanceBetween(envelope, region) < required Then
                failures = AppendFailure(failures, _
                    "ProtectedIntrusion:" & envelope.Name & "/" & _
                    region.Name)
            End If
        Next p
    Next i

    If Len(failures) = 0 Then failures = "None"

    VerifyClearances = "clearanceChecks=" & CStr(checks) & _
        "|viewClearanceM=" & FormatMetres(VIEW_CLEARANCE_M) & _
        "|sectionClearanceM=" & FormatMetres(SECTION_CLEARANCE_M) & _
        "|clearanceFailures=" & failures
    Exit Function

Failed:
    VerifyClearances = "clearanceFailures=Error:" & CStr(Err.Number)
End Function

' R23-902 and R23-908. Constraint-based placement from the envelopes' own
' sizes and the usable rectangle. Read-only: it returns target centres and
' moves nothing.
'
' Rows are packed by envelope width and the whole block is centred in the
' usable band. Nothing is pinned to a boundary, because a row pinned to the
' top boundary has nowhere to put the annotations that hang above it - which
' is what the fixed P-0251 upward bias did.
Public Function PlanPlacement( _
    ByRef viewEnvelopes As Collection, _
    ByRef evidence As CRunEvidence, _
    ByRef planProof As String) As Collection

    Dim result As Collection
    Set result = New Collection
    Set PlanPlacement = result

    On Error GoTo Failed

    planProof = "plan=NotAttempted"

    Dim usableWidth As Double
    Dim usableHeight As Double
    usableWidth = evidence.UsableRight - evidence.UsableLeft
    usableHeight = evidence.UsableTop - evidence.UsableBottom

    If usableWidth <= 0# Or usableHeight <= 0# Then
        planProof = "plan=Reject|reason=UsableRegionEmpty"
        Exit Function
    End If

    ' Greedy row packing in the given order. Order is the caller's, so the
    ' plan is deterministic for a given sheet.
    Dim rowWidths As Collection
    Dim rowHeights As Collection
    Dim rowFirst As Collection
    Dim rowLast As Collection
    Set rowWidths = New Collection
    Set rowHeights = New Collection
    Set rowFirst = New Collection
    Set rowLast = New Collection

    Dim currentWidth As Double
    Dim currentHeight As Double
    Dim currentFirst As Long
    currentFirst = 1

    Dim i As Long
    For i = 1 To viewEnvelopes.Count
        Dim envelope As CContentEnvelope
        Set envelope = viewEnvelopes(i)

        Dim addedWidth As Double
        addedWidth = envelope.Width()
        If i > currentFirst Then
            addedWidth = addedWidth + VIEW_CLEARANCE_M
        End If

        If (currentWidth + addedWidth) > usableWidth And _
            i > currentFirst Then

            rowWidths.Add currentWidth
            rowHeights.Add currentHeight
            rowFirst.Add currentFirst
            rowLast.Add i - 1

            currentFirst = i
            currentWidth = envelope.Width()
            currentHeight = envelope.Height()
        Else
            currentWidth = currentWidth + addedWidth
            If envelope.Height() > currentHeight Then
                currentHeight = envelope.Height()
            End If
        End If
    Next i

    If viewEnvelopes.Count > 0 Then
        rowWidths.Add currentWidth
        rowHeights.Add currentHeight
        rowFirst.Add currentFirst
        rowLast.Add viewEnvelopes.Count
    End If

    Dim totalHeight As Double
    Dim r As Long
    For r = 1 To rowHeights.Count
        totalHeight = totalHeight + CDbl(rowHeights(r))
        If r > 1 Then totalHeight = totalHeight + VIEW_CLEARANCE_M
    Next r

    Dim widestRow As Double
    For r = 1 To rowWidths.Count
        If CDbl(rowWidths(r)) > widestRow Then
            widestRow = CDbl(rowWidths(r))
        End If
    Next r

    ' R23-908. The complete content either fits or the sheet is too small.
    ' R23-907 forbids shrinking a view to make it fit, so this is a failure
    ' and a request, not a fallback.
    If widestRow > usableWidth Or totalHeight > usableHeight Then
        planProof = "plan=Reject|reason=LargerSheetRequired" & _
            "|requiredWidthM=" & FormatMetres(widestRow) & _
            "|requiredHeightM=" & FormatMetres(totalHeight) & _
            "|usableWidthM=" & FormatMetres(usableWidth) & _
            "|usableHeightM=" & FormatMetres(usableHeight)
        Exit Function
    End If

    Dim blockTop As Double
    blockTop = evidence.UsableBottom + _
        (usableHeight + totalHeight) / 2#

    Dim cursorY As Double
    cursorY = blockTop

    For r = 1 To rowHeights.Count
        Dim rowHeight As Double
        rowHeight = CDbl(rowHeights(r))

        Dim rowCentreY As Double
        rowCentreY = cursorY - rowHeight / 2#

        Dim cursorX As Double
        cursorX = evidence.UsableLeft + _
            (usableWidth - CDbl(rowWidths(r))) / 2#

        Dim first As Long
        Dim last As Long
        first = CLng(rowFirst(r))
        last = CLng(rowLast(r))

        Dim c As Long
        For c = first To last
            Dim cellEnvelope As CContentEnvelope
            Set cellEnvelope = viewEnvelopes(c)

            If c > first Then cursorX = cursorX + VIEW_CLEARANCE_M

            Dim centre(0 To 1) As Double
            centre(0) = cursorX + cellEnvelope.Width() / 2#
            centre(1) = rowCentreY
            result.Add centre

            cursorX = cursorX + cellEnvelope.Width()
        Next c

        cursorY = cursorY - rowHeight - VIEW_CLEARANCE_M
    Next r

    planProof = "plan=Accept|rows=" & CStr(rowHeights.Count) & _
        "|cells=" & CStr(result.Count) & _
        "|blockWidthM=" & FormatMetres(widestRow) & _
        "|blockHeightM=" & FormatMetres(totalHeight) & _
        "|bias=None"
    Exit Function

Failed:
    planProof = "plan=Error:" & CStr(Err.Number)
    Set PlanPlacement = result
End Function

' R23-907. Records each view's approved scale so a later readback can prove
' none of them was reduced to force a fit.
Public Function CaptureViewScales( _
    ByRef views As Collection) As Collection

    Dim result As Collection
    Set result = New Collection
    Set CaptureViewScales = result

    On Error GoTo Failed

    Dim i As Long
    For i = 1 To views.Count
        Dim swView As SldWorks.View
        Set swView = views(i)
        result.Add CDbl(swView.ScaleDecimal)
    Next i

    Exit Function

Failed:
    Set CaptureViewScales = result
End Function

Public Function VerifyScalesUnchanged( _
    ByRef views As Collection, _
    ByRef capturedScales As Collection) As String

    On Error GoTo Failed

    Dim failures As String

    Dim i As Long
    For i = 1 To views.Count
        Dim swView As SldWorks.View
        Set swView = views(i)

        Dim current As Double
        current = CDbl(swView.ScaleDecimal)

        Dim approved As Double
        approved = CDbl(capturedScales(i))

        If Abs(current - approved) > 0.000001 Then
            failures = AppendFailure(failures, _
                "ScaleChanged:" & SafeViewName(swView) & _
                ":" & Format$(approved, "0.000000") & _
                ">" & Format$(current, "0.000000"))
        End If
    Next i

    If Len(failures) = 0 Then failures = "None"
    VerifyScalesUnchanged = "scaleFailures=" & failures
    Exit Function

Failed:
    VerifyScalesUnchanged = "scaleFailures=Error:" & CStr(Err.Number)
End Function

' R23-909. Seals the layout. Anything created afterwards changes the
' envelopes the layout was proved against, so the mutation ledger is
' photographed here and compared later.
Public Sub SealLayout(ByRef evidence As CRunEvidence)
    mLayoutSealed = True
    mSealedMutationSequence = evidence.SolidWorksMutationSequence
    mSealedLastMutation = evidence.LastSolidWorksMutation
End Sub

Public Function VerifyNothingCreatedAfterLayout( _
    ByRef evidence As CRunEvidence) As String

    If Not mLayoutSealed Then
        VerifyNothingCreatedAfterLayout = "postLayout=NotSealed"
        Exit Function
    End If

    Dim added As Long
    added = evidence.SolidWorksMutationSequence - mSealedMutationSequence

    VerifyNothingCreatedAfterLayout = "postLayout=" & _
        CStr(added = 0) & _
        "|sealedSequence=" & CStr(mSealedMutationSequence) & _
        "|currentSequence=" & _
            CStr(evidence.SolidWorksMutationSequence) & _
        "|sealedLast=" & EnvelopeToken(mSealedLastMutation) & _
        "|currentLast=" & _
            EnvelopeToken(evidence.LastSolidWorksMutation)
End Function

' R23-903 and R23-904. MUTATES THE DRAWING.
'
' Moves each view by the difference between its ENVELOPE centre and its
' assigned cell centre - not its outline centre, which is what
' Module9_LayoutEngine moves by and is precisely why annotations end up
' outside the region the layout believed it was filling.
'
' After the moves it rebuilds, re-acquires every envelope from scratch, and
' allows at most one correction pass. A view move relocates the section line
' and every annotation attached to the view, so geometry read before the
' move no longer describes the sheet.
Public Function ApplyPlacementPlan( _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef views As Collection, _
    ByRef viewEnvelopes As Collection, _
    ByRef targetCentres As Collection, _
    ByVal sheetWidth As Double, _
    ByVal sheetHeight As Double, _
    ByVal allowMutation As Boolean, _
    ByRef evidence As CRunEvidence) As String

    On Error GoTo Failed

    If Not allowMutation Then
        EmitWarning evidence, "LAYOUT_APPLY_REFUSED" & _
            "|reason=MutationNotAuthorized"
        ApplyPlacementPlan = "layout=Refused"
        Exit Function
    End If

    If views.Count <> targetCentres.Count Then
        EmitFailure evidence, "LAYOUT_APPLY_REFUSED" & _
            "|reason=PlanViewCountMismatch" & _
            "|views=" & CStr(views.Count) & _
            "|cells=" & CStr(targetCentres.Count)
        ApplyPlacementPlan = "layout=PlanMismatch"
        Exit Function
    End If

    Dim capturedScales As Collection
    Set capturedScales = CaptureViewScales(views)

    Dim pass As Long
    Dim verdict As String

    For pass = 0 To MAX_CORRECTION_PASSES
        Dim i As Long
        For i = 1 To views.Count
            Dim swView As SldWorks.View
            Set swView = views(i)

            Dim envelope As CContentEnvelope
            Set envelope = viewEnvelopes(i)

            Dim centre As Variant
            centre = targetCentres(i)

            Dim currentPosition As Variant
            currentPosition = swView.Position

            Dim target(0 To 1) As Double
            target(0) = CDbl(currentPosition(0)) + _
                (CDbl(centre(0)) - envelope.CentreX())
            target(1) = CDbl(currentPosition(1)) + _
                (CDbl(centre(1)) - envelope.CentreY())

            evidence.RecordSolidWorksMutation _
                "IView.Position:" & SafeViewName(swView)
            swView.Position = target
            evidence.LayoutMoves = evidence.LayoutMoves + 1
        Next i

        ' R23-904. Rebuild, then re-acquire everything. Nothing read before
        ' the move survives it.
        Dim rebuilt As Boolean
        rebuilt = Module11_GeometryIdentity.NormalizeSwBoolean( _
            swDraw.EditRebuild3)

        If Not rebuilt Then
            EmitWarning evidence, "LAYOUT_REBUILD|pass=" & CStr(pass) & _
                "|editRebuild3=False"
        End If

        Dim rebuiltEnvelopes As Collection
        Set rebuiltEnvelopes = New Collection

        For i = 1 To views.Count
            rebuiltEnvelopes.Add BuildViewEnvelope(views(i), _
                sheetWidth, sheetHeight, evidence)
        Next i

        Set viewEnvelopes = rebuiltEnvelopes

        Dim passRegions As Collection
        Set passRegions = BuildProtectedRegions(evidence)

        verdict = VerifyClearances(viewEnvelopes, views, passRegions)

        EmitInfo evidence, "LAYOUT_PASS|pass=" & CStr(pass) & _
            "|" & verdict

        If InStr(1, verdict, "clearanceFailures=None", _
            vbBinaryCompare) > 0 Then

            Exit For
        End If
    Next pass

    SealLayout evidence

    ApplyPlacementPlan = "layout=Applied|passes=" & CStr(pass) & _
        "|" & verdict & _
        "|" & VerifyScalesUnchanged(views, capturedScales)
    Exit Function

Failed:
    EmitFailure evidence, "LAYOUT_APPLY_ERROR|error=" & _
        CStr(Err.Number) & "|description=" & Err.Description
    ApplyPlacementPlan = "layout=Error:" & CStr(Err.Number)
End Function

' R23-900 to R23-909 evidence entry point. STRICTLY READ-ONLY: it contains
' no ApplyPlacementPlan call, no IView.Position assignment and no
' EditRebuild3.
Public Sub R23_ProbeContentEnvelope()
    On Error GoTo Failed

    mEmitDiagnostics = False
    mLayoutSealed = False

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Debug.Print "R23_ENVELOPE_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swDraw As SldWorks.ModelDoc2
    Set swDraw = swApp.ActiveDoc

    If swDraw Is Nothing Then
        Debug.Print "R23_ENVELOPE_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swDraw.GetType <> swDocDRAWING Then
        Debug.Print "R23_ENVELOPE_FATAL|reason=ActiveDocumentNotDrawing"
        Exit Sub
    End If

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet

    Dim views As Variant
    views = swSheet.GetViews

    If IsEmpty(views) Or Not IsArray(views) Then
        Debug.Print "R23_ENVELOPE_FATAL|reason=NoViewsOnSheet"
        Exit Sub
    End If

    Dim swPart As SldWorks.ModelDoc2
    Set swPart = Nothing

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim candidateView As SldWorks.View
        Set candidateView = views(i)
        If Not candidateView Is Nothing Then
            If Not candidateView.ReferencedDocument Is Nothing Then
                Set swPart = candidateView.ReferencedDocument
                Exit For
            End If
        End If
    Next i

    If swPart Is Nothing Then
        Debug.Print "R23_ENVELOPE_FATAL|reason=NoReferencedDocument"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Debug.Print "R23_ENVELOPE_FATAL|reason=UnauthorizedFixture" & _
            "|path=" & partPath
        Exit Sub
    End If

    Dim drawingSaveBefore As Boolean
    drawingSaveBefore = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Dim initialSelectionCount As Long
    initialSelectionCount = _
        swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim evidence As CRunEvidence
    Set evidence = New CRunEvidence

    Debug.Print "R23_ENVELOPE_BEGIN|drawing=" & swDraw.GetPathName & _
        "|part=" & partPath & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath) & _
        "|mode=ReadOnly|creations=0|mutations=0"

    ' The protected rectangles come from the same measurement the
    ' production pipeline uses, so the probe cannot report a boundary the
    ' pipeline does not have.
    If Not Module8_RuntimeSupport.MeasureControlledSheetRegions( _
        swSheet, evidence) Then

        Debug.Print "R23_ENVELOPE_FATAL|reason=SheetRegionsUnmeasured"
        Exit Sub
    End If

    Debug.Print "QA INFO: ENVELOPE_SHEET" & _
        "|sheetWidthM=" & FormatMetres(evidence.SheetWidth) & _
        "|sheetHeightM=" & FormatMetres(evidence.SheetHeight) & _
        "|usable=" & FormatMetres(evidence.UsableLeft) & _
        "," & FormatMetres(evidence.UsableBottom) & _
        "," & FormatMetres(evidence.UsableRight) & _
        "," & FormatMetres(evidence.UsableTop)

    Dim viewList As Collection
    Dim envelopes As Collection
    Set viewList = New Collection
    Set envelopes = New Collection

    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueView

        Dim envelope As CContentEnvelope
        Set envelope = BuildViewEnvelope(swView, _
            evidence.SheetWidth, evidence.SheetHeight, evidence)

        If Not envelope.Seeded Then GoTo ContinueView

        viewList.Add swView
        envelopes.Add envelope

        Debug.Print "QA INFO: ENVELOPE|" & envelope.Summary()

ContinueView:
    Next i

    If envelopes.Count = 0 Then
        Debug.Print "R23_ENVELOPE_FATAL|reason=NoSeededEnvelopes"
        Exit Sub
    End If

    Dim protectedRegions As Collection
    Set protectedRegions = BuildProtectedRegions(evidence)

    Dim p As Long
    For p = 1 To protectedRegions.Count
        Dim region As CContentEnvelope
        Set region = protectedRegions(p)
        Debug.Print "QA INFO: PROTECTED|" & region.Summary()
    Next p

    Dim clearanceVerdict As String
    clearanceVerdict = VerifyClearances(envelopes, viewList, _
        protectedRegions)

    Debug.Print "QA INFO: ENVELOPE_CLEARANCE|" & clearanceVerdict

    Dim planProof As String
    Dim plan As Collection
    Set plan = PlanPlacement(envelopes, evidence, planProof)

    Debug.Print "QA INFO: ENVELOPE_PLAN|" & planProof

    For i = 1 To plan.Count
        Dim cell As Variant
        cell = plan(i)

        Dim planned As CContentEnvelope
        Set planned = envelopes(i)

        Debug.Print "QA INFO: ENVELOPE_CELL|view=" & _
            EnvelopeToken(planned.Name) & _
            "|envelopeCentre=" & FormatMetres(planned.CentreX()) & _
            "," & FormatMetres(planned.CentreY()) & _
            "|cellCentre=" & FormatMetres(CDbl(cell(0))) & _
            "," & FormatMetres(CDbl(cell(1))) & _
            "|deltaX=" & FormatMetres(CDbl(cell(0)) - planned.CentreX()) & _
            "|deltaY=" & FormatMetres(CDbl(cell(1)) - planned.CentreY())
    Next i

    Dim annotationEnvelopes As Long
    For i = 1 To envelopes.Count
        Dim counted As CContentEnvelope
        Set counted = envelopes(i)
        If counted.HasAnnotationContent() Then
            annotationEnvelopes = annotationEnvelopes + 1
        End If
    Next i

    Dim finalSelectionCount As Long
    finalSelectionCount = _
        swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim drawingSaveAfter As Boolean
    drawingSaveAfter = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Debug.Print "R23_ENVELOPE_END|envelopes=" & CStr(envelopes.Count) & _
        "|annotationEnvelopes=" & CStr(annotationEnvelopes) & _
        "|protectedRegions=" & CStr(protectedRegions.Count) & _
        "|" & clearanceVerdict & _
        "|" & planProof & _
        "|creations=0|mutations=0" & _
        "|initialSelectionCount=" & CStr(initialSelectionCount) & _
        "|finalSelectionCount=" & CStr(finalSelectionCount) & _
        "|drawingUnchanged=" & _
        CStr(drawingSaveBefore = drawingSaveAfter)
    Exit Sub

Failed:
    Debug.Print "R23_ENVELOPE_FATAL|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description

    On Error Resume Next
    If Not swDraw Is Nothing Then
        swDraw.SetPickMode
        swDraw.ClearSelection2 True
    End If
End Sub
