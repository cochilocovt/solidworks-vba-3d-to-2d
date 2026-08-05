Option Explicit

' R23 typed record for one physical location as projected into one drawing
' view.
'
' A projection is only usable for annotation when it carries a selectable
' drawing-context entity. The Phase 0 evidence established that the working
' acquisition route on this build is
' IView.GetCorrespondingEntity(modelEntity); IComponent2 mediation returned
' Nothing for every counterbore edge and every body vertex tested, and
' mapping is per-edge rather than per-face. AnchorRoute records which route
' actually produced the anchor so a later regression is visible.

Public PhysicalLocation As CPhysicalHoleLocation
Public PhysicalInstanceKey As String

Public DrawingView As SldWorks.View
Public Component As SldWorks.Component2
Public ViewName As String
Public ComponentName As String
Public ReferencedConfiguration As String
Public ConfigurationProven As Boolean

' Model-context entities that were mapped, and the drawing-context entities
' they produced. Kept separate because only the drawing-context objects may
' be selected in the drawing document.
Public ModelEntityAliases As Collection
Public DrawingEntityAliases As Collection

Public PrimaryAnchor As SldWorks.Entity
Public AnchorRoute As String
Public AnchorProofSource As String
Public AnchorVisibleEntityIndex As Long

' Page-frame projected centre. Frame is stated explicitly because the
' section-line payload proved that mixing view-sketch and page values
' silently produces false clearance results.
Public PageX As Double
Public PageY As Double
Public CoordinateFrameProof As String

Public ProjectedAxisX As Double
Public ProjectedAxisY As Double
Public ProjectedAxisZ As Double
Public ProjectedRadiusM As Double
Public AxisNormalToView As Boolean

Public CoveredX As Boolean
Public CoveredY As Boolean
Public CoverageSourceX As String
Public CoverageSourceY As String

' A scheme is a read-only plan. Keep its ledger separate from the actual
' created-and-read-back coverage above so semantic QA cannot pass merely
' because it constructed a candidate ordinate group.
Public SchemeCoveredX As Boolean
Public SchemeCoveredY As Boolean
Public SchemeCoverageSourceX As String
Public SchemeCoverageSourceY As String

Public AttachedAnnotations As Collection

' Set when this location projects onto the same page point as another
' location that DID obtain an anchor in this view. Two holes on one axis seen
' from the side are one circle on the sheet, so SOLIDWORKS has a single
' drawing entity and IView.GetCorrespondingEntity maps only one of the two
' model edges to it. Without this, the second location's failure reads as a
' mapping defect rather than the geometric fact that it is.
Public CoincidentWithAnchoredKey As String

Public Accepted As Boolean
Public RejectionReason As String

Private Sub Class_Initialize()
    Set ModelEntityAliases = New Collection
    Set DrawingEntityAliases = New Collection
    Set AttachedAnnotations = New Collection

    AnchorRoute = "None"
    AnchorProofSource = vbNullString
    AnchorVisibleEntityIndex = -1
    CoordinateFrameProof = "Unproven"
    ConfigurationProven = False
    AxisNormalToView = False
    CoveredX = False
    CoveredY = False
    CoverageSourceX = vbNullString
    CoverageSourceY = vbNullString
    SchemeCoveredX = False
    SchemeCoveredY = False
    SchemeCoverageSourceX = vbNullString
    SchemeCoverageSourceY = vbNullString
    CoincidentWithAnchoredKey = vbNullString
    Accepted = False
    RejectionReason = "NotEvaluated"
End Sub

Public Function ProjectionKey() As String
    ProjectionKey = Module11_GeometryIdentity.BuildViewProjectionKey( _
        ViewName, PhysicalInstanceKey)
End Function

Public Function HasSelectableAnchor() As Boolean
    HasSelectableAnchor = Not (PrimaryAnchor Is Nothing)
End Function

Public Sub RecordModelAlias(ByVal modelEntity As Object)
    If modelEntity Is Nothing Then Exit Sub
    ModelEntityAliases.Add modelEntity
End Sub

Public Sub RecordDrawingAlias(ByVal drawingEntity As Object)
    If drawingEntity Is Nothing Then Exit Sub
    DrawingEntityAliases.Add drawingEntity
End Sub

Public Sub AttachAnnotation( _
    ByRef importedAnnotation As CImportedAnnotation)

    If importedAnnotation Is Nothing Then Exit Sub
    AttachedAnnotations.Add importedAnnotation
End Sub

' A projection may only be published when its anchor, coordinate frame and
' referenced configuration are all proved. Anything short of that fails
' closed with ProjectionAnchorUnavailable rather than silently dropping the
' physical location.
Public Function QualificationFailureReason() As String
    If PhysicalLocation Is Nothing Then
        QualificationFailureReason = "PhysicalLocationMissing"
        Exit Function
    End If

    If Len(Trim$(PhysicalInstanceKey)) = 0 Then
        QualificationFailureReason = "PhysicalInstanceKeyMissing"
        Exit Function
    End If

    If Not ConfigurationProven Then
        QualificationFailureReason = "ReferencedConfigurationUnproven"
        Exit Function
    End If

    If Not HasSelectableAnchor() Then
        QualificationFailureReason = "ProjectionAnchorUnavailable"
        Exit Function
    End If

    ' R23-306. A hole only projects as a circle when its axis is parallel to
    ' the view normal. An anchor can still be mapped when it is not: the
    ' first anchored run accepted all four M5 tapped holes in the front view,
    ' where their Y axis lies in the page plane and the "circular" edge is
    ' actually seen edge-on. Such a projection must not be published as a
    ' circular anchor, but the mapped anchor is real, so it fails with its
    ' own reason rather than ProjectionAnchorUnavailable.
    If Not AxisNormalToView Then
        QualificationFailureReason = "AxisNotNormalToView"
        Exit Function
    End If

    If StrComp(CoordinateFrameProof, "Unproven", vbTextCompare) = 0 Then
        QualificationFailureReason = "CoordinateFrameUnproven"
        Exit Function
    End If

    QualificationFailureReason = vbNullString
End Function

' A projection may carry a proved page POSITION without carrying a
' selectable anchor. Measured 2026-08-05 in run
' macro_qa/20260805_033146_P-0251-14A-001: the stepped bore of P-0251 maps
' all four of its circular edges through Route D
' (ISelectionMgr.GetSelectedObjectsDrawingView2 proves Drawing View1 owns
' them) yet none appears in IView.GetVisibleEntities2, which by contract
' holds only entities "not completely obscured by other entities in the
' view". The bore is hidden behind the front face in every orthographic
' view, so no orthographic anchor for it can ever exist - that feature is
' shown in the J-J section, exactly as the reference drawing does.
'
' Consumers that must SELECT the projection - dimensioning, callout
' attachment, ordinate anchoring - keep requiring QualificationFailureReason
' and are unaffected by this. Consumers that need only WHERE the feature
' lands on the page - the section-line waypoints, which read PageX, PageY
' and ProjectedRadiusM and never touch PrimaryAnchor - may use this weaker
' proof instead. User decision, 2026-08-05.
'
' Everything except the anchor is still proved: the physical identity, the
' referenced configuration, the page coordinate frame, and the fact that the
' axis is normal to the view, without which the location does not project as
' a circle at all and its "centre" would be meaningless.
Public Function PositionFailureReason() As String
    If PhysicalLocation Is Nothing Then
        PositionFailureReason = "PhysicalLocationMissing"
        Exit Function
    End If

    If Len(Trim$(PhysicalInstanceKey)) = 0 Then
        PositionFailureReason = "PhysicalInstanceKeyMissing"
        Exit Function
    End If

    If Not ConfigurationProven Then
        PositionFailureReason = "ReferencedConfigurationUnproven"
        Exit Function
    End If

    If Not AxisNormalToView Then
        PositionFailureReason = "AxisNotNormalToView"
        Exit Function
    End If

    If StrComp(CoordinateFrameProof, "Unproven", vbTextCompare) = 0 Then
        PositionFailureReason = "CoordinateFrameUnproven"
        Exit Function
    End If

    PositionFailureReason = vbNullString
End Function

Public Function HasProvedPosition() As Boolean
    HasProvedPosition = (Len(PositionFailureReason()) = 0)
End Function

Public Function CoverageSummary() As String
    CoverageSummary = _
        "view=" & ViewName & _
        "|physical=" & PhysicalInstanceKey & _
        "|frame=Page" & _
        "|pageX=" & Format$(PageX, "0.000000000") & _
        "|pageY=" & Format$(PageY, "0.000000000") & _
        "|anchorRoute=" & AnchorRoute & _
        "|anchor=" & CStr(HasSelectableAnchor()) & _
        "|axisNormal=" & CStr(AxisNormalToView) & _
        "|projectedRadiusM=" & Format$(ProjectedRadiusM, "0.000000000") & _
        "|coveredX=" & CStr(CoveredX) & _
        "|coveredY=" & CStr(CoveredY) & _
        "|schemeCoveredX=" & CStr(SchemeCoveredX) & _
        "|schemeCoveredY=" & CStr(SchemeCoveredY) & _
        "|annotations=" & CStr(AttachedAnnotations.Count)
End Function
