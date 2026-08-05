Option Explicit

' R23 Phase 3. Resolves each physical location in the location graph into a
' CViewHoleProjection for one drawing view: a selectable drawing-context
' anchor, a page-frame centre, and the proofs that justify both.
' Everything here is read-only against the model and the drawing. No
' annotation is created, no entity is left selected, and no document is
' saved. Selection proof refuses a non-empty caller selection rather than
' clearing state it cannot restore exactly.
' The acquisition routes are the ones Phase 0 settled on this build, not the
' ones the documentation implies:
' Route A  IView.GetCorrespondingEntity(modelEntity)  -- works.
' Route B  IComponent2.GetCorrespondingEntity         -- returned Nothing
' for every counterbore edge and every body vertex tested, so it
' is attempted and recorded but never depended on.
' Route C  IView.GetVisibleEntities2 + ISldWorks.IsSame -- not an
' acquisition route. It is the independent confirmation that what
' route A returned is genuinely a drawing-context entity of this
' view, and it supplies the visible-inventory index for evidence.
' Route D  IView.SelectEntity + SelectionMgr -- guarded alternative only
' when Route A declines and a part drawing exposes no inventory.
' Mapping is per-EDGE, not per-face. For each counterbore only one of the two
' owned circular edges maps, so every edge of every contributing face is
' tried before a location is failed.

' Drawing entity types. Confirmed values of swViewEntityType_e.
Private Const VIEW_ENTITY_EDGE As Long = 1
Private Const VIEW_ENTITY_VERTEX As Long = 2
Private Const VIEW_ENTITY_FACE As Long = 3
Private Const VIEW_ENTITY_SILHOUETTE As Long = 4

' Circular-projection acceptance. A projected hole centre is only usable if
' the mapped drawing edge is a complete circle, so the curve must close on
' itself within this tolerance in metres.
Private Const CIRCLE_CLOSURE_TOLERANCE_M As Double = 0.000001

' Axis-to-view normality. The projected axis must be parallel to the view
' normal for the hole to appear as a circle rather than an ellipse.
Private Const AXIS_NORMAL_TOLERANCE As Double = 0.001

' ISldWorks.IsSame returns swObjectEquality, NOT a Boolean: 0 not same,
' 1 same, 2 unable to determine. It must never be read through
' NormalizeSwBoolean, which would treat "unable to determine" as a match.
' Only an exact 1 proves identity.
Private Const OBJECT_EQUALITY_SAME As Long = 1
Private Const OBJECT_EQUALITY_UNSUPPORTED As Long = 2

' swCrossHatchFilter_e.swCrossHatchExclude. MCP corpus value; verify on the
' installed SOLIDWORKS 2025 build through the read-only probe runner.
Private Const CROSSHATCH_EXCLUDE As Long = 1

' R23-308 anchor priority. Lower is better.
' Tier 1, the imported native-callout attachment, cannot be resolved yet:
' nothing has been imported until Phase 4 attaches annotations to the
' projection. The tier is defined here so the ordering is complete and the
' Phase 4 hook is explicit, and AnchorTierFor records it as unavailable
' rather than silently collapsing the order to two tiers.
Private Const ANCHOR_TIER_NATIVE_CALLOUT As Long = 1
Private Const ANCHOR_TIER_PRIMARY_DIAMETER As Long = 2
Private Const ANCHOR_TIER_SMALLEST_CIRCLE As Long = 3
Private Const ANCHOR_TIER_NONE As Long = 99

' Page-space coincidence tolerance in metres. Two locations closer than this
' on the sheet are the same circle to the drawing.
Private Const PAGE_COINCIDENCE_M As Double = 0.000001

Private mEmitDiagnostics As Boolean

Private Sub EmitInfo( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddInfo message
    If mEmitDiagnostics Then
        Module21_EvidenceSink.LogLine message
    End If
End Sub

Private Sub EmitWarning( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddWarning message
    If mEmitDiagnostics Then
        Module21_EvidenceSink.LogLine message
    End If
End Sub

Private Sub EmitFailure( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddFailure message
    If mEmitDiagnostics Then
        Module21_EvidenceSink.LogLine message
    End If
End Sub

' Builds one CViewHoleProjection per physical location for the supplied view
' and adds it to the graph. Returns True when the pass completed; individual
' locations that cannot be projected are recorded with an explicit reason
' rather than dropped, so a later phase can report what is missing.
Public Function BuildViewProjections( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If swApp Is Nothing Then Exit Function
    If swView Is Nothing Then Exit Function
    If graph Is Nothing Then Exit Function

    Dim viewName As String
    viewName = SafeViewName(swView)

    Dim referencedConfiguration As String
    Dim configurationProven As Boolean
    configurationProven = ReadReferencedConfiguration( _
        swView, referencedConfiguration)

    Dim drawingComponent As SldWorks.Component2
    Dim modelComponent As SldWorks.Component2
    Dim componentProof As String
    ResolveViewComponentContexts _
        swView, drawingComponent, modelComponent, componentProof

    EmitInfo evidence, "PROJECTION_VIEW|view=" & viewName & _
        "|viewType=" & CStr(SafeViewType(swView)) & _
        "|displayMode=" & DisplayModeName(SafeDisplayMode(swView)) & _
        "|referencedConfiguration=" & referencedConfiguration & _
        "|configurationProven=" & CStr(configurationProven) & _
        "|" & componentProof

    ' GetCorrespondingEntity is view-scoped. A runner opens a drawing without
    ' making any individual view active, unlike the interactive Phase 0 pass.
    ' Activating the target view changes UI focus only; it does not create,
    ' move, dimension, or save the drawing. The result is recorded because a
    ' failed activation means Route A cannot be interpreted as a map failure.
    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim viewActivated As Boolean
    On Error Resume Next
    viewActivated = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDrawing.ActivateView(viewName))
    Dim activationError As Long
    activationError = Err.Number
    Err.Clear
    On Error GoTo Failed

    EmitInfo evidence, "PROJECTION_VIEW_ACTIVATION|view=" & viewName & _
        "|succeeded=" & CStr(viewActivated) & _
        "|error=" & CStr(activationError)

    Dim visibleEdges As Variant
    visibleEdges = ReadVisibleEntities( _
        swView, drawingComponent, VIEW_ENTITY_EDGE)

    Dim visibleCount As Long
    visibleCount = VariantItemCount(visibleEdges)

    Dim visibleInventoryAvailable As Boolean
    visibleInventoryAvailable = (visibleCount > 0)

    EmitInfo evidence, "PROJECTION_VISIBLE|view=" & viewName & _
        "|type=Edge" & _
        "|count=" & CStr(visibleCount) & _
        "|source=IView.GetVisibleEntities2"

    ' A sheet's standard-view placeholders are returned by ISheet.GetViews
    ' alongside real views. They carry no drawing geometry, so no location
    ' can anchor in them. Do not apply that rule to a part drawing's real
    ' view: GetVisibleEntities2 requires a Component2, and this build gives
    ' a part drawing no Component2 even when the view is visibly populated.
    If visibleCount = 0 And IsSheetOrientationPlaceholder(swView) Then
        EmitInfo evidence, "PROJECTION_VIEW_SKIPPED|view=" & viewName & _
            "|viewType=" & CStr(SafeViewType(swView)) & _
            "|reason=SheetOrientationPlaceholder"
        BuildViewProjections = True
        Exit Function
    End If

    If visibleCount = 0 And Not drawingComponent Is Nothing Then
        EmitInfo evidence, "PROJECTION_VIEW_SKIPPED|view=" & viewName & _
            "|viewType=" & CStr(SafeViewType(swView)) & _
            "|reason=NoVisibleDrawingEntities"
        BuildViewProjections = True
        Exit Function
    End If

    If visibleCount = 0 Then
        EmitWarning evidence, "PROJECTION_INVENTORY_UNAVAILABLE|view=" & _
            viewName & "|reason=PartDrawingHasNoComponent2" & _
            "|confirmation=SelectionOwnership"
    End If

    Dim locations As Collection
    Set locations = graph.Locations()

    Dim i As Long
    For i = 1 To locations.Count
        Dim location As CPhysicalHoleLocation
        Set location = locations(i)

        Dim projection As CViewHoleProjection
        Set projection = New CViewHoleProjection

        Set projection.PhysicalLocation = location
        projection.PhysicalInstanceKey = location.PhysicalInstanceKey
        Set projection.DrawingView = swView
        projection.ViewName = viewName
        Set projection.Component = drawingComponent
        projection.ComponentName = SafeComponentName(drawingComponent)
        projection.ReferencedConfiguration = referencedConfiguration
        projection.ConfigurationProven = configurationProven

        ResolveProjection swApp, swDraw, swView, location, visibleEdges, _
            visibleInventoryAvailable, drawingComponent, projection, evidence

        Dim failureReason As String
        failureReason = projection.QualificationFailureReason()

        ' GetVisibleEntities2 cannot inventory a part drawing with no
        ' Component2. The mapped object is still a drawing-context IEntity,
        ' so prove its owner through the selection manager before accepting
        ' it. Selection is cleared inside SelectAnchorInView on every path.
        If Len(failureReason) = 0 And Not visibleInventoryAvailable Then
            Dim confirmationProof As String
            If SelectAnchorInView(swDraw, projection, confirmationProof) Then
                projection.AnchorProofSource = _
                    projection.AnchorProofSource & _
                    "|inventory=Unavailable" & _
                    "|confirmation=SelectionOwnership"
                EmitInfo evidence, "PROJECTION_ROUTE_CONFIRMATION|view=" & _
                    viewName & "|physical=" & _
                    projection.PhysicalInstanceKey & _
                    "|source=ISelectionMgr.GetSelectedObjectsDrawingView2" & _
                    "|proved=True"
            Else
                failureReason = "SelectionOwnershipUnproven"
                EmitWarning evidence, "PROJECTION_ROUTE_CONFIRMATION|view=" & _
                    viewName & "|physical=" & _
                    projection.PhysicalInstanceKey & _
                    "|source=ISelectionMgr.GetSelectedObjectsDrawingView2" & _
                    "|proved=False|" & confirmationProof
            End If
        End If

        If Len(failureReason) = 0 Then
            projection.Accepted = True
            projection.RejectionReason = vbNullString
        Else
            projection.Accepted = False
            projection.RejectionReason = failureReason
        End If

        graph.AddProjection projection

        EmitInfo evidence, "PROJECTION|" & projection.CoverageSummary() & _
            "|accepted=" & CStr(projection.Accepted) & _
            "|reason=" & IIf(Len(failureReason) = 0, "None", failureReason)
    Next i

    MarkCoincidentProjections graph, viewName, evidence

    BuildViewProjections = True
    Exit Function

Failed:
    ' Capture first: SafeViewName contains On Error Resume Next, which
    ' resets the global Err before the concatenation reaches Err.Number.
    Dim passErrorNumber As Long
    Dim passErrorDescription As String
    passErrorNumber = Err.Number
    passErrorDescription = Err.Description

    EmitFailure evidence, "PROJECTION_PASS_ERROR|view=" & _
        SafeViewName(swView) & _
        "|error=" & CStr(passErrorNumber) & _
        "|description=" & passErrorDescription
End Function

' Attributes each unanchored projection that shares a page point with an
' anchored one, so the failure is reported as geometry rather than as a
' mapping defect.
' Two holes on a common axis, viewed along that axis, are ONE circle on the
' sheet. SOLIDWORKS therefore has a single drawing entity and
' IView.GetCorrespondingEntity maps only one of the two model edges to it.
' In the P-0251 side view this is exactly why six counterbores yield three
' anchors and four tapped holes yield two: the mapped count tracks distinct
' page positions, not holes. No search strategy can produce more anchors
' than the drawing has entities.
Private Sub MarkCoincidentProjections( _
    ByRef graph As CLocationGraph, _
    ByVal viewName As String, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim projections As Collection
    Set projections = graph.ProjectionsForView(viewName)

    Dim i As Long
    Dim j As Long

    For i = 1 To projections.Count
        Dim candidate As CViewHoleProjection
        Set candidate = projections(i)

        If candidate.HasSelectableAnchor() Then GoTo ContinueCandidate

        If candidate.CoordinateFrameProof = "Unproven" Then
            GoTo ContinueCandidate
        End If

        For j = 1 To projections.Count
            Dim anchored As CViewHoleProjection
            Set anchored = projections(j)

            If Not anchored.HasSelectableAnchor() Then GoTo ContinueAnchored

            If Module11_GeometryIdentity.ValuesMatchWithin( _
                candidate.PageX, anchored.PageX, PAGE_COINCIDENCE_M) Then

                If Module11_GeometryIdentity.ValuesMatchWithin( _
                    candidate.PageY, anchored.PageY, PAGE_COINCIDENCE_M) Then

                    candidate.CoincidentWithAnchoredKey = _
                        anchored.PhysicalInstanceKey

                    EmitInfo evidence, "PROJECTION_COINCIDENT|view=" & _
                        viewName & _
                        "|physical=" & candidate.PhysicalInstanceKey & _
                        "|pageX=" & Format$(candidate.PageX, "0.000000000") & _
                        "|pageY=" & Format$(candidate.PageY, "0.000000000") & _
                        "|sharesPagePointWith=" & _
                            anchored.PhysicalInstanceKey & _
                        "|reason=OneDrawingEntityForTwoCoaxialHoles"
                    GoTo ContinueCandidate
                End If
            End If

ContinueAnchored:
        Next j

ContinueCandidate:
    Next i
    Exit Sub

Failed:
    EmitWarning evidence, "PROJECTION_COINCIDENCE_ERROR|view=" & viewName & _
        "|error=" & CStr(Err.Number)
End Sub

' R23-300. The two component handles are kept explicitly separate.
' IView.GetVisibleComponents returns the limited drawing-context Component2
' that GetVisibleEntities2 requires. IView.GetVisibleDrawingComponents is
' documented for ASSEMBLY drawings and yields IDrawingComponent, whose
' .Component is a full model-capable IComponent2. A part drawing is expected
' to supply the first and not the second, so the absence of the model-capable
' handle is recorded as context rather than treated as a failure. When both
' are present their names are compared so a divergence becomes visible.
Private Sub ResolveViewComponentContexts( _
    ByRef swView As SldWorks.View, _
    ByRef drawingComponent As SldWorks.Component2, _
    ByRef modelComponent As SldWorks.Component2, _
    ByRef componentProof As String)

    On Error GoTo Failed

    componentProof = "componentContext=Unresolved"

    Set drawingComponent = SingleVisibleComponent(swView)
    Set modelComponent = SingleModelCapableComponent(swView)

    Dim drawingName As String
    Dim modelName As String
    drawingName = SafeComponentName(drawingComponent)
    modelName = SafeComponentName(modelComponent)

    Dim convergence As String

    If drawingComponent Is Nothing And modelComponent Is Nothing Then
        convergence = "NoComponentContext"
    ElseIf modelComponent Is Nothing Then
        convergence = "DrawingContextOnly"
    ElseIf drawingComponent Is Nothing Then
        convergence = "ModelContextOnly"
    ElseIf StrComp( _
        Module11_GeometryIdentity.IdentityToken(drawingName), _
        Module11_GeometryIdentity.IdentityToken(modelName), _
        vbBinaryCompare) = 0 Then

        convergence = "Converged"
    Else
        convergence = "Diverged"
    End If

    componentProof = "componentContext=" & convergence & _
        "|drawingComponent=" & drawingName & _
        "|drawingSource=IView.GetVisibleComponents" & _
        "|modelComponent=" & modelName & _
        "|modelSource=IView.GetVisibleDrawingComponents.Component"
    Exit Sub

Failed:
    componentProof = "componentContext=ReadError:" & CStr(Err.Number)
End Sub

Private Function SingleVisibleComponent( _
    ByRef swView As SldWorks.View) As SldWorks.Component2

    On Error GoTo Failed

    Dim visible As Variant
    visible = swView.GetVisibleComponents

    If IsEmpty(visible) Then Exit Function
    If Not IsArray(visible) Then Exit Function

    Dim count As Long
    Dim i As Long

    For i = LBound(visible) To UBound(visible)
        Dim component As SldWorks.Component2
        Set component = Nothing
        On Error Resume Next
        Set component = visible(i)
        On Error GoTo Failed

        If Not component Is Nothing Then
            count = count + 1
            If count = 1 Then Set SingleVisibleComponent = component
        End If
    Next i

    ' More than one component is an assembly context this phase does not
    ' claim to resolve; refuse rather than pick an arbitrary one.
    If count <> 1 Then Set SingleVisibleComponent = Nothing
    Exit Function

Failed:
    Set SingleVisibleComponent = Nothing
End Function

Private Function SingleModelCapableComponent( _
    ByRef swView As SldWorks.View) As SldWorks.Component2

    On Error GoTo Failed

    Dim visible As Variant
    visible = swView.GetVisibleDrawingComponents

    If IsEmpty(visible) Then Exit Function
    If Not IsArray(visible) Then Exit Function

    Dim count As Long
    Dim i As Long

    For i = LBound(visible) To UBound(visible)
        Dim drawingComponent As SldWorks.DrawingComponent
        Set drawingComponent = Nothing
        On Error Resume Next
        Set drawingComponent = visible(i)
        On Error GoTo Failed

        If Not drawingComponent Is Nothing Then
            Dim component As SldWorks.Component2
            Set component = Nothing
            On Error Resume Next
            Set component = drawingComponent.Component
            On Error GoTo Failed

            If Not component Is Nothing Then
                count = count + 1
                If count = 1 Then Set SingleModelCapableComponent = component
            End If
        End If
    Next i

    If count <> 1 Then Set SingleModelCapableComponent = Nothing
    Exit Function

Failed:
    Set SingleModelCapableComponent = Nothing
End Function

' R23-301. Drawing-context entities for the view. The component argument is
' passed through exactly as obtained; a part drawing may legitimately supply
' Nothing and the call still returns the view's own entities.
Private Function ReadVisibleEntities( _
    ByRef swView As SldWorks.View, _
    ByRef component As SldWorks.Component2, _
    ByVal entityType As Long) As Variant

    On Error GoTo Failed

    ReadVisibleEntities = Empty

    Dim entities As Variant
    entities = swView.GetVisibleEntities2(component, entityType)

    If IsEmpty(entities) Then Exit Function
    If Not IsArray(entities) Then Exit Function

    ReadVisibleEntities = entities
    Exit Function

Failed:
    ReadVisibleEntities = Empty
End Function

' The named orientation entries returned by ISheet.GetViews begin with *.
' This is a sheet-list fact, not a drawing-view-type guess.
Private Function IsSheetOrientationPlaceholder( _
    ByRef swView As SldWorks.View) As Boolean

    IsSheetOrientationPlaceholder = _
        (Left$(SafeViewName(swView), 1) = "*")
End Function

' R23-302 through R23-308. Walks every contributing face and every boundary
' edge of that face, keeping every candidate that maps into drawing context
' AND proves a complete circle, then picks by the R23-308 priority order.
' Every candidate is evaluated rather than stopping at the first: taking the
' first mappable edge would make the anchor depend on face and edge order,
' so a counterbore could be dimensioned on its 11 mm mouth instead of its
' 6.6 mm through hole purely by traversal accident.
Private Sub ResolveProjection( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef location As CPhysicalHoleLocation, _
    ByVal visibleEdges As Variant, _
    ByVal visibleInventoryAvailable As Boolean, _
    ByRef component As SldWorks.Component2, _
    ByRef projection As CViewHoleProjection, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    If location Is Nothing Then Exit Sub

    ' R23-306. Axis compatibility is a property of the location and the view,
    ' independent of any edge, so it is proved once up front.
    Dim projectedAxisX As Double
    Dim projectedAxisY As Double
    Dim projectedAxisZ As Double

    projection.AxisNormalToView = AxisIsNormalToView( _
        swApp, swView, location, _
        projectedAxisX, projectedAxisY, projectedAxisZ)

    ' The out-parameters must be locals. A class's Public variable is exposed
    ' as a property, so passing projection.ProjectedAxisX directly as a ByRef
    ' argument hands the callee a temporary that is discarded on return: the
    ' fourth live run logged projectedAxis=0,0,0 on every line while
    ' axisNormal was correct, because the function's own locals were sound
    ' and only the write-back was lost.
    projection.ProjectedAxisX = projectedAxisX
    projection.ProjectedAxisY = projectedAxisY
    projection.ProjectedAxisZ = projectedAxisZ

    projection.ProjectedRadiusM = location.PrimaryRadiusM

    Dim bestTier As Long
    Dim bestRadiusM As Double
    bestTier = ANCHOR_TIER_NONE
    bestRadiusM = 0#

    ' Stage counters. The first catalog projection run reported only
    ' candidates=0, which could not distinguish "no faces retained" from
    ' "no circular edge" from "nothing mapped". Each stage is counted so one
    ' run isolates the break, the way the Phase 0 rejectGate did.
    Dim stageFaces As Long
    Dim stageProjected As Long
    Dim stageEdges As Long
    Dim stageCircles As Long
    Dim stageMapped As Long
    Dim stageConfirmed As Long
    Dim firstReject As String

    ' Phase 2 diagnosis, read-only. Both GetCorrespondingEntity and
    ' GetVisibleEntities2 only ever see entities that are "not completely
    ' obscured by other entities in the view" (2025 Help, GetVisibleEntities2
    ' Remarks). So a circular edge that no route maps may be obscured
    ' geometry behaving correctly rather than a mapping defect, and the two
    ' cases need different fixes. GetPolylines7 returns the visible MODEL
    ' edges, so the candidate can be classified without first needing the
    ' drawing entity that is missing. Counting only; acceptance is unchanged.
    ' Counts edges that declined every route. The r40-r42 visibility
    ' classifiers that used to sit here were an instrument for one question -
    ' what GetPolylines7 actually returns - and that question is now settled
    ' and recorded in SOLIDWORKS_API_VALIDATION.md. They are removed rather
    ' than kept: the model-space variant measured nothing, and the
    ' drawing-space control cost a GetPolylines7 read per mapped edge.
    Dim stageUnmappedAllRoutes As Long
    Dim firstUnmappedRoute As String
    firstUnmappedRoute = "None"

    stageFaces = location.SourceFaces.Count
    firstReject = "None"

    Dim faceIndex As Long
    For faceIndex = 1 To location.SourceFaces.Count
        Dim modelFace As SldWorks.Face2
        Set modelFace = Nothing
        On Error Resume Next
        Set modelFace = location.SourceFaces(faceIndex)
        On Error GoTo Failed

        If modelFace Is Nothing Then GoTo ContinueFace

        Dim pageX As Double
        Dim pageY As Double
        Dim pageZ As Double
        Dim frameProof As String

        If Not ProjectFaceCentreToPage( _
            swApp, swView, modelFace, pageX, pageY, pageZ, frameProof) Then

            If firstReject = "None" Then
                firstReject = "FaceCentreNotProjected:" & frameProof
            End If
            GoTo ContinueFace
        End If

        stageProjected = stageProjected + 1

        ' Record the page centre from the first face that projects, whether
        ' or not an anchor is ever found. An unanchored location still has a
        ' provable position, and without it a coincident projection cannot be
        ' distinguished from a mapping defect. Acceptance is unaffected: the
        ' anchor test runs before the coordinate-frame test.
        If projection.CoordinateFrameProof = "Unproven" Then
            projection.PageX = pageX
            projection.PageY = pageY
            projection.CoordinateFrameProof = frameProof
        End If

        Dim edges As Variant
        On Error Resume Next
        edges = modelFace.GetEdges
        On Error GoTo Failed

        If IsEmpty(edges) Or Not IsArray(edges) Then
            If firstReject = "None" Then firstReject = "FaceEdgesUnavailable"
            GoTo ContinueFace
        End If

        stageEdges = stageEdges + VariantItemCount(edges)

        Dim edgeIndex As Long
        For edgeIndex = LBound(edges) To UBound(edges)
            Dim modelEdge As SldWorks.Edge
            Set modelEdge = Nothing
            On Error Resume Next
            Set modelEdge = edges(edgeIndex)
            On Error GoTo Failed

            If modelEdge Is Nothing Then GoTo ContinueEdge

            ' Qualification is not weakened to "any visible circle": the edge
            ' must be a complete circle owned by this location's own face.
            Dim edgeRadiusM As Double
            Dim circleProof As String
            If Not EdgeIsCompleteCircle( _
                modelEdge, edgeRadiusM, circleProof) Then

                If firstReject = "None" Then firstReject = circleProof
                GoTo ContinueEdge
            End If

            stageCircles = stageCircles + 1
            projection.RecordModelAlias modelEdge

            Dim mappedRoute As String
            Dim mapped As SldWorks.Entity
            Set mapped = MapModelEntityToDrawing( _
                swView, component, modelEdge, mappedRoute)

            ' Route D uses the documented view-selection API and accepts its
            ' result only after target-view ownership is read back from the
            ' selection manager; its temporary selection is cleared on every
            ' path.
            '
            ' The gate used to be "Not visibleInventoryAvailable", which is
            ' the wrong condition and is the root of the P-0251 section
            ' failure. Drawing View1 HAS an inventory (39 edges), so Route D
            ' was never attempted there - even though Route A demonstrably
            ' declines for real geometry in that same view: the six
            ' counterbores map two edges each while the stepped bore maps
            ' none of its four, with A:Nothing:err0 on every one. With no
            ' second route, that bore has no accepted projection, so
            ' ResolveBoreProjection finds no singleton bore and no section
            ' path resolves.
            '
            ' Keying it on "Route A and B both declined" is not a weakening.
            ' A Route-D entity still has to pass the Route C inventory check
            ' immediately below when an inventory exists, so it is proved
            ' twice - ownership through ISelectionMgr.GetSelectedObjectsDrawingView2
            ' AND membership of IView.GetVisibleEntities2 - where the
            ' inventory-less path proves only ownership. An edge that is
            ' genuinely obscured cannot be in the visible inventory, so it is
            ' still rejected as MappedEntityNotInVisibleInventory. That
            ' rejection is now a real answer to the open hidden-geometry
            ' question rather than the silence of a route never tried.
            If mapped Is Nothing Then
                Dim selectionRoute As String
                Set mapped = SelectModelEntityInView( _
                    swDraw, swView, modelEdge, selectionRoute)

                If Not mapped Is Nothing Then
                    mappedRoute = selectionRoute
                Else
                    mappedRoute = mappedRoute & ";" & selectionRoute
                End If
            End If

            ' An edge that reaches here declined Route A, Route B and Route D.
            ' r40 and r41 classified these against GetPolylines7 with the
            ' MODEL edge and reported them as "obscured"; r42's control proved
            ' that comparison matches nothing at all, so the label was void
            ' and has been removed rather than left to mislead. Drawing-space
            ' classification is not possible here by construction: it needs a
            ' drawing entity, and having none is precisely the condition being
            ' reported.
            If mapped Is Nothing Then
                stageUnmappedAllRoutes = stageUnmappedAllRoutes + 1

                If firstUnmappedRoute = "None" Then
                    firstUnmappedRoute = mappedRoute
                End If

                If firstReject = "None" Then
                    firstReject = "NoRouteMappedThisEdge:" & mappedRoute
                End If
                GoTo ContinueEdge
            End If

            stageMapped = stageMapped + 1
            projection.RecordDrawingAlias mapped

            ' R23-304. Route C must independently agree that the mapped
            ' object is a drawing-context entity of this view before it is
            ' accepted as an anchor.
            Dim visibleIndex As Long
            Dim unsupportedComparisons As Long
            visibleIndex = -1

            If visibleInventoryAvailable Then
                visibleIndex = FindVisibleEntityIndex( _
                    swApp, visibleEdges, mapped, unsupportedComparisons)

                If visibleIndex < 0 Then
                    EmitWarning evidence, _
                        "PROJECTION_ROUTE_DISAGREEMENT|view=" & _
                        projection.ViewName & _
                        "|physical=" & projection.PhysicalInstanceKey & _
                        "|route=" & mappedRoute & _
                        "|unsupportedComparisons=" & _
                            CStr(unsupportedComparisons) & _
                        "|reason=MappedEntityNotInVisibleInventory"

                    If firstReject = "None" Then
                        firstReject = "MappedEntityNotInVisibleInventory"
                    End If
                    GoTo ContinueEdge
                End If

                stageConfirmed = stageConfirmed + 1
            End If

            Dim candidateTier As Long
            candidateTier = AnchorTierFor(location, edgeRadiusM)

            If BetterAnchor( _
                candidateTier, edgeRadiusM, bestTier, bestRadiusM) Then

                bestTier = candidateTier
                bestRadiusM = edgeRadiusM

                Set projection.PrimaryAnchor = mapped
                projection.AnchorRoute = mappedRoute
                projection.AnchorVisibleEntityIndex = visibleIndex
                projection.AnchorProofSource = _
                    "anchorTier=" & AnchorTierName(candidateTier) & _
                    "|faceIndex=" & CStr(faceIndex) & _
                    "|edgeIndex=" & CStr(edgeIndex) & _
                    "|" & circleProof & _
                    "|edgeRadiusM=" & _
                        Format$(edgeRadiusM, "0.000000000") & _
                    "|visibleIndex=" & CStr(visibleIndex) & _
                    "|identity=" & IIf(visibleInventoryAvailable, _
                        "ISldWorks.IsSame", "SelectionOwnershipPending")

                projection.PageX = pageX
                projection.PageY = pageY
                projection.CoordinateFrameProof = frameProof
            End If


ContinueEdge:
        Next edgeIndex

ContinueFace:
    Next faceIndex

    EmitInfo evidence, "PROJECTION_ANCHOR|view=" & projection.ViewName & _
        "|physical=" & projection.PhysicalInstanceKey & _
        "|axisNormal=" & CStr(projection.AxisNormalToView) & _
        "|projectedAxis=" & _
            Format$(projection.ProjectedAxisX, "0.000000") & "," & _
            Format$(projection.ProjectedAxisY, "0.000000") & "," & _
            Format$(projection.ProjectedAxisZ, "0.000000") & _
        "|sourceFaces=" & CStr(stageFaces) & _
        "|facesProjected=" & CStr(stageProjected) & _
        "|boundaryEdges=" & CStr(stageEdges) & _
        "|circularEdges=" & CStr(stageCircles) & _
        "|mappedEdges=" & CStr(stageMapped) & _
        "|inventoryConfirmed=" & CStr(stageConfirmed) & _
        "|unmappedAllRoutes=" & CStr(stageUnmappedAllRoutes) & _
        "|firstUnmappedRoute=" & firstUnmappedRoute & _
        "|firstReject=" & firstReject & _
        "|chosenTier=" & AnchorTierName(bestTier) & _
        "|" & IIf(Len(projection.AnchorProofSource) = 0, _
            "anchorProof=None", projection.AnchorProofSource)
    Exit Sub

Failed:
    EmitWarning evidence, "PROJECTION_RESOLVE_ERROR|view=" & _
        projection.ViewName & _
        "|physical=" & projection.PhysicalInstanceKey & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description
End Sub

' R23-308. Classifies one mappable circular edge against the anchor priority
' order. The primary typed hole diameter is the location's smallest coaxial
' radius, which is the feature a manufacturing callout must dimension: for a
' counterbore that is the through hole, not the 11 mm mouth.
Private Function AnchorTierFor( _
    ByRef location As CPhysicalHoleLocation, _
    ByVal edgeRadiusM As Double) As Long

    If Module11_GeometryIdentity.RadiiMatch( _
        edgeRadiusM, location.PrimaryRadiusM) Then

        AnchorTierFor = ANCHOR_TIER_PRIMARY_DIAMETER
        Exit Function
    End If

    AnchorTierFor = ANCHOR_TIER_SMALLEST_CIRCLE
End Function

' Within a tier the smaller circle wins, which makes the choice depend on
' geometry rather than on the order faces and edges happen to be returned in.
Private Function BetterAnchor( _
    ByVal candidateTier As Long, _
    ByVal candidateRadiusM As Double, _
    ByVal bestTier As Long, _
    ByVal bestRadiusM As Double) As Boolean

    If bestTier = ANCHOR_TIER_NONE Then
        BetterAnchor = True
        Exit Function
    End If

    If candidateTier < bestTier Then
        BetterAnchor = True
        Exit Function
    End If

    If candidateTier > bestTier Then Exit Function

    BetterAnchor = (candidateRadiusM < bestRadiusM)
End Function

Private Function AnchorTierName(ByVal tier As Long) As String
    Select Case tier
        Case ANCHOR_TIER_NATIVE_CALLOUT
            AnchorTierName = "NativeCalloutAttachment"
        Case ANCHOR_TIER_PRIMARY_DIAMETER
            AnchorTierName = "PrimaryTypedHoleDiameter"
        Case ANCHOR_TIER_SMALLEST_CIRCLE
            AnchorTierName = "SmallestCompleteCircle"
        Case Else
            AnchorTierName = "None"
    End Select
End Function

' R23-303. Route A is the model-owned forward correspondence and is the only
' route Phase 0 found to work on this build. Route B is still attempted so a
' future build that fixes it is detected rather than assumed, but nothing
' depends on it.
Private Function MapModelEntityToDrawing( _
    ByRef swView As SldWorks.View, _
    ByRef component As SldWorks.Component2, _
    ByVal modelEntity As Object, _
    ByRef mappedRoute As String) As SldWorks.Entity

    mappedRoute = "A:NotAttempted"

    Dim direct As SldWorks.Entity
    Set direct = Nothing
    On Error Resume Next
    Set direct = swView.GetCorrespondingEntity(modelEntity)
    Dim directError As Long
    directError = Err.Number
    Err.Clear
    On Error GoTo 0

    If Not direct Is Nothing Then
        mappedRoute = "A:Direct"
        Set MapModelEntityToDrawing = direct
        Exit Function
    End If

    If component Is Nothing Then
        mappedRoute = "A:Nothing:err" & CStr(directError) & _
            ";B:UnavailableNoComponent"
        Exit Function
    End If

    Dim viaComponent As SldWorks.Entity
    Set viaComponent = Nothing
    On Error Resume Next
    Set viaComponent = component.GetCorrespondingEntity(modelEntity)
    Dim componentError As Long
    componentError = Err.Number
    Err.Clear
    On Error GoTo 0

    If Not viaComponent Is Nothing Then
        mappedRoute = "B:ComponentMediated"
        Set MapModelEntityToDrawing = viaComponent
    Else
        mappedRoute = "A:Nothing:err" & CStr(directError) & _
            ";B:Nothing:err" & CStr(componentError)
    End If
End Function

' R23-502. Reuse the two acquisition routes already live-proved for a
' part drawing when an outline edge, rather than a hole edge, must become a
' selectable drawing datum. Route A is preferred; Route D is only used when
' direct correspondence declines. Route D clears its temporary selection on
' every exit and proves target-view ownership before returning an entity.
Public Function MapModelEntityToDrawingForDatum( _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByVal modelEntity As Object, _
    ByRef mappedRoute As String) As SldWorks.Entity

    On Error GoTo Failed

    Dim noComponent As SldWorks.Component2
    Dim mapped As SldWorks.Entity
    Set mapped = MapModelEntityToDrawing( _
        swView, noComponent, modelEntity, mappedRoute)

    If Not mapped Is Nothing Then
        Set MapModelEntityToDrawingForDatum = mapped
        Exit Function
    End If

    Dim selectionRoute As String
    Set mapped = SelectModelEntityInView( _
        swDraw, swView, modelEntity, selectionRoute)

    If mapped Is Nothing Then
        mappedRoute = mappedRoute & ";" & selectionRoute
        Exit Function
    End If

    mappedRoute = selectionRoute
    Set MapModelEntityToDrawingForDatum = mapped
    Exit Function

Failed:
    mappedRoute = "DatumMapError:" & CStr(Err.Number)
    Set MapModelEntityToDrawingForDatum = Nothing
End Function

' R23-502. A datum edge must be visible in the drawing, not only mappable.
' GetVisibleEntities2 requires a Component2 and supplies none for this part
' drawing. GetPolylines7 is preferred because it returns visible model
' edges. Some valid HLV part views return no polyline-edge array; only that
' documented absence falls back to a scoped IView.SelectEntity proof.
Public Function MapVisibleDatumEntity( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByVal modelEntity As Object, _
    ByRef mappedRoute As String) As SldWorks.Entity

    On Error GoTo Failed

    mappedRoute = "DatumMap:NotAttempted"
    If swApp Is Nothing Then
        mappedRoute = "DatumMap:UnavailableNoApplication"
        Exit Function
    End If
    If swDraw Is Nothing Then
        mappedRoute = "DatumMap:UnavailableNoDrawing"
        Exit Function
    End If
    If swView Is Nothing Then
        mappedRoute = "DatumMap:UnavailableNoView"
        Exit Function
    End If
    If modelEntity Is Nothing Then
        mappedRoute = "DatumMap:UnavailableNoModelEntity"
        Exit Function
    End If

    ' r42 settled what this array actually holds. The same comparison, run
    ' twice against the same GetPolylines7 result, matched the DRAWING entity
    ' 12 times out of 12 and the MODEL edge 0 times out of 12, on edges
    ' GetVisibleEntities2 had independently confirmed present. The 2025 Help
    ' calls the return "modeling edges", but on this build it compares equal
    ' to the drawing entity Route A returns, not to the model edge it came
    ' from. Testing the model entity - which is what this function did from
    ' r33 until now - could therefore never match, so every datum failed
    ' closed with PolylineVisibilityUnavailable no matter how visible its
    ' edge was. The order below is corrected: map first, then prove the
    ' MAPPED entity is in the visible array.
    '
    ' The r37 fail-closed guard is deliberately unchanged. A non-empty array
    ' that does not contain the mapped entity still refuses; only the
    ' documented empty-array case falls back to scoped selection.
    Dim noComponent As SldWorks.Component2
    Dim mapped As SldWorks.Entity
    Set mapped = MapModelEntityToDrawing( _
        swView, noComponent, modelEntity, mappedRoute)

    Dim visibilityProof As String
    Dim visibleIndex As Long

    If Not mapped Is Nothing Then
        visibleIndex = FindVisibleModelEdgeIndex( _
            swApp, swView, mapped, visibilityProof)

        If visibleIndex >= 0 Then
            mappedRoute = mappedRoute & "|visibility=GetPolylines7" & _
                "|space=Drawing" & _
                "|visibleEntityIndex=" & CStr(visibleIndex) & _
                "|" & visibilityProof
            Set MapVisibleDatumEntity = mapped
            Exit Function
        End If

        ' A mapped entity absent from a non-empty visible array is the case
        ' the guard exists for: the edge maps but is not shown. Refuse.
        If InStr(1, visibilityProof, "status=NoEdges", _
            vbBinaryCompare) = 0 Then

            mappedRoute = "DatumMap:MappedEntityNotVisible|" & _
                mappedRoute & "|" & visibilityProof
            Exit Function
        End If
    Else
        ' Nothing to test in drawing space yet. Read the array once so the
        ' empty-array fallback condition can still be evaluated.
        visibleIndex = FindVisibleModelEdgeIndex( _
            swApp, swView, modelEntity, visibilityProof)

        If InStr(1, visibilityProof, "status=NoEdges", _
            vbBinaryCompare) = 0 Then

            mappedRoute = "DatumMap:UnmappedWithVisibleArray|" & _
                mappedRoute & "|" & visibilityProof
            Exit Function
        End If
    End If

    ' Documented empty-array case only. IView.SelectEntity accepts the model
    ' edge only in the target drawing view; its returned drawing entity is
    ' then checked with GetSelectedObjectsDrawingView2 before it is retained.
    Dim selectionRoute As String
    Set mapped = SelectModelEntityInView( _
        swDraw, swView, modelEntity, selectionRoute)
    If mapped Is Nothing Then
        mappedRoute = "DatumMap:ScopedSelectionUnavailable|" & _
            visibilityProof & ";" & selectionRoute
        Exit Function
    End If

    mappedRoute = selectionRoute & _
        "|visibility=ScopedViewSelection" & _
        "|polyline=" & visibilityProof
    Set MapVisibleDatumEntity = mapped
    Exit Function

Failed:
    mappedRoute = "DatumMap:Error:" & CStr(Err.Number)
    Set MapVisibleDatumEntity = Nothing
End Function

' R23-502. GetPolylines7 returns model edges for all visible solid-model
' edges in a view. Unlike GetVisibleEntities2, it works without a drawing
' Component2, which this installed part-drawing context does not expose.
Private Function FindVisibleModelEdgeIndex( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef modelEntity As Object, _
    ByRef visibilityProof As String) As Long

    FindVisibleModelEdgeIndex = -1
    visibilityProof = "source=IView.GetPolylines7|status=NotAttempted"
    If swApp Is Nothing Then
        visibilityProof = "source=IView.GetPolylines7|status=NoApplication"
        Exit Function
    End If
    If swView Is Nothing Then
        visibilityProof = "source=IView.GetPolylines7|status=NoView"
        Exit Function
    End If
    If modelEntity Is Nothing Then
        visibilityProof = "source=IView.GetPolylines7|status=NoModelEntity"
        Exit Function
    End If

    Dim polylineData As Variant
    Dim visibleModelEdges As Variant
    Dim readError As Long
    On Error Resume Next
    visibleModelEdges = swView.GetPolylines7( _
        CROSSHATCH_EXCLUDE, polylineData)
    readError = Err.Number
    Err.Clear
    On Error GoTo 0

    Dim edgeCount As Long
    edgeCount = VariantItemCount(visibleModelEdges)
    If edgeCount = 0 Then
        visibilityProof = "source=IView.GetPolylines7|status=NoEdges" & _
            "|error=" & CStr(readError)
        Exit Function
    End If

    Dim unsupportedComparisons As Long
    Dim visibleIndex As Long
    visibleIndex = FindVisibleEntityIndex( _
        swApp, visibleModelEdges, modelEntity, unsupportedComparisons)
    visibilityProof = "source=IView.GetPolylines7|modelEdges=" & _
        CStr(edgeCount) & "|unsupported=" & _
        CStr(unsupportedComparisons) & "|error=" & CStr(readError)

    If visibleIndex < 0 Then Exit Function

    FindVisibleModelEdgeIndex = visibleIndex
End Function

' Route D starts with a model edge but returns an entity from the drawing
' selection list. It is intentionally narrower than a coordinate search:
' IView.SelectEntity scopes the selection and the selection manager proves
' target-view ownership before its object is used as an anchor.
Private Function SelectModelEntityInView( _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByVal modelEntity As Object, _
    ByRef selectionRoute As String) As SldWorks.Entity

    On Error GoTo Failed

    selectionRoute = "D:NotAttempted"
    If swDraw Is Nothing Then
        selectionRoute = "D:UnavailableNoDrawing"
        Exit Function
    End If
    If swView Is Nothing Then
        selectionRoute = "D:UnavailableNoView"
        Exit Function
    End If
    If modelEntity Is Nothing Then
        selectionRoute = "D:UnavailableNoModelEntity"
        Exit Function
    End If

    Dim selectionMgr As SldWorks.SelectionMgr
    Set selectionMgr = swDraw.SelectionManager
    If selectionMgr Is Nothing Then
        selectionRoute = "D:UnavailableNoSelectionManager"
        Exit Function
    End If

    ' IView.SelectEntity with AppendFlag=False replaces the current
    ' selection. Do not erase an interactive caller's state: refuse before
    ' selecting when anything is already selected. R23 probes start clean;
    ' a non-empty selection is a named precondition failure, not cleanup.
    Dim initialSelectionCount As Long
    initialSelectionCount = selectionMgr.GetSelectedObjectCount2(-1)
    If initialSelectionCount <> 0 Then
        selectionRoute = "D:RefusedPreexistingSelection:count" & _
            CStr(initialSelectionCount)
        Exit Function
    End If

    Dim selected As Boolean
    Dim selectError As Long
    On Error Resume Next
    selected = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swView.SelectEntity(modelEntity, False))
    selectError = Err.Number
    Err.Clear
    On Error GoTo Failed

    Dim selectedCount As Long
    selectedCount = selectionMgr.GetSelectedObjectCount2(-1)
    If Not selected Or selectedCount <> 1 Then
        selectionRoute = "D:SelectEntityFailed:result" & CStr(selected) & _
            ":count" & CStr(selectedCount) & ":err" & CStr(selectError)
        GoTo CleanUp
    End If

    Dim owningView As SldWorks.View
    Set owningView = selectionMgr.GetSelectedObjectsDrawingView2(1, -1)
    If owningView Is Nothing Then
        selectionRoute = "D:OwnerUnavailable"
        GoTo CleanUp
    End If

    If StrComp(Module11_GeometryIdentity.IdentityToken( _
        SafeViewName(owningView)), Module11_GeometryIdentity.IdentityToken( _
        SafeViewName(swView)), vbBinaryCompare) <> 0 Then

        selectionRoute = "D:WrongOwner:" & SafeViewName(owningView)
        GoTo CleanUp
    End If

    Dim selectedObject As Object
    Set selectedObject = selectionMgr.GetSelectedObject6(1, -1)
    If selectedObject Is Nothing Then
        selectionRoute = "D:SelectedObjectUnavailable"
        GoTo CleanUp
    End If

    Dim selectedEntity As SldWorks.Entity
    Set selectedEntity = selectedObject
    If selectedEntity Is Nothing Then
        selectionRoute = "D:SelectedObjectNotEntity"
        GoTo CleanUp
    End If

    selectionRoute = "D:ViewSelectEntity|owner=" & SafeViewName(owningView)
    Set SelectModelEntityInView = selectedEntity

CleanUp:
    swDraw.ClearSelection2 True
    Exit Function

Failed:
    selectionRoute = "D:Error:" & CStr(Err.Number)
    On Error Resume Next
    swDraw.ClearSelection2 True
    Set SelectModelEntityInView = Nothing
End Function

' R23-304. Identity comparison, not geometric similarity. ISldWorks.IsSame is
' the only test that proves two dispatch pointers denote the same entity.
Private Function FindVisibleEntityIndex( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal visibleEntities As Variant, _
    ByRef candidate As Object, _
    ByRef unsupportedComparisons As Long) As Long

    FindVisibleEntityIndex = -1
    unsupportedComparisons = 0

    If candidate Is Nothing Then Exit Function
    If IsEmpty(visibleEntities) Then Exit Function
    If Not IsArray(visibleEntities) Then Exit Function

    On Error GoTo Failed

    Dim i As Long
    For i = LBound(visibleEntities) To UBound(visibleEntities)
        Dim entry As Object
        Set entry = Nothing
        On Error Resume Next
        Set entry = visibleEntities(i)
        On Error GoTo Failed

        If Not entry Is Nothing Then
            Dim equality As Long
            equality = SafeObjectEquality(swApp, entry, candidate)

            If equality = OBJECT_EQUALITY_SAME Then
                FindVisibleEntityIndex = i
                Exit Function
            End If

            If equality = OBJECT_EQUALITY_UNSUPPORTED Then
                unsupportedComparisons = unsupportedComparisons + 1
            End If
        End If
    Next i
    Exit Function

Failed:
    FindVisibleEntityIndex = -1
End Function

' Returns swObjectEquality, defaulting to swObjectNotSame. An unreadable
' comparison must never present as identity.
Private Function SafeObjectEquality( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef first As Object, _
    ByRef second As Object) As Long

    On Error GoTo Failed
    SafeObjectEquality = CLng(swApp.IsSame(first, second))
    Exit Function
Failed:
    SafeObjectEquality = 0
End Function

' R23-305. A complete circle, proved the way Phase 0 proved it: ICurve.IsCircle
' through the shared Boolean normalizer, plus a GetCurveParams3 parameter
' range whose endpoints coincide. The raw Boolean is never negated directly.
Private Function EdgeIsCompleteCircle( _
    ByRef modelEdge As SldWorks.Edge, _
    ByRef radiusM As Double, _
    ByRef circleProof As String) As Boolean

    On Error GoTo Failed

    circleProof = "circle=Unproven"
    radiusM = 0#

    Dim curve As SldWorks.Curve
    Set curve = Nothing
    On Error Resume Next
    Set curve = modelEdge.GetCurve
    On Error GoTo Failed

    If curve Is Nothing Then
        circleProof = "circle=Reject|reason=CurveUnavailable"
        Exit Function
    End If

    Dim rawIsCircle As Variant
    rawIsCircle = curve.IsCircle

    If Not Module11_GeometryIdentity.NormalizeSwBoolean(rawIsCircle) Then
        circleProof = "circle=Reject|reason=NotCircle" & _
            "|isCircleRaw=" & CStr(rawIsCircle)
        Exit Function
    End If

    Dim circleData As Variant
    circleData = curve.CircleParams

    If VariantItemCount(circleData) <> 7 Then
        circleProof = "circle=Reject|reason=CircleParamsItemCount:" & _
            CStr(VariantItemCount(circleData))
        Exit Function
    End If

    radiusM = Abs(CDbl(circleData(LBound(circleData) + 6)))

    ' IEdge.GetCurveParams3 returns an ICurveParamData OBJECT, not an array
    ' of doubles. A Let assignment into a Variant asks it for a default
    ' property it does not have and raises error 438, which is what rejected
    ' every edge on the first instrumented run. GetCurve must already have
    ' been called, which it has above.
    Dim curveParams As SldWorks.CurveParamData
    Set curveParams = Nothing
    Set curveParams = modelEdge.GetCurveParams3

    If curveParams Is Nothing Then
        circleProof = "circle=Reject|reason=CurveParamsNothing"
        Exit Function
    End If

    Dim startPoint As Variant
    Dim endPoint As Variant
    startPoint = curveParams.StartPoint
    endPoint = curveParams.EndPoint

    If VariantItemCount(startPoint) < 3 Or _
       VariantItemCount(endPoint) < 3 Then

        circleProof = "circle=Reject|reason=CurveParamsEndpointsUnavailable"
        Exit Function
    End If

    Dim closureM As Double
    closureM = PointDistance(startPoint, endPoint)

    ' PointDistance reports failure as a negative value, which would slip
    ' under the closure tolerance and read as a perfectly closed curve.
    If closureM < 0# Then
        circleProof = "circle=Reject|reason=ClosureNotMeasurable"
        Exit Function
    End If

    If closureM > CIRCLE_CLOSURE_TOLERANCE_M Then
        circleProof = "circle=Reject|reason=CurveNotClosed" & _
            "|uMin=" & Format$(curveParams.UMinValue, "0.000000000") & _
            "|uMax=" & Format$(curveParams.UMaxValue, "0.000000000") & _
            "|closureM=" & Format$(closureM, "0.000000000")
        Exit Function
    End If

    circleProof = "circle=Proven" & _
        "|isCircleRaw=" & CStr(rawIsCircle) & _
        "|uMin=" & Format$(curveParams.UMinValue, "0.000000000") & _
        "|uMax=" & Format$(curveParams.UMaxValue, "0.000000000") & _
        "|closureM=" & Format$(closureM, "0.000000000") & _
        "|proof=IsCircleAndClosedCurveParams"
    EdgeIsCompleteCircle = True
    Exit Function

Failed:
    circleProof = "circle=Reject|reason=ReadError:" & CStr(Err.Number)
    EdgeIsCompleteCircle = False
End Function

' R23-306. The hole projects as a circle only when its axis is parallel to
' the view normal. Both senses are accepted because the axis is already
' sign-normalized and a hole seen from either side still projects circular.
Private Function AxisIsNormalToView( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef location As CPhysicalHoleLocation, _
    ByRef projectedAxisX As Double, _
    ByRef projectedAxisY As Double, _
    ByRef projectedAxisZ As Double) As Boolean

    On Error GoTo Failed

    Dim originX As Double
    Dim originY As Double
    Dim originZ As Double
    Dim tipX As Double
    Dim tipY As Double
    Dim tipZ As Double
    Dim originProof As String
    Dim tipProof As String

    ' The transform is applied to two model points and differenced, so the
    ' result is the axis direction expressed in the view frame. Transforming
    ' a direction as if it were a point would fold in the translation.
    If Not Module8_RuntimeSupport.TransformPointToView( _
        swApp, swView, 0#, 0#, 0#, _
        originX, originY, originZ, originProof, False) Then

        Exit Function
    End If

    If Not Module8_RuntimeSupport.TransformPointToView( _
        swApp, swView, location.AxisX, location.AxisY, location.AxisZ, _
        tipX, tipY, tipZ, tipProof, False) Then

        Exit Function
    End If

    projectedAxisX = tipX - originX
    projectedAxisY = tipY - originY
    projectedAxisZ = tipZ - originZ

    ' A view-frame direction whose in-plane components vanish is parallel to
    ' the view normal.
    Dim inPlane As Double
    inPlane = Sqr(projectedAxisX * projectedAxisX + _
        projectedAxisY * projectedAxisY)

    AxisIsNormalToView = (inPlane <= AXIS_NORMAL_TOLERANCE)
    Exit Function

Failed:
    AxisIsNormalToView = False
End Function

' Page-frame centre of the cylindrical face, through the already-proven
' production transform. The frame is stated in the proof string because the
' Phase 0 section-line payload showed that mixing view-sketch and page values
' silently produces false results.
Private Function ProjectFaceCentreToPage( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef modelFace As SldWorks.Face2, _
    ByRef pageX As Double, _
    ByRef pageY As Double, _
    ByRef pageZ As Double, _
    ByRef frameProof As String) As Boolean

    On Error GoTo Failed

    frameProof = "Unproven"

    Dim surface As SldWorks.Surface
    Set surface = Nothing
    On Error Resume Next
    Set surface = modelFace.GetSurface
    On Error GoTo Failed

    If surface Is Nothing Then Exit Function

    Dim rawIsCylinder As Variant
    rawIsCylinder = surface.IsCylinder

    If Not Module11_GeometryIdentity.NormalizeSwBoolean(rawIsCylinder) Then
        Exit Function
    End If

    Dim cylinderData As Variant
    cylinderData = surface.CylinderParams

    If VariantItemCount(cylinderData) <> 7 Then Exit Function

    Dim baseIndex As Long
    baseIndex = LBound(cylinderData)

    Dim transformProof As String
    If Not Module8_RuntimeSupport.TransformPointToView( _
        swApp, swView, _
        CDbl(cylinderData(baseIndex)), _
        CDbl(cylinderData(baseIndex + 1)), _
        CDbl(cylinderData(baseIndex + 2)), _
        pageX, pageY, pageZ, transformProof, True) Then

        frameProof = "Unproven"
        Exit Function
    End If

    frameProof = "Page|source=Module8_RuntimeSupport.TransformPointToView" & _
        "|origin=ISurface.CylinderParams|" & transformProof
    ProjectFaceCentreToPage = True
    Exit Function

Failed:
    frameProof = "Unproven"
    ProjectFaceCentreToPage = False
End Function

' R23-307. Selects the projection's anchor only from an empty selection.
' Only drawing-document entities are ever selected. A caller's non-empty
' selection is refused intact because a proof pass cannot restore marks,
' points, or selection order exactly.
' ISelectData.View is documented get/set but raises runtime error 91 on this
' build, which Module2_DrawingPipeline.CreatePrimarySection already records.
' The binding is therefore attempted in a guarded helper and, bound or not,
' ownership is proved AFTER the fact through
' ISelectionMgr.GetSelectedObjectsDrawingView2 rather than assumed.
Public Function SelectAnchorInView( _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef projection As CViewHoleProjection, _
    ByRef selectionProof As String) As Boolean

    On Error GoTo Failed

    selectionProof = "selection=NotAttempted"

    If swDraw Is Nothing Then Exit Function
    If projection Is Nothing Then Exit Function
    If Not projection.HasSelectableAnchor() Then
        selectionProof = "selection=Reject|reason=NoAnchor"
        Exit Function
    End If

    Dim selectionMgr As SldWorks.SelectionMgr
    Set selectionMgr = swDraw.SelectionManager
    If selectionMgr Is Nothing Then
        selectionProof = "selection=Reject|reason=SelectionManagerUnavailable"
        Exit Function
    End If

    Dim initialSelectionCount As Long
    initialSelectionCount = selectionMgr.GetSelectedObjectCount2(-1)
    If initialSelectionCount <> 0 Then
        selectionProof = "selection=Reject|reason=PreexistingSelection" & _
            "|initialSelectionCount=" & CStr(initialSelectionCount)
        Exit Function
    End If

    Dim selectData As SldWorks.SelectData
    Set selectData = selectionMgr.CreateSelectData

    Dim viewBinding As String
    viewBinding = TryBindSelectDataView(selectData, projection.DrawingView)

    Dim selected As Boolean
    selected = Module11_GeometryIdentity.NormalizeSwBoolean( _
        projection.PrimaryAnchor.Select4(False, selectData))

    Dim selectedCount As Long
    selectedCount = selectionMgr.GetSelectedObjectCount2(-1)

    Dim owningViewName As String
    owningViewName = "Unproven"

    If selected And selectedCount = 1 Then
        Dim owningView As SldWorks.View
        Set owningView = Nothing
        On Error Resume Next
        Set owningView = selectionMgr.GetSelectedObjectsDrawingView2(1, -1)
        On Error GoTo Failed

        If Not owningView Is Nothing Then
            owningViewName = SafeViewName(owningView)
        End If
    End If

    Dim ownershipProven As Boolean
    ownershipProven = (StrComp( _
        Module11_GeometryIdentity.IdentityToken(owningViewName), _
        Module11_GeometryIdentity.IdentityToken(projection.ViewName), _
        vbBinaryCompare) = 0)

    selectionProof = "selection=" & CStr(selected) & _
        "|viewBinding=" & viewBinding & _
        "|selectedCount=" & CStr(selectedCount) & _
        "|owningView=" & owningViewName & _
        "|ownershipProven=" & CStr(ownershipProven) & _
        "|ownershipSource=ISelectionMgr.GetSelectedObjectsDrawingView2"

    ' This is a proof pass, not an editing pass: leave nothing selected.
    swDraw.ClearSelection2 True

    SelectAnchorInView = (selected And selectedCount = 1 And ownershipProven)
    Exit Function

Failed:
    selectionProof = "selection=Reject|reason=Error:" & CStr(Err.Number)
    On Error Resume Next
    swDraw.ClearSelection2 True
    SelectAnchorInView = False
End Function

Private Function TryBindSelectDataView( _
    ByRef selectData As SldWorks.SelectData, _
    ByRef swView As SldWorks.View) As String

    On Error GoTo Failed
    Set selectData.View = swView
    TryBindSelectDataView = "Bound"
    Exit Function
Failed:
    TryBindSelectDataView = "UnboundAfterError:" & CStr(Err.Number)
End Function

Private Function ReadReferencedConfiguration( _
    ByRef swView As SldWorks.View, _
    ByRef referencedConfiguration As String) As Boolean

    On Error GoTo Failed

    referencedConfiguration = vbNullString
    referencedConfiguration = swView.ReferencedConfiguration

    ReadReferencedConfiguration = _
        (Len(Trim$(referencedConfiguration)) > 0)
    Exit Function

Failed:
    referencedConfiguration = vbNullString
    ReadReferencedConfiguration = False
End Function

' R23-310. Reports what is missing rather than returning a bare count, so a
' partial result cannot be mistaken for success.
Public Function VerifyExpectedProjections( _
    ByRef graph As CLocationGraph, _
    ByVal viewName As String, _
    ByVal expectedAcceptedCount As Long) As String

    Dim failures As String

    If graph Is Nothing Then
        VerifyExpectedProjections = "GraphMissing"
        Exit Function
    End If

    Dim projections As Collection
    Set projections = graph.ProjectionsForView(viewName)

    Dim accepted As Long
    Dim unanchored As Long
    Dim axisNotNormal As Long
    Dim frameUnproven As Long
    Dim configurationUnproven As Long

    Dim i As Long
    For i = 1 To projections.Count
        Dim projection As CViewHoleProjection
        Set projection = projections(i)

        If projection.Accepted Then
            accepted = accepted + 1
        Else
            Select Case projection.RejectionReason
                Case "ProjectionAnchorUnavailable"
                    unanchored = unanchored + 1
                Case "AxisNotNormalToView"
                    axisNotNormal = axisNotNormal + 1
                Case "CoordinateFrameUnproven"
                    frameUnproven = frameUnproven + 1
                Case "ReferencedConfigurationUnproven"
                    configurationUnproven = configurationUnproven + 1
            End Select
        End If
    Next i

    If accepted <> expectedAcceptedCount Then
        failures = AppendFailure(failures, _
            "AcceptedProjectionCount:" & CStr(accepted) & _
            "Expected:" & CStr(expectedAcceptedCount))
    End If

    ' Not a defect on its own: a view legitimately sees only the holes whose
    ' axes are normal to it. Reported so the count is visible rather than
    ' silently folded into the unanchored total.
    If axisNotNormal > 0 Then
        failures = AppendFailure(failures, _
            "AxisNotNormalToView:" & CStr(axisNotNormal))
    End If

    If unanchored > 0 Then
        failures = AppendFailure(failures, _
            "ProjectionAnchorUnavailable:" & CStr(unanchored))
    End If

    If frameUnproven > 0 Then
        failures = AppendFailure(failures, _
            "CoordinateFrameUnproven:" & CStr(frameUnproven))
    End If

    If configurationUnproven > 0 Then
        failures = AppendFailure(failures, _
            "ReferencedConfigurationUnproven:" & CStr(configurationUnproven))
    End If

    VerifyExpectedProjections = failures
End Function

' Per-view acceptance tally for the R23-310 report. A single global count
' cannot express the requirement, which is about which holes are usable in
' WHICH view.
Public Function ViewAcceptanceSummary( _
    ByRef graph As CLocationGraph, _
    ByVal viewName As String) As String

    If graph Is Nothing Then
        ViewAcceptanceSummary = "GraphMissing"
        Exit Function
    End If

    Dim projections As Collection
    Set projections = graph.ProjectionsForView(viewName)

    Dim accepted As Long
    Dim axisNormal As Long
    Dim anchored As Long
    Dim coincident As Long

    Dim i As Long
    For i = 1 To projections.Count
        Dim projection As CViewHoleProjection
        Set projection = projections(i)

        If projection.Accepted Then accepted = accepted + 1
        If projection.AxisNormalToView Then axisNormal = axisNormal + 1
        If projection.HasSelectableAnchor() Then anchored = anchored + 1

        If Len(projection.CoincidentWithAnchoredKey) > 0 Then
            coincident = coincident + 1
        End If
    Next i

    ViewAcceptanceSummary = "view=" & viewName & _
        "|projections=" & CStr(projections.Count) & _
        "|axisNormal=" & CStr(axisNormal) & _
        "|anchored=" & CStr(anchored) & _
        "|coincidentUnanchored=" & CStr(coincident) & _
        "|accepted=" & CStr(accepted)
End Function

Private Function AppendFailure( _
    ByVal current As String, _
    ByVal reasonCode As String) As String

    If Len(current) = 0 Then
        AppendFailure = reasonCode
    Else
        AppendFailure = current & ";" & reasonCode
    End If
End Function

' Euclidean distance between two 3-double coordinate arrays, used to prove a
' curve closes on itself.
Private Function PointDistance( _
    ByVal firstPoint As Variant, _
    ByVal secondPoint As Variant) As Double

    On Error GoTo Failed

    PointDistance = -1#

    Dim firstBase As Long
    Dim secondBase As Long
    firstBase = LBound(firstPoint)
    secondBase = LBound(secondPoint)

    Dim deltaX As Double
    Dim deltaY As Double
    Dim deltaZ As Double
    deltaX = CDbl(firstPoint(firstBase)) - CDbl(secondPoint(secondBase))
    deltaY = CDbl(firstPoint(firstBase + 1)) - _
        CDbl(secondPoint(secondBase + 1))
    deltaZ = CDbl(firstPoint(firstBase + 2)) - _
        CDbl(secondPoint(secondBase + 2))

    PointDistance = Sqr(deltaX * deltaX + deltaY * deltaY + deltaZ * deltaZ)
    Exit Function

Failed:
    PointDistance = -1#
End Function

Private Function VariantItemCount(ByVal value As Variant) As Long
    On Error GoTo Failed

    If IsEmpty(value) Then Exit Function
    If Not IsArray(value) Then Exit Function

    VariantItemCount = UBound(value) - LBound(value) + 1
    Exit Function

Failed:
    VariantItemCount = 0
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

' swDrawingViewTypes_e. Logged rather than used as a filter: which code the
' sheet placeholders carry is not established on this build, and the visible
' entity count is a fact rather than an assumption.
Private Function SafeViewType( _
    ByRef swView As SldWorks.View) As Long

    On Error Resume Next
    If swView Is Nothing Then Exit Function
    SafeViewType = -1
    SafeViewType = swView.Type
End Function

' swDisplayMode_e. This decides whether a hole on the far side of the part
' exists as a drawing entity at all: under Hidden Lines Removed it is never
' drawn, so IView.GetCorrespondingEntity has nothing to return and the
' location cannot anchor in that view no matter how the search is written.
' Recorded so an unanchorable location can be attributed to the drawing's
' display setting rather than mistaken for a mapping defect.
Private Function SafeDisplayMode( _
    ByRef swView As SldWorks.View) As Long

    On Error Resume Next
    SafeDisplayMode = -1
    If swView Is Nothing Then Exit Function
    SafeDisplayMode = swView.GetDisplayMode2
End Function

Private Function DisplayModeName(ByVal displayMode As Long) As String
    Select Case displayMode
        Case 0
            DisplayModeName = "Wireframe"
        Case 1
            DisplayModeName = "HiddenLinesVisible"
        Case 2
            DisplayModeName = "HiddenLinesRemoved"
        Case 3
            DisplayModeName = "Shaded"
        Case 4
            DisplayModeName = "FacetedWireframe"
        Case 5
            DisplayModeName = "FacetedHiddenGreyed"
        Case 6
            DisplayModeName = "FacetedHidden"
        Case 7
            DisplayModeName = "ShadedWithEdges"
        Case 8
            DisplayModeName = "Default"
        Case Else
            DisplayModeName = "Unknown:" & CStr(displayMode)
    End Select
End Function

Private Function SafeComponentName( _
    ByRef component As SldWorks.Component2) As String

    On Error Resume Next
    If component Is Nothing Then
        SafeComponentName = vbNullString
        Exit Function
    End If

    SafeComponentName = vbNullString
    SafeComponentName = component.Name2
End Function

' Read-only evidence entry point for R23-310. Requires the authorized P-0251
' fixture and an open drawing. Creates nothing, selects nothing on exit, and
' never saves.
Public Sub R23_ProbeViewProjections()
    On Error GoTo Failed

    mEmitDiagnostics = False

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_PROJECTION_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swDraw As SldWorks.ModelDoc2
    Set swDraw = swApp.ActiveDoc

    If swDraw Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_PROJECTION_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swDraw.GetType <> swDocDRAWING Then
        Module21_EvidenceSink.LogLine _
            "R23_PROJECTION_FATAL|reason=ActiveDocumentNotDrawing"
        Exit Sub
    End If

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet

    Dim views As Variant
    views = swSheet.GetViews

    If IsEmpty(views) Or Not IsArray(views) Then
        Module21_EvidenceSink.LogLine _
            "R23_PROJECTION_FATAL|reason=NoViewsOnSheet"
        Exit Sub
    End If

    ' The model is read through the first view that references one, so the
    ' probe never guesses which document the drawing is of.
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
        Module21_EvidenceSink.LogLine _
            "R23_PROJECTION_FATAL|reason=NoReferencedDocument"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Module21_EvidenceSink.LogLine _
            "R23_PROJECTION_FATAL|reason=UnauthorizedFixture" & _
            "|path=" & partPath
        Exit Sub
    End If

    Dim drawingSaveFlagBefore As Boolean
    drawingSaveFlagBefore = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Dim partSaveFlagBefore As Boolean
    partSaveFlagBefore = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swPart.GetSaveFlag)

    ' The operator may already have something selected. A raw final count is
    ' therefore not evidence of cleanliness: the first run reported
    ' finalSelectionCount=1 with no accepted projection and nothing selected
    ' by this code. The baseline is captured so the report can state whether
    ' THIS pass left anything behind.
    Dim initialSelectionCount As Long
    initialSelectionCount = swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim evidence As CRunEvidence
    Set evidence = New CRunEvidence

    Dim graph As CLocationGraph
    Set graph = New CLocationGraph

    Module21_EvidenceSink.LogLine _
        "R23_PROJECTION_BEGIN|drawing=" & swDraw.GetPathName & _
        "|initialSelectionCount=" & CStr(initialSelectionCount) & _
        "|part=" & partPath & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath) & _
        "|views=" & CStr(VariantItemCount(views))

    Dim configurationName As String
    configurationName = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    If Not Module12_FeatureQualification.BuildFeatureCatalog( _
        swApp, swPart, configurationName, graph, evidence) Then

        Module21_EvidenceSink.LogLine _
            "R23_PROJECTION_FATAL|reason=CatalogUnavailable"
        Exit Sub
    End If

    Module21_EvidenceSink.LogLine _
        "R23_PROJECTION_CATALOG|" & graph.GraphSummary()

    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueView

        BuildViewProjections swApp, swDraw, swView, graph, evidence

ContinueView:
    Next i

    ' Per-view tally. R23-310 is a statement about which holes are usable in
    ' which view, so a single total cannot answer it.
    For i = LBound(views) To UBound(views)
        Dim summaryView As SldWorks.View
        Set summaryView = views(i)
        If summaryView Is Nothing Then GoTo ContinueSummary

        Dim summaryName As String
        summaryName = SafeViewName(summaryView)

        If graph.ProjectionsForView(summaryName).Count > 0 Then
            Module21_EvidenceSink.LogLine _
                "R23_PROJECTION_VIEW_SUMMARY|" & _
                ViewAcceptanceSummary(graph, summaryName)
        End If

ContinueSummary:
    Next i

    ' Selection is proved on accepted projections only, one at a time, with
    ' the selection cleared again before the next.
    Dim projections As Collection
    Set projections = graph.Projections()

    Dim selectionProved As Long
    Dim selectionAttempted As Long
    For i = 1 To projections.Count
        Dim projection As CViewHoleProjection
        Set projection = projections(i)

        If projection.Accepted Then
            selectionAttempted = selectionAttempted + 1
            Dim selectionProof As String
            If SelectAnchorInView(swDraw, projection, selectionProof) Then
                selectionProved = selectionProved + 1
            End If

            Module21_EvidenceSink.LogLine _
                "R23_PROJECTION_SELECT|view=" & _
                projection.ViewName & _
                "|physical=" & projection.PhysicalInstanceKey & _
                "|" & selectionProof
        End If
    Next i

    Dim drawingSaveFlagAfter As Boolean
    drawingSaveFlagAfter = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Dim partSaveFlagAfter As Boolean
    partSaveFlagAfter = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swPart.GetSaveFlag)

    Dim finalSelectionCount As Long
    finalSelectionCount = swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    ' Selection is only this pass's responsibility if this pass selected
    ' anything. When no anchor was proved, the operator's own pre-existing
    ' selection is untouched and a non-zero count says nothing about us.
    Dim selectionClean As String
    If selectionAttempted = 0 Then
        selectionClean = "NotAttempted|operatorSelectionPreserved=" & _
            CStr(finalSelectionCount = initialSelectionCount)
    Else
        selectionClean = CStr(finalSelectionCount = 0)
    End If

    Module21_EvidenceSink.LogLine _
        "R23_PROJECTION_END|" & graph.GraphSummary() & _
        "|selectionAttempted=" & CStr(selectionAttempted) & _
        "|selectionProved=" & CStr(selectionProved) & _
        "|initialSelectionCount=" & CStr(initialSelectionCount) & _
        "|finalSelectionCount=" & CStr(finalSelectionCount) & _
        "|selectionClean=" & selectionClean & _
        "|drawingUnchanged=" & _
            CStr(drawingSaveFlagBefore = drawingSaveFlagAfter) & _
        "|partUnchanged=" & CStr(partSaveFlagBefore = partSaveFlagAfter)
    Exit Sub

Failed:
    Module21_EvidenceSink.LogLine _
        "R23_PROJECTION_FATAL|reason=UnhandledError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description
End Sub

' R23 probe-runner compile-failure localisation. A no-op; VBA compiles
' at module granularity, so a module that loads this has compiled.
Public Sub R23_CompileTouch()
End Sub
