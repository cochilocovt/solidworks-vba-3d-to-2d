Option Explicit

' R23 Phase 4. Model-item import, annotation classification, and
' reconciliation against the location graph.
' SAFETY BOUNDARY. This module contains the only R23 code that can modify a
' drawing. Two procedures mutate: ImportModelAnnotations inserts annotations,
' and RemoveR23CreatedAnnotations deletes them. Both require an explicit
' allowMutation argument that defaults to nothing and is never passed by the
' evidence entry point. R23_ProbeAnnotationReconciliation is strictly
' read-only: it inventories, classifies and reconciles the annotations a
' drawing ALREADY has, so it can be run against a manual reference drawing
' without altering it.
' That read-only path is not a reduced form of the work. The reference
' drawing already carries the manufacturing intent R23 must reproduce, so
' reconciling against it exercises R23-406 through R23-409, R23-411 and
' R23-412 on real data.

' R23-400. Import constants, verified member by member against the 2025
' swInsertAnnotation_e table. The Phase 0 mask 18055274 decomposes with no
' unaccounted bit:
' 2         swInsertDatums
' 8         swInsertDimensions
' 32        swInsertGTols
' 64        swInsertNotes
' 32768     swInsertDimensionsMarkedForDrawing
' 65536     swInsertHoleWizardProfileDimensions
' 131072    swInsertHoleWizardLocationDimensions
' 1048576   swInsertholeCallout
' 16777216  swInsertTolerancedDims
' The callout member really is spelled with a lowercase h in the enum.
Private Const INSERT_DATUMS As Long = 2
Private Const INSERT_DIMENSIONS As Long = 8
Private Const INSERT_GTOLS As Long = 32
Private Const INSERT_NOTES As Long = 64
Private Const INSERT_DIMS_MARKED As Long = 32768
Private Const INSERT_HW_PROFILE As Long = 65536
Private Const INSERT_HW_LOCATION As Long = 131072
Private Const INSERT_HOLE_CALLOUT As Long = 1048576
Private Const INSERT_TOLERANCED As Long = 16777216

' The expanded mask proved in Phase 0 to return 25 unique identities.
Public Const IMPORT_MASK_FULL As Long = 18055274

' Section-first mask: general, marked and toleranced dimensions, datums and
' GTols. Deliberately excludes hole-wizard and callout bits, which Phase 0
' showed belong in the side and primary passes.
Public Const IMPORT_MASK_SECTION As Long = 16810026

' Source scope. 0 = entire model.
Private Const IMPORT_FROM_ENTIRE_MODEL As Long = 0

' swAnnotationType_e.
Private Const ANN_CTHREAD As Long = 1
Private Const ANN_DATUM_TAG As Long = 2
Private Const ANN_DISPLAY_DIMENSION As Long = 4
Private Const ANN_GTOL As Long = 5
Private Const ANN_NOTE As Long = 6
Private Const ANN_DATUM_ORIGIN As Long = 16

Private Const ANN_CENTER_MARK As Long = 13
Private Const ANN_CENTER_LINE As Long = 15

' swDimensionType_e as observed live. Phase 0 proved imported section
' diameters arrive as 6 (swDiameterDimension), NOT 15, and that created
' ordinates carried Type2 of 1 and 7, so ordinate classification accepts
' 1, 7 and 8.
' The reference drawing's own section diameters, including the 47 H7, are
' authored as swLinearDimension (2), not as diameter dimensions. A linear
' dimension can therefore carry a diameter's fit, and type alone must never
' be used to decide whether something is a diameter.
Private Const DIM_ORDINATE As Long = 1
Private Const DIM_LINEAR As Long = 2
Private Const DIM_RADIAL As Long = 5
Private Const DIM_DIAMETER As Long = 6
Private Const DIM_HOR_ORDINATE As Long = 7
Private Const DIM_VERT_ORDINATE As Long = 8

' Standing instruction from the user, 2026-08-01: the tolerances in the
' designers' existing drawings were added manually to signal that SOME
' tolerance is acceptable, not to state that the part holds them. They are
' evidence that a designer typed a number, and nothing more. Every tolerance
' read off a drawing is therefore labelled with this authority so no later
' phase can mistake it for a manufacturing requirement.
' The rule for deciding when a tolerance SHOULD be added is still open; the
' user is establishing with their designer what part information can drive
' it. Nothing here guesses at that rule.
Private Const TOLERANCE_AUTHORITY_DRAWING As String = _
    "DrawingAuthoredNonAuthoritative"
Private Const TOLERANCE_AUTHORITY_NONE As String = "NoTolerance"

' swTolType_e. Note that swTolFIT and swTolMETRIC are BOTH 7 in the 2025
' table, so a value of 7 cannot be reported as one rather than the other.
Private Const TOL_NONE As Long = 0
Private Const TOL_FIT As Long = 7
Private Const TOL_FIT_WITH_TOL As Long = 8
Private Const TOL_FIT_TOL_ONLY As Long = 9

' swDrawingViewTypes_e.
Private Const VIEW_TYPE_SECTION As Long = 2
Private Const VIEW_TYPE_DETAIL As Long = 3
Private Const VIEW_TYPE_PROJECTED As Long = 4
Private Const VIEW_TYPE_AUXILIARY As Long = 5
Private Const VIEW_TYPE_NAMED As Long = 7

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

' R23-401 and R23-402. A view may receive model items when it carries real
' drawing geometry. Primary, side and section all qualify; the sheet's
' standard-view placeholders do not, because they hold no entities.
Public Function IsModelImportEligibleView( _
    ByRef swView As SldWorks.View, _
    ByRef reason As String) As Boolean

    On Error GoTo Failed

    reason = "Unevaluated"

    If swView Is Nothing Then
        reason = "ViewIsNothing"
        Exit Function
    End If

    Dim viewType As Long
    viewType = SafeViewType(swView)

    Select Case viewType
        Case VIEW_TYPE_SECTION, VIEW_TYPE_DETAIL, VIEW_TYPE_PROJECTED, _
             VIEW_TYPE_AUXILIARY, VIEW_TYPE_NAMED
            ' Candidate; the entity test below decides.
        Case Else
            reason = "UnsupportedViewType:" & CStr(viewType)
            Exit Function
    End Select

    If Not ViewHasDrawingGeometry(swView) Then
        reason = "NoVisibleDrawingEntities"
        Exit Function
    End If

    reason = "Eligible"
    IsModelImportEligibleView = True
    Exit Function

Failed:
    reason = "ReadError:" & CStr(Err.Number)
    IsModelImportEligibleView = False
End Function

' R23-401 and R23-403. Ordinate eligibility is stricter than import
' eligibility and is decided from proven projection data, not from view
' names.
' A section view is excluded by type. An isometric is excluded because no
' hole axis is normal to it, which Phase 3 measured directly: the isometric
' resolved 9 anchors and zero axisNormal. Using the graph's own measurement
' rather than a name test means a renamed or differently oriented view is
' still classified correctly.
Public Function IsOrdinateEligibleView( _
    ByRef graph As CLocationGraph, _
    ByRef swView As SldWorks.View, _
    ByRef reason As String) As Boolean

    On Error GoTo Failed

    reason = "Unevaluated"

    If Not IsModelImportEligibleView(swView, reason) Then Exit Function

    If SafeViewType(swView) = VIEW_TYPE_SECTION Then
        reason = "SectionViewOrdinateIneligible"
        Exit Function
    End If

    If graph Is Nothing Then
        reason = "GraphUnavailable"
        Exit Function
    End If

    Dim projections As Collection
    Set projections = graph.ProjectionsForView(SafeViewName(swView))

    Dim normalCount As Long
    Dim i As Long
    For i = 1 To projections.Count
        Dim projection As CViewHoleProjection
        Set projection = projections(i)
        If projection.AxisNormalToView Then normalCount = normalCount + 1
    Next i

    If normalCount = 0 Then
        reason = "NoAxisNormalToView"
        Exit Function
    End If

    reason = "Eligible"
    IsOrdinateEligibleView = True
    Exit Function

Failed:
    reason = "ReadError:" & CStr(Err.Number)
    IsOrdinateEligibleView = False
End Function

' R23-404. Identifies a view whose creation must wait until model import has
' finished. An isometric attracts imported items that belong in the
' orthographic views, so it is created last. The test is the same
' evidence-based one used for ordinate eligibility: a view in which no hole
' axis is normal cannot be an intended dimension host.
Public Function IsDeferredCreationView( _
    ByRef graph As CLocationGraph, _
    ByRef swView As SldWorks.View) As Boolean

    Dim reason As String

    If swView Is Nothing Then Exit Function
    If SafeViewType(swView) = VIEW_TYPE_SECTION Then Exit Function
    If Not ViewHasDrawingGeometry(swView) Then Exit Function

    IsDeferredCreationView = _
        Not IsOrdinateEligibleView(graph, swView, reason)
End Function

' R23-405. MUTATES THE DRAWING. Runs the Phase 0-selected strategy: section
' first with the restricted mask, then side, then primary with the full
' mask, always AllViews=False and DuplicateDims=True.
' allowMutation must be True. The argument exists so that no caller can
' insert annotations by accident, and so a reader can see at the call site
' whether a transaction is read-only.
Public Function ImportModelAnnotations( _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef orderedViews As Collection, _
    ByVal allowMutation As Boolean, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Long

    On Error GoTo Failed

    ImportModelAnnotations = -1

    If Not allowMutation Then
        EmitFailure evidence, "IMPORT_REFUSED|reason=MutationNotAuthorized"
        Exit Function
    End If

    If swDraw Is Nothing Then Exit Function
    If orderedViews Is Nothing Then Exit Function

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim totalInserted As Long

    Dim i As Long
    For i = 1 To orderedViews.Count
        Dim swView As SldWorks.View
        Set swView = orderedViews(i)
        If swView Is Nothing Then GoTo ContinueView

        Dim viewName As String
        viewName = SafeViewName(swView)

        Dim eligibility As String
        If Not IsModelImportEligibleView(swView, eligibility) Then
            EmitInfo evidence, "IMPORT_VIEW_SKIPPED|view=" & viewName & _
                "|reason=" & eligibility
            GoTo ContinueView
        End If

        Dim mask As Long
        If SafeViewType(swView) = VIEW_TYPE_SECTION Then
            mask = IMPORT_MASK_SECTION
        Else
            mask = IMPORT_MASK_FULL
        End If

        ' IDrawingDoc.ActivateView returns False on this build even when the
        ' view does become active. The 2026-08-04 18:45 production run refused
        ' every import on that false negative and inserted zero annotations
        ' into a drawing whose views were activating correctly. Activation is
        ' proved by reading the active view back by name, which is what
        ' Module8_RuntimeSupport.ActivateDrawingView does; its raw setter
        ' result is recorded as a warning, never used as the verdict.
        If Not Module8_RuntimeSupport.ActivateDrawingView( _
            swDraw, swDrawing, swView, evidence, _
            "Annotation import into '" & viewName & "'") Then

            EmitWarning evidence, "IMPORT_VIEW_NOT_ACTIVATED|view=" & viewName
            GoTo ContinueView
        End If

        ' IDrawingDoc.InsertModelAnnotations4 takes EIGHT arguments and
        ' returns an ARRAY of the inserted IAnnotation objects, not a count.
        ' Returning the objects matters: R23-410 can only delete what this
        ' run created, and identity is the only safe way to know that.
        ' DuplicateDims=True means "eliminate duplicates" per the 2025 Help.
        Dim insertedAnnotations As Variant
        insertedAnnotations = swDrawing.InsertModelAnnotations4( _
            IMPORT_FROM_ENTIRE_MODEL, mask, False, True, _
            False, False, False, False)

        Dim insertedCount As Long
        insertedCount = VariantItemCount(insertedAnnotations)
        totalInserted = totalInserted + insertedCount

        RecordCreatedAnnotations _
            insertedAnnotations, viewName, graph, evidence

        If Not evidence Is Nothing Then
            evidence.RecordSolidWorksMutation _
                "InsertModelAnnotations4:" & viewName
        End If

        EmitInfo evidence, "IMPORT_VIEW|view=" & viewName & _
            "|mask=" & CStr(mask) & _
            "|allViews=False|duplicateDims=True" & _
            "|inserted=" & CStr(insertedCount)

ContinueView:
    Next i

    ImportModelAnnotations = totalInserted
    Exit Function

Failed:
    EmitFailure evidence, "IMPORT_ERROR|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description
    ImportModelAnnotations = -1
End Function

' R23-406, first half. Records exactly what the insert returned, tagged as
' R23-created so R23-410 can later delete these and only these.
Private Sub RecordCreatedAnnotations( _
    ByVal insertedAnnotations As Variant, _
    ByVal viewName As String, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    If VariantItemCount(insertedAnnotations) = 0 Then Exit Sub

    Dim i As Long
    For i = LBound(insertedAnnotations) To UBound(insertedAnnotations)
        Dim swAnnotation As SldWorks.Annotation
        Set swAnnotation = Nothing
        On Error Resume Next
        Set swAnnotation = insertedAnnotations(i)
        On Error GoTo Failed

        If swAnnotation Is Nothing Then GoTo ContinueAnnotation

        Dim record As CImportedAnnotation
        Set record = New CImportedAnnotation

        Set record.Annotation = swAnnotation
        record.OwnerViewName = viewName
        record.ProvenanceSource = "Model"

        ClassifyAnnotation swAnnotation, record
        If Not record.DisplayDimension Is Nothing Then
            ReadDimensionSemantics record
        End If

        EmitInfo evidence, "IMPORT_CREATED|view=" & viewName & _
            "|index=" & CStr(i) & _
            "|source=" & record.SourceIdentity

ContinueAnnotation:
    Next i
    Exit Sub

Failed:
    EmitWarning evidence, "IMPORT_CREATED_ERROR|view=" & viewName & _
        "|error=" & CStr(Err.Number)
End Sub

' R23-406, second half. Independent traversal. The import call returns a set
' of objects, and a returned set is not coverage; every annotation actually
' present in each view is enumerated regardless of what any insert reported.
Public Function BuildAnnotationInventory( _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If swDraw Is Nothing Then Exit Function
    If graph Is Nothing Then Exit Function

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet

    Dim views As Variant
    views = swSheet.GetViews

    If IsEmpty(views) Or Not IsArray(views) Then Exit Function

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueView

        If Not ViewHasDrawingGeometry(swView) Then GoTo ContinueView

        InventoryViewAnnotations swView, graph, evidence

ContinueView:
    Next i

    BuildAnnotationInventory = True
    Exit Function

Failed:
    EmitFailure evidence, "INVENTORY_ERROR|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description
End Function

Private Sub InventoryViewAnnotations( _
    ByRef swView As SldWorks.View, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim viewName As String
    viewName = SafeViewName(swView)

    Dim annotations As Variant
    annotations = swView.GetAnnotations

    Dim total As Long
    total = VariantItemCount(annotations)

    EmitInfo evidence, "ANNOTATION_VIEW|view=" & viewName & _
        "|viewType=" & CStr(SafeViewType(swView)) & _
        "|annotations=" & CStr(total) & _
        "|source=IView.GetAnnotations"

    If total = 0 Then Exit Sub

    Dim i As Long
    For i = LBound(annotations) To UBound(annotations)
        Dim swAnnotation As SldWorks.Annotation
        Set swAnnotation = Nothing
        On Error Resume Next
        Set swAnnotation = annotations(i)
        On Error GoTo Failed

        If swAnnotation Is Nothing Then GoTo ContinueAnnotation

        Dim record As CImportedAnnotation
        Set record = New CImportedAnnotation

        Set record.Annotation = swAnnotation
        record.OwnerViewName = viewName

        Dim category As String
        category = ClassifyAnnotation(swAnnotation, record)

        ReadAnnotationPosition swAnnotation, record

        ' Attachments are read here, during inventory. The first run read
        ' them inside the later reconciliation pass, so every ANNOTATION line
        ' reported attachments=0 even for the annotation that went on to
        ' reconcile successfully. Evidence must not be printed before the
        ' field it reports has been populated.
        ReadAnnotationAttachments record

        If category = "DisplayDimension" Then
            ReadDimensionSemantics record
        End If

        graph.AddImportedAnnotation record

        Dim toleranceDetail As String
        toleranceDetail = vbNullString

        If Not record.DisplayDimension Is Nothing Then
            toleranceDetail = _
                "|dimensionCategory=" & ClassifyDimension(record) & _
                "|nominalSource=" & record.NominalSource & _
                "|tolMinM=" & Format$(record.ToleranceMinM, "0.000000000") & _
                "|tolMaxM=" & Format$(record.ToleranceMaxM, "0.000000000") & _
                "|tolStatus=" & CStr(record.ToleranceMinStatus) & _
                    "/" & CStr(record.ToleranceMaxStatus) & _
                "|holeFit=" & record.HoleFit & _
                "|shaftFit=" & record.ShaftFit & _
                "|toleranceAuthority=" & record.ToleranceAuthority & _
                "|tolRead=" & record.ToleranceReadStatus
        End If

        EmitInfo evidence, "ANNOTATION|view=" & viewName & _
            "|index=" & CStr(i) & _
            "|category=" & category & _
            "|" & record.ProvenanceSummary() & toleranceDetail

ContinueAnnotation:
    Next i
    Exit Sub

Failed:
    ' Capture first: SafeViewName contains On Error Resume Next, which
    ' resets the global Err before the concatenation reaches Err.Number.
    Dim annotationErrorNumber As Long
    annotationErrorNumber = Err.Number

    EmitWarning evidence, "ANNOTATION_VIEW_ERROR|view=" & _
        SafeViewName(swView) & _
        "|error=" & CStr(annotationErrorNumber)
End Sub

' R23-408. Distinguishes the annotation categories by swAnnotationType_e and,
' for display dimensions, by whether the item is a native hole callout. A
' hole callout is a display dimension, not a note, so a type test alone would
' merge it with ordinary dimensions.
Public Function ClassifyAnnotation( _
    ByRef swAnnotation As SldWorks.Annotation, _
    ByRef record As CImportedAnnotation) As String

    On Error GoTo Failed

    ClassifyAnnotation = "Unknown"

    If swAnnotation Is Nothing Then
        ClassifyAnnotation = "AnnotationIsNothing"
        Exit Function
    End If

    Dim annotationType As Long
    annotationType = swAnnotation.GetType

    Select Case annotationType
        Case ANN_DISPLAY_DIMENSION
            Dim displayDimension As SldWorks.DisplayDimension
            Set displayDimension = Nothing
            On Error Resume Next
            Set displayDimension = swAnnotation.GetSpecificAnnotation
            On Error GoTo Failed

            If displayDimension Is Nothing Then
                ClassifyAnnotation = "DisplayDimensionUnavailable"
                Exit Function
            End If

            Set record.DisplayDimension = displayDimension
            ClassifyAnnotation = "DisplayDimension"

        Case ANN_GTOL
            ClassifyAnnotation = "GTol"
        Case ANN_DATUM_TAG
            ClassifyAnnotation = "DatumTag"
        Case ANN_DATUM_ORIGIN
            ClassifyAnnotation = "DatumOrigin"
        Case ANN_NOTE
            ClassifyAnnotation = "Note"
        Case ANN_CTHREAD
            ClassifyAnnotation = "CosmeticThread"
        Case ANN_CENTER_MARK
            ClassifyAnnotation = "CenterMark"
        Case ANN_CENTER_LINE
            ClassifyAnnotation = "CenterLine"
        Case Else
            ClassifyAnnotation = "Other:" & CStr(annotationType)
    End Select

    Exit Function

Failed:
    ClassifyAnnotation = "ClassifyError:" & CStr(Err.Number)
End Function

' Returns the sub-category of a display dimension. Kept separate from
' ClassifyAnnotation so an ordinate is never silently counted as an ordinary
' linear dimension.
Public Function ClassifyDimension( _
    ByRef record As CImportedAnnotation) As String

    If record Is Nothing Then
        ClassifyDimension = "Unknown"
        Exit Function
    End If

    If record.IsHoleCallout Then
        ClassifyDimension = "NativeHoleCallout"
        Exit Function
    End If

    Select Case record.DimensionType2
        Case DIM_ORDINATE, DIM_HOR_ORDINATE, DIM_VERT_ORDINATE
            ClassifyDimension = "Ordinate"
        Case DIM_DIAMETER
            ClassifyDimension = "Diameter"
        Case DIM_RADIAL
            ClassifyDimension = "Radial"
        Case DIM_LINEAR
            ClassifyDimension = "Linear"
        Case Else
            ClassifyDimension = "Dimension:" & CStr(record.DimensionType2)
    End Select
End Function

' Labels how much authority a tolerance read from a drawing carries. Every
' nonzero tolerance found on an existing drawing gets the non-authoritative
' label, per the standing instruction. There is deliberately no branch that
' can return a stronger authority: nothing R23 can currently read
' distinguishes a binding tolerance from an indicative one, and inventing a
' distinction would manufacture authority that does not exist.
Public Function ClassifyToleranceAuthority( _
    ByRef record As CImportedAnnotation) As String

    ClassifyToleranceAuthority = TOLERANCE_AUTHORITY_NONE

    If record Is Nothing Then Exit Function
    If Not record.HasNonZeroTolerance() Then
        If Not record.HasFitData() Then Exit Function
    End If

    ClassifyToleranceAuthority = TOLERANCE_AUTHORITY_DRAWING
End Function

' R23-409. Reads the typed manufacturing content of a display dimension:
' type, hole-callout flag, nominal, tolerance type and bounds, and fit.
Public Sub ReadDimensionSemantics( _
    ByRef record As CImportedAnnotation)

    On Error GoTo Failed

    record.ToleranceReadStatus = "NotRead"

    If record Is Nothing Then Exit Sub
    If record.DisplayDimension Is Nothing Then Exit Sub

    Dim displayDimension As SldWorks.DisplayDimension
    Set displayDimension = record.DisplayDimension

    record.DimensionType2 = displayDimension.Type2

    record.IsHoleCallout = Module11_GeometryIdentity.NormalizeSwBoolean( _
        displayDimension.IsHoleCallout)

    Dim swDimension As SldWorks.Dimension
    Set swDimension = Nothing
    On Error Resume Next
    Set swDimension = displayDimension.GetDimension
    On Error GoTo Failed

    If swDimension Is Nothing Then
        record.ToleranceReadStatus = "DimensionUnavailable"
        Exit Sub
    End If

    Set record.Dimension = swDimension
    record.SourceIdentity = SafeDimensionName(swDimension)

    ' Nominal value. The first read-only run returned 0 for EVERY dimension
    ' from GetSystemValue3(swThisConfiguration, Empty): these are drawing
    ' dimensions (RD4@Drawing View6@....Drawing) and a drawing document has
    ' no configurations, so a configuration-scoped read has nothing to
    ' resolve. IDimension.SystemValue is the configuration-free property and
    ' is read first. Both are recorded so the disagreement stays visible
    ' rather than being asserted away.
    Dim systemValue As Double
    Dim configuredValue As Variant

    systemValue = 0#
    On Error Resume Next
    systemValue = swDimension.SystemValue
    configuredValue = swDimension.GetSystemValue3(swThisConfiguration, Empty)
    On Error GoTo Failed

    record.NominalM = systemValue
    record.NominalAvailable = (systemValue <> 0#)

    If IsNumeric(configuredValue) Then
        record.NominalSource = "IDimension.SystemValue=" & _
            Format$(systemValue, "0.000000000") & _
            "|GetSystemValue3=" & Format$(CDbl(configuredValue), "0.000000000")
    Else
        record.NominalSource = "IDimension.SystemValue=" & _
            Format$(systemValue, "0.000000000") & _
            "|GetSystemValue3=NonNumeric"
    End If

    Dim tolerance As SldWorks.DimensionTolerance
    Set tolerance = Nothing
    On Error Resume Next
    Set tolerance = swDimension.Tolerance
    On Error GoTo Failed

    If tolerance Is Nothing Then
        record.ToleranceType = TOL_NONE
        record.ToleranceReadStatus = "NoToleranceObject"
        Exit Sub
    End If

    record.ToleranceType = tolerance.Type

    ' GetMinValue2/GetMaxValue2 return the STATUS
    ' (swDimensionToleranceWarning_e) and deliver the value through an out
    ' parameter. The out parameter must be a local: a class Public variable
    ' is exposed as a property, so passing record.ToleranceMinM directly
    ' would hand the callee a temporary and silently discard the value.
    ' This is the same trap that produced projectedAxis=0,0,0 in Phase 3.
    Dim minValue As Double
    Dim maxValue As Double

    record.ToleranceMinStatus = tolerance.GetMinValue2(minValue)
    record.ToleranceMaxStatus = tolerance.GetMaxValue2(maxValue)
    record.ToleranceMinM = minValue
    record.ToleranceMaxM = maxValue

    ' A fit is only claimed when the tolerance type says so. swTolFIT and
    ' swTolMETRIC share the value 7 in the 2025 enum, so 7 is recorded as
    ' "fit or metric" and never reported as one specifically.
    Select Case record.ToleranceType
        Case TOL_FIT, TOL_FIT_WITH_TOL, TOL_FIT_TOL_ONLY
            record.FitType = tolerance.FitType
            record.FitDisplayStyle = tolerance.FitDisplayStyle
            record.HoleFit = Trim$(CStr(tolerance.GetHoleFitValue))
            record.ShaftFit = Trim$(CStr(tolerance.GetShaftFitValue))
            record.FitValues = "type=" & CStr(record.ToleranceType) & _
                IIf(record.ToleranceType = TOL_FIT, _
                    "(swTolFIT/swTolMETRIC ambiguous)", "")
    End Select

    record.ToleranceReadStatus = "Read"
    record.ToleranceAuthority = ClassifyToleranceAuthority(record)
    Exit Sub

Failed:
    record.ToleranceReadStatus = "ReadError:" & CStr(Err.Number)
End Sub

' R23-407 and R23-411. Reconciles imported display dimensions against the
' location graph.
' Matching is by ATTACHED ENTITY identity against each projection's proven
' drawing anchor, not by page proximity. Proximity would attach a dimension
' to whichever hole happened to be nearest, which is exactly the failure the
' physical-location model exists to prevent.
Public Function ReconcileWithLocationGraph( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Long

    On Error GoTo Failed

    ReconcileWithLocationGraph = 0

    If graph Is Nothing Then Exit Function

    Dim annotations As Collection
    Set annotations = graph.Annotations()

    Dim seenKeys As Object
    Set seenKeys = CreateObject("Scripting.Dictionary")
    seenKeys.CompareMode = 1

    Dim reconciled As Long

    Dim i As Long
    For i = 1 To annotations.Count
        Dim record As CImportedAnnotation
        Set record = annotations(i)

        If record.DisplayDimension Is Nothing Then
            record.ReconciliationStatus = "NotADisplayDimension"
            GoTo ContinueAnnotation
        End If

        ' R23-411. DuplicateDims=True suppresses duplicates the API knows
        ' about; it does not suppress two imports that carry the same model
        ' dimension identity into the same view.
        Dim identityKey As String
        identityKey = record.AnnotationIdentityKey()

        If seenKeys.Exists(identityKey) Then
            record.ReconciliationStatus = "DuplicateModelDimensionIdentity"
            EmitWarning evidence, "ANNOTATION_DUPLICATE|key=" & identityKey
            GoTo ContinueAnnotation
        End If

        seenKeys.Add identityKey, True

        Dim matchDiagnostics As String
        Dim reverseDiagnostics As String
        Dim matchRoute As String
        Dim matchedProjection As CViewHoleProjection

        reverseDiagnostics = "reverse=NotRun"

        Set matchedProjection = MatchAnnotationToProjection( _
            swApp, graph, record, matchDiagnostics)

        If Not matchedProjection Is Nothing Then
            matchRoute = "ForwardAlias"
        Else
            ' Route D. The forward map is partial: the first instrumented
            ' run showed the counterbore callout attaches to a drawing EDGE
            ' (swSelEDGES) that is none of the 18 aliases
            ' IView.GetCorrespondingEntity produced for that view. The 2025
            ' Help names the reverse member for exactly this direction, and
            ' R23-302 asked for it: map the DRAWING entity back to the model
            ' and test ownership there instead.
            Set matchedProjection = MatchByReverseCorrespondence( _
                swApp, swDraw, graph, record, reverseDiagnostics)

            If Not matchedProjection Is Nothing Then
                matchRoute = "ReverseCorrespondence"
            End If
        End If

        If matchedProjection Is Nothing Then
            ' Distinguish "this drawing entity has no model counterpart at
            ' all" from "it has one and no location owns it". Only the
            ' second would indicate a defect in the ownership model.
            If InStr(reverseDiagnostics, _
                "UnavailableNoModelCounterpart") > 0 Then

                record.ReconciliationStatus = _
                    "AuthoredDrawingEntityNoModelCounterpart"
            Else
                record.ReconciliationStatus = "NoAttachedProjection"
            End If

            EmitInfo evidence, "ANNOTATION_UNMATCHED|view=" & _
                record.OwnerViewName & _
                "|source=" & record.SourceIdentity & _
                "|category=" & ClassifyDimension(record) & _
                "|routesTried=ForwardAlias,ReverseCorrespondence" & _
                "|" & matchDiagnostics & _
                "|" & reverseDiagnostics
            GoTo ContinueAnnotation
        End If

        matchedProjection.AttachAnnotation record
        record.CoversRequirementKey = matchedProjection.ProjectionKey()
        record.ReconciliationStatus = "AttachedToProjection:" & matchRoute
        reconciled = reconciled + 1

        EmitInfo evidence, "ANNOTATION_RECONCILED|view=" & _
            record.OwnerViewName & _
            "|source=" & record.SourceIdentity & _
            "|matchRoute=" & matchRoute & _
            "|covers=" & record.CoversRequirementKey & _
            "|dimensionCategory=" & ClassifyDimension(record)

ContinueAnnotation:
    Next i

    ReconcileWithLocationGraph = reconciled
    Exit Function

Failed:
    EmitFailure evidence, "RECONCILE_ERROR|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description
    ReconcileWithLocationGraph = -1
End Function

' Attaches by COM identity between the dimension's attached entities and a
' projection's proven anchor. Returns Nothing when nothing matches, which is
' a real answer rather than a fallback to the nearest hole.
Private Function MatchAnnotationToProjection( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef graph As CLocationGraph, _
    ByRef record As CImportedAnnotation, _
    ByRef diagnostics As String) As CViewHoleProjection

    On Error GoTo Failed

    diagnostics = "attachments=0"

    If record.Annotation Is Nothing Then Exit Function
    If record.AttachedEntityCount = 0 Then Exit Function

    Dim attachments As Variant
    attachments = record.Annotation.GetAttachedEntities3

    If VariantItemCount(attachments) = 0 Then Exit Function

    Dim projections As Collection
    Set projections = graph.ProjectionsForView(record.OwnerViewName)

    ' Diagnostic accumulation. The alias fix did not raise the reconciled
    ' count, so the evidence must say WHAT each attachment actually is
    ' rather than only that it failed to match. swSelectType_e: 1 edge,
    ' 2 face, 3 vertex, 10 sketch segment, 11 sketch point, 28 centre mark,
    ' 46 silhouette.
    Dim anchoredProjections As Long
    Dim totalAliases As Long
    Dim attachmentTypes As String

    Dim d As Long
    For d = LBound(attachments) To UBound(attachments)
        Dim probe As Object
        Set probe = Nothing
        On Error Resume Next
        Set probe = attachments(d)
        On Error GoTo Failed

        If Len(attachmentTypes) > 0 Then
            attachmentTypes = attachmentTypes & ","
        End If

        attachmentTypes = attachmentTypes & CStr(SafeEntityType(probe))
    Next d

    Dim p As Long
    For p = 1 To projections.Count
        Dim countedProjection As CViewHoleProjection
        Set countedProjection = projections(p)
        If countedProjection.HasSelectableAnchor() Then
            anchoredProjections = anchoredProjections + 1
            totalAliases = totalAliases + _
                countedProjection.DrawingEntityAliases.Count
        End If
    Next p

    diagnostics = "attachments=" & CStr(VariantItemCount(attachments)) & _
        "|attachmentTypes=" & attachmentTypes & _
        "|anchoredProjections=" & CStr(anchoredProjections) & _
        "|aliasesAvailable=" & CStr(totalAliases)

    Dim i As Long
    For i = 1 To projections.Count
        Dim projection As CViewHoleProjection
        Set projection = projections(i)

        If Not projection.HasSelectableAnchor() Then GoTo ContinueProjection

        Dim j As Long
        For j = LBound(attachments) To UBound(attachments)
            Dim attached As Object
            Set attached = Nothing
            On Error Resume Next
            Set attached = attachments(j)
            On Error GoTo Failed

            If attached Is Nothing Then GoTo ContinueAttachment

            ' Match against EVERY drawing entity this projection owns, not
            ' only the chosen anchor. A counterbore maps two edges, and the
            ' native hole callout attaches to the 11 mm mouth while the
            ' anchor tier deliberately prefers the 6.6 mm through hole. The
            ' first run reconciled 1 of 38 for exactly this reason: the
            ' callout and the anchor are different entities of the same
            ' location. Ownership is what reconciliation needs; the anchor
            ' is only the preferred attachment point for NEW annotations.
            If ProjectionOwnsEntity(swApp, projection, attached) Then
                Set MatchAnnotationToProjection = projection
                Exit Function
            End If

ContinueAttachment:
        Next j

ContinueProjection:
    Next i
    Exit Function

Failed:
    Set MatchAnnotationToProjection = Nothing
End Function

' Route D. Maps each attached DRAWING entity back to the model with
' IModelDocExtension.GetCorrespondingEntity2, then tests that model entity
' against the geometry each physical location actually owns.
' This is the direction R23-302 asked for and the 2025 Help recommends for
' drawing-to-model resolution. It matters because the forward map is
' partial: the instrumented run showed the counterbore hole callout attached
' to a drawing edge of type swSelEDGES that was none of the 18 aliases
' IView.GetCorrespondingEntity had produced for that view. It also reaches
' views where the forward map produced nothing at all, such as the section
' view, whose projections have no anchors.
' Still identity only. A location owns the entity or it does not; there is no
' positional or dimensional fallback anywhere in this path.
Private Function MatchByReverseCorrespondence( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef graph As CLocationGraph, _
    ByRef record As CImportedAnnotation, _
    ByRef diagnostics As String) As CViewHoleProjection

    On Error GoTo Failed

    diagnostics = "reverse=NoAttachments"

    If swDraw Is Nothing Then
        diagnostics = "reverse=NoDrawingDoc"
        Exit Function
    End If

    If record.Annotation Is Nothing Then Exit Function
    If record.AttachedEntityCount = 0 Then Exit Function

    Dim attachments As Variant
    attachments = record.Annotation.GetAttachedEntities3
    If VariantItemCount(attachments) = 0 Then Exit Function

    Dim projections As Collection
    Set projections = graph.ProjectionsForView(record.OwnerViewName)

    ' Per-attachment outcome. The previous run could not distinguish
    ' "GetCorrespondingEntity2 returned Nothing" from "it resolved and no
    ' location owned the result", so both hypotheses survived a live run.
    ' eqMax carries the strongest swObjectEquality any IsSame comparison
    ' returned: 0 NotSame, 1 Same, 2 Unsupported. A 2 means the comparison
    ' itself is invalid across these documents, which reads exactly like a
    ' non-match in ObjectsAreSame and would explain a silent zero.
    Dim outcomes As String
    Dim eqMax As Long
    Dim resolvedCount As Long
    Dim modelEdgesTested As Long

    eqMax = -1

    Dim j As Long
    For j = LBound(attachments) To UBound(attachments)
        Dim drawingEntity As Object
        Set drawingEntity = Nothing
        On Error Resume Next
        Set drawingEntity = attachments(j)
        On Error GoTo Failed

        If Len(outcomes) > 0 Then outcomes = outcomes & ","

        If drawingEntity Is Nothing Then
            outcomes = outcomes & "nullAttachment"
            GoTo ContinueAttachment
        End If

        Dim modelEntity As Object
        Dim resolveError As Long
        Set modelEntity = Nothing
        On Error Resume Next
        Err.Clear
        Set modelEntity = _
            swDraw.Extension.GetCorrespondingEntity2(drawingEntity)
        resolveError = Err.Number
        On Error GoTo Failed

        If modelEntity Is Nothing Then
            outcomes = outcomes & "draw" & _
                CStr(SafeEntityType(drawingEntity)) & _
                ":unresolved:err" & CStr(resolveError)
            GoTo ContinueAttachment
        End If

        resolvedCount = resolvedCount + 1
        outcomes = outcomes & "draw" & _
            CStr(SafeEntityType(drawingEntity)) & _
            ":model" & CStr(SafeEntityType(modelEntity))

        Dim i As Long
        For i = 1 To projections.Count
            Dim projection As CViewHoleProjection
            Set projection = projections(i)

            If LocationOwnsModelEntity( _
                swApp, projection.PhysicalLocation, modelEntity, _
                eqMax, modelEdgesTested) Then

                diagnostics = "reverse=Matched|outcomes=" & outcomes
                Set MatchByReverseCorrespondence = projection
                Exit Function
            End If
        Next i

ContinueAttachment:
    Next j

    Dim verdict As String
    If resolvedCount = 0 Then
        ' Proven live: 38 of 38 annotations, every attachment, error 0. The
        ' call does not fail, it declines. This is a PART drawing and the
        ' member is documented as returning the entity "in the underlying
        ' part or subassembly" - there is no underlying component to descend
        ' into, which matches componentContext=DrawingContextOnly.
        verdict = "reverse=UnavailableNoModelCounterpart"
    Else
        verdict = "reverse=NoOwner"
    End If

    diagnostics = verdict & "|outcomes=" & outcomes & _
        "|resolved=" & CStr(resolvedCount) & _
        "|projectionsInView=" & CStr(projections.Count) & _
        "|modelEdgesTested=" & CStr(modelEdgesTested) & _
        "|eqMax=" & CStr(eqMax)
    Exit Function

Failed:
    diagnostics = "reverse=Error:" & CStr(Err.Number)
    Set MatchByReverseCorrespondence = Nothing
End Function

' True when the model entity is one of the location's own cylindrical faces
' or one of those faces' boundary edges.
Private Function LocationOwnsModelEntity( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef location As CPhysicalHoleLocation, _
    ByRef modelEntity As Object, _
    ByRef eqMax As Long, _
    ByRef comparisons As Long) As Boolean

    On Error GoTo Failed

    If location Is Nothing Then Exit Function
    If modelEntity Is Nothing Then Exit Function

    Dim f As Long
    For f = 1 To location.SourceFaces.Count
        Dim modelFace As SldWorks.Face2
        Set modelFace = Nothing
        On Error Resume Next
        Set modelFace = location.SourceFaces(f)
        On Error GoTo Failed

        If modelFace Is Nothing Then GoTo ContinueFace

        comparisons = comparisons + 1
        If RecordEquality(swApp, modelEntity, modelFace, eqMax) Then
            LocationOwnsModelEntity = True
            Exit Function
        End If

        Dim edges As Variant
        On Error Resume Next
        edges = modelFace.GetEdges
        On Error GoTo Failed

        If IsEmpty(edges) Or Not IsArray(edges) Then GoTo ContinueFace

        Dim e As Long
        For e = LBound(edges) To UBound(edges)
            Dim modelEdge As Object
            Set modelEdge = Nothing
            On Error Resume Next
            Set modelEdge = edges(e)
            On Error GoTo Failed

            If Not modelEdge Is Nothing Then
                comparisons = comparisons + 1
                If RecordEquality( _
                    swApp, modelEntity, modelEdge, eqMax) Then

                    LocationOwnsModelEntity = True
                    Exit Function
                End If
            End If
        Next e

ContinueFace:
    Next f
    Exit Function

Failed:
    LocationOwnsModelEntity = False
End Function

' True when the entity is the projection's anchor or any other drawing entity
' the projection mapped. Identity only; no geometric or positional fallback.
Private Function ProjectionOwnsEntity( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef projection As CViewHoleProjection, _
    ByRef candidate As Object) As Boolean

    On Error GoTo Failed

    If projection Is Nothing Then Exit Function
    If candidate Is Nothing Then Exit Function

    If ObjectsAreSame(swApp, candidate, projection.PrimaryAnchor) Then
        ProjectionOwnsEntity = True
        Exit Function
    End If

    Dim i As Long
    For i = 1 To projection.DrawingEntityAliases.Count
        ' Not named "alias": that is a VBA reserved word (Declare ... Alias).
        Dim aliasEntity As Object
        Set aliasEntity = Nothing
        On Error Resume Next
        Set aliasEntity = projection.DrawingEntityAliases(i)
        On Error GoTo Failed

        If Not aliasEntity Is Nothing Then
            If ObjectsAreSame(swApp, candidate, aliasEntity) Then
                ProjectionOwnsEntity = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

Failed:
    ProjectionOwnsEntity = False
End Function

' ISldWorks.IsSame returns swObjectEquality, not a Boolean: 0 not same,
' 1 same, 2 unable to determine. Only an exact 1 proves identity.
Private Function ObjectsAreSame( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef first As Object, _
    ByRef second As Object) As Boolean

    On Error GoTo Failed

    If first Is Nothing Or second Is Nothing Then Exit Function

    ObjectsAreSame = (CLng(swApp.IsSame(first, second)) = 1)
    Exit Function

Failed:
    ObjectsAreSame = False
End Function

' ISldWorks.IsSame returns swObjectEquality, NOT a Boolean: 0 NotSame,
' 1 Same, 2 Unsupported. Only 1 is a match. Unsupported is reported
' separately through eqMax because it means the comparison could not be
' performed at all, which is a different failure from a genuine non-match
' and is indistinguishable from one in the Boolean wrapper.
Private Function RecordEquality( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef first As Object, _
    ByRef second As Object, _
    ByRef eqMax As Long) As Boolean

    Dim equality As Long
    equality = -1

    On Error Resume Next
    equality = CLng(swApp.IsSame(first, second))
    On Error GoTo 0

    If equality > eqMax Then eqMax = equality

    RecordEquality = (equality = 1)
End Function

' R23-410. MUTATES THE DRAWING. Deletes only annotations this run created and
' recorded, identified by object identity against the supplied collection.
' Nothing is matched by name, position or appearance, so pre-existing manual
' content can never be selected for deletion.
Public Function RemoveR23CreatedAnnotations( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef createdAnnotations As Collection, _
    ByVal allowMutation As Boolean, _
    ByRef evidence As CRunEvidence) As Long

    On Error GoTo Failed

    RemoveR23CreatedAnnotations = -1

    If Not allowMutation Then
        EmitFailure evidence, "CLEANUP_REFUSED|reason=MutationNotAuthorized"
        Exit Function
    End If

    If swDraw Is Nothing Then Exit Function
    If createdAnnotations Is Nothing Then Exit Function

    Dim removed As Long

    Dim i As Long
    For i = 1 To createdAnnotations.Count
        Dim record As CImportedAnnotation
        Set record = createdAnnotations(i)

        If record Is Nothing Then GoTo ContinueRecord
        If record.Annotation Is Nothing Then GoTo ContinueRecord

        ' Only this run's own records reach here, and each still carries the
        ' live annotation object it was built from.
        swDraw.ClearSelection2 True

        If Module11_GeometryIdentity.NormalizeSwBoolean( _
            record.Annotation.Select3(False, Nothing)) Then

            If Module11_GeometryIdentity.NormalizeSwBoolean( _
                swDraw.Extension.DeleteSelection2(0)) Then

                removed = removed + 1

                If Not evidence Is Nothing Then
                    evidence.RecordSolidWorksMutation _
                        "DeleteSelection2:" & record.SourceIdentity
                End If
            End If
        End If

        swDraw.ClearSelection2 True

ContinueRecord:
    Next i

    EmitInfo evidence, "CLEANUP|requested=" & _
        CStr(createdAnnotations.Count) & _
        "|removed=" & CStr(removed)

    RemoveR23CreatedAnnotations = removed
    Exit Function

Failed:
    EmitFailure evidence, "CLEANUP_ERROR|error=" & CStr(Err.Number)
    On Error Resume Next
    swDraw.ClearSelection2 True
    RemoveR23CreatedAnnotations = -1
End Function

' R23-412. Success is required-category and required-view coverage, never a
' count. An import that returns 25 items and covers no required category has
' failed.
Public Function VerifyRequiredCoverage( _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As String

    On Error GoTo Failed

    Dim failures As String

    If graph Is Nothing Then
        VerifyRequiredCoverage = "GraphMissing"
        Exit Function
    End If

    Dim annotations As Collection
    Set annotations = graph.Annotations()

    Dim holeCallouts As Long
    Dim ordinates As Long
    Dim diameters As Long
    Dim toleranced As Long
    Dim withFit As Long
    Dim reconciled As Long

    Dim i As Long
    For i = 1 To annotations.Count
        Dim record As CImportedAnnotation
        Set record = annotations(i)

        If record.DisplayDimension Is Nothing Then GoTo ContinueRecord

        Select Case ClassifyDimension(record)
            Case "NativeHoleCallout"
                holeCallouts = holeCallouts + 1
            Case "Ordinate"
                ordinates = ordinates + 1
            Case "Diameter"
                diameters = diameters + 1
        End Select

        If record.HasNonZeroTolerance() Then toleranced = toleranced + 1
        If record.HasFitData() Then withFit = withFit + 1

        If record.ReconciliationStatus = "AttachedToProjection" Then
            reconciled = reconciled + 1
        End If

ContinueRecord:
    Next i

    EmitInfo evidence, "COVERAGE|holeCallouts=" & CStr(holeCallouts) & _
        "|ordinates=" & CStr(ordinates) & _
        "|diameters=" & CStr(diameters) & _
        "|toleranced=" & CStr(toleranced) & _
        "|withFit=" & CStr(withFit) & _
        "|reconciledToProjections=" & CStr(reconciled)

    If holeCallouts = 0 Then
        failures = AppendFailure(failures, "NoNativeHoleCallout")
    End If

    If toleranced = 0 Then
        failures = AppendFailure(failures, "NoTolerancedDimension")
    End If

    If ordinates = 0 Then
        failures = AppendFailure(failures, "NoOrdinateDimension")
    End If

    VerifyRequiredCoverage = failures
    Exit Function

Failed:
    VerifyRequiredCoverage = "CoverageReadError:" & CStr(Err.Number)
End Function

' GetAttachedEntities3 lives on IAnnotation, not IDisplayDimension. The
' inventory already holds the IAnnotation, so no GetAnnotation round trip is
' needed.
Private Sub ReadAnnotationAttachments( _
    ByRef record As CImportedAnnotation)

    On Error GoTo Failed

    record.AttachedEntityCount = 0
    record.AttachmentFingerprint = "None"

    If record Is Nothing Then Exit Sub
    If record.Annotation Is Nothing Then Exit Sub

    Dim attachments As Variant
    attachments = record.Annotation.GetAttachedEntities3

    record.AttachedEntityCount = VariantItemCount(attachments)

    If record.AttachedEntityCount > 0 Then
        record.AttachmentFingerprint = "count=" & _
            CStr(record.AttachedEntityCount) & _
            "|source=IAnnotation.GetAttachedEntities3"
    End If
    Exit Sub

Failed:
    record.AttachedEntityCount = 0
    record.AttachmentFingerprint = "ReadError:" & CStr(Err.Number)
End Sub

Private Sub ReadAnnotationPosition( _
    ByRef swAnnotation As SldWorks.Annotation, _
    ByRef record As CImportedAnnotation)

    On Error GoTo Failed

    record.PositionReadStatus = "NotRead"

    Dim position As Variant
    position = swAnnotation.GetPosition

    If VariantItemCount(position) < 2 Then
        record.PositionReadStatus = "PositionUnavailable"
        Exit Sub
    End If

    Dim baseIndex As Long
    baseIndex = LBound(position)

    record.PageX = CDbl(position(baseIndex))
    record.PageY = CDbl(position(baseIndex + 1))
    record.PositionReadStatus = "Page|source=IAnnotation.GetPosition"
    Exit Sub

Failed:
    record.PositionReadStatus = "ReadError:" & CStr(Err.Number)
End Sub

Private Function ViewHasDrawingGeometry( _
    ByRef swView As SldWorks.View) As Boolean

    On Error GoTo Failed

    If swView Is Nothing Then Exit Function

    ' GetVisibleEntityCount2 requires a Component2. A part drawing supplies
    ' none on this build, so calling it with Nothing makes every real view
    ' look empty and suppresses the entire annotation inventory. Sheet
    ' orientation placeholders have * names; a real drawing view additionally
    ' has a non-zero page-frame outline.
    If Left$(SafeViewName(swView), 1) = "*" Then Exit Function

    Dim outline As Variant
    outline = swView.GetOutline
    If VariantItemCount(outline) < 4 Then Exit Function

    Dim widthM As Double
    Dim heightM As Double
    widthM = Abs(CDbl(outline(2)) - CDbl(outline(0)))
    heightM = Abs(CDbl(outline(3)) - CDbl(outline(1)))

    ViewHasDrawingGeometry = (widthM > 0# And heightM > 0#)
    Exit Function

Failed:
    ViewHasDrawingGeometry = False
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

Private Function VariantItemCount(ByVal value As Variant) As Long
    On Error GoTo Failed

    If IsEmpty(value) Then Exit Function
    If Not IsArray(value) Then Exit Function

    VariantItemCount = UBound(value) - LBound(value) + 1
    Exit Function

Failed:
    VariantItemCount = 0
End Function

' swSelectType_e of an attached entity, or -1 when unreadable. Reported so an
' unmatched attachment can be attributed to WHAT it is rather than left as an
' unexplained failure.
Private Function SafeEntityType( _
    ByRef candidate As Object) As Long

    On Error Resume Next
    SafeEntityType = -1
    If candidate Is Nothing Then Exit Function
    SafeEntityType = candidate.GetType
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

Private Function SafeViewType( _
    ByRef swView As SldWorks.View) As Long

    On Error Resume Next
    SafeViewType = -1
    If swView Is Nothing Then Exit Function
    SafeViewType = swView.Type
End Function

Private Function SafeDimensionName( _
    ByRef swDimension As SldWorks.Dimension) As String

    On Error Resume Next
    SafeDimensionName = "Unnamed"
    If swDimension Is Nothing Then Exit Function
    SafeDimensionName = swDimension.FullName
End Function

' READ-ONLY evidence entry point for Phase 4.
' Inventories, classifies and reconciles the annotations the drawing already
' has. It never imports and never deletes, so it is safe to run against a
' manual reference drawing. The mutation counters at the end are the proof of
' that, not a claim about it.
Public Sub R23_ProbeAnnotationReconciliation()
    On Error GoTo Failed

    mEmitDiagnostics = False

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_ANNOTATION_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swDraw As SldWorks.ModelDoc2
    Set swDraw = swApp.ActiveDoc

    If swDraw Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_ANNOTATION_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swDraw.GetType <> swDocDRAWING Then
        Module21_EvidenceSink.LogLine _
            "R23_ANNOTATION_FATAL|reason=ActiveDocumentNotDrawing"
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
            "R23_ANNOTATION_FATAL|reason=NoViewsOnSheet"
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
        Module21_EvidenceSink.LogLine _
            "R23_ANNOTATION_FATAL|reason=NoReferencedDocument"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Module21_EvidenceSink.LogLine _
            "R23_ANNOTATION_FATAL|reason=UnauthorizedFixture" & _
            "|path=" & partPath
        Exit Sub
    End If

    Dim drawingSaveBefore As Boolean
    drawingSaveBefore = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Dim initialSelectionCount As Long
    initialSelectionCount = swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim evidence As CRunEvidence
    Set evidence = New CRunEvidence

    Dim graph As CLocationGraph
    Set graph = New CLocationGraph

    Module21_EvidenceSink.LogLine _
        "R23_ANNOTATION_BEGIN|drawing=" & swDraw.GetPathName & _
        "|part=" & partPath & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath) & _
        "|mode=ReadOnly|imports=0|deletions=0"

    Dim configurationName As String
    configurationName = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    If Not Module12_FeatureQualification.BuildFeatureCatalog( _
        swApp, swPart, configurationName, graph, evidence) Then

        Module21_EvidenceSink.LogLine _
            "R23_ANNOTATION_FATAL|reason=CatalogUnavailable"
        Exit Sub
    End If

    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueProjectionView

        Module13_ProjectionResolution.BuildViewProjections _
            swApp, swDraw, swView, graph, evidence

ContinueProjectionView:
    Next i

    Module21_EvidenceSink.LogLine _
        "R23_ANNOTATION_GRAPH|" & graph.GraphSummary()

    ' Eligibility is reported for every real view before anything is read,
    ' so the policy is visible independently of what the drawing contains.
    For i = LBound(views) To UBound(views)
        Dim policyView As SldWorks.View
        Set policyView = views(i)
        If policyView Is Nothing Then GoTo ContinuePolicyView
        If Not ViewHasDrawingGeometry(policyView) Then GoTo ContinuePolicyView

        Dim importReason As String
        Dim ordinateReason As String
        Dim importEligible As Boolean
        Dim ordinateEligible As Boolean

        importEligible = IsModelImportEligibleView(policyView, importReason)
        ordinateEligible = IsOrdinateEligibleView( _
            graph, policyView, ordinateReason)

        Module21_EvidenceSink.LogLine _
            "R23_ANNOTATION_POLICY|view=" & _
            SafeViewName(policyView) & _
            "|viewType=" & CStr(SafeViewType(policyView)) & _
            "|importEligible=" & CStr(importEligible) & _
            "|importReason=" & importReason & _
            "|ordinateEligible=" & CStr(ordinateEligible) & _
            "|ordinateReason=" & ordinateReason & _
            "|deferredCreation=" & _
                CStr(IsDeferredCreationView(graph, policyView))

ContinuePolicyView:
    Next i

    BuildAnnotationInventory swDraw, graph, evidence

    Dim reconciled As Long
    reconciled = ReconcileWithLocationGraph(swApp, swDraw, graph, evidence)

    Dim coverageFailures As String
    coverageFailures = VerifyRequiredCoverage(graph, evidence)

    Dim drawingSaveAfter As Boolean
    drawingSaveAfter = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Dim finalSelectionCount As Long
    finalSelectionCount = swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Module21_EvidenceSink.LogLine _
        "R23_ANNOTATION_END|" & graph.GraphSummary() & _
        "|reconciled=" & CStr(reconciled) & _
        "|coverageFailures=" & _
            IIf(Len(coverageFailures) = 0, "None", coverageFailures) & _
        "|mutations=" & CStr(evidence.SolidWorksMutationSequence) & _
        "|initialSelectionCount=" & CStr(initialSelectionCount) & _
        "|finalSelectionCount=" & CStr(finalSelectionCount) & _
        "|drawingUnchanged=" & CStr(drawingSaveBefore = drawingSaveAfter)
    Exit Sub

Failed:
    Module21_EvidenceSink.LogLine _
        "R23_ANNOTATION_FATAL|reason=UnhandledError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description
End Sub

' R23 probe-runner compile-failure localisation. A no-op; VBA compiles
' at module granularity, so a module that loads this has compiled.
Public Sub R23_CompileTouch()
End Sub
