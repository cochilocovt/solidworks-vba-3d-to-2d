Option Explicit

' R23 Phase 8. Semantic section-dimension engine.
'
' SAFETY BOUNDARY. Exactly two procedures change a drawing:
' CreateSectionDimension and ApplyReferenceFit. Both refuse unless passed an
' explicit allowMutation argument. R23_ProbeSectionDimensions calls neither
' and contains no AddDimension2, SetFitValues or SetValues2 call.
'
' R23-802. Reconciliation runs BEFORE creation, always. The section already
' carries imported dimensions - Phase 0 counted seventeen - and creating a
' second dimension for a requirement that is already satisfied is the defect
' this ordering exists to prevent. Each requirement records six independent
' observations about whatever it matched: source dimension identity,
' attached geometry, semantic role, nominal, type, and tolerance.
'
' R23-804. The imported 47 and 40 records are live-proven
' swDimensionType_e.swDiameterDimension = 6, read back from the fixture on
' 2026-07-31 as D1@Sketch4 and D1@Sketch6. swDiametricLinearDimension = 15
' is recorded when seen and never required; requiring it would reject the
' real drawing.
'
' R23-806 and R23-807. The 47 H7 tolerance does NOT exist in the model.
' Phase 0's corrected probe read the part-source dimension directly and got
' toleranceType=0, fitType=-1 and empty fit strings. The user selected
' controlled target-spec authority, so the fit is applied to a real
' associative dimension and its provenance is recorded as reference
' authority - never as model data. A free-text label is never substituted
' for a dimension that failed: a note is not a dimension, does not update
' with the geometry, and cannot be inspected as one.

' MCP corpus values for swDimensionType_e; verify in the SW2025 Object
' Browser before acceptance. Type 6 is additionally live-proven on the
' fixture (see the module header).
Private Const DIM_TYPE_ORDINATE_BASE As Long = 1
Private Const DIM_TYPE_LINEAR As Long = 2
Private Const DIM_TYPE_DIAMETER As Long = 6
Private Const DIM_TYPE_HOR_ORDINATE As Long = 7
Private Const DIM_TYPE_VERT_ORDINATE As Long = 8
Private Const DIM_TYPE_HOR_LINEAR As Long = 11
Private Const DIM_TYPE_VERT_LINEAR As Long = 12
Private Const DIM_TYPE_DIAMETRIC_LINEAR As Long = 15
Private Const DIM_TYPE_ANGULAR_ORDINATE As Long = 16

' MCP corpus value for swTolType_e.swTolFITWITHTOL; verify in the SW2025
' Object Browser. FITWITHTOL rather than FIT because the reference states
' both the H7 symbol and the +0.025/0.000 deviations, and the drawing has
' to show what the reference states.
Private Const TOL_TYPE_FIT_WITH_TOL As Long = 8

' MCP corpus value for swSetValueInConfiguration_e. A drawing dimension has
' no configuration of its own, and this member exists precisely to say so.
Private Const SET_VALUE_NO_CONFIGURATION As Long = -1

' MCP corpus value for swDrawingViewTypes_e.swDrawingSectionView.
Private Const VIEW_TYPE_SECTION As Long = 2

' Two nominals are the same when they agree to a micrometre. Every
' requirement below is separated from every other by at least 5.5 mm, so
' this is nowhere near a discriminating threshold - it exists to absorb
' floating-point representation, not to widen a match.
Public Const NOMINAL_TOLERANCE_M As Double = 0.000001

' R23-801 requirement keys.
Public Const REQ_OVERALL_THICKNESS As String = "OVERALL_THICKNESS_18"
Public Const REQ_BORE_STEP_DEPTH As String = "BORE_STEP_DEPTH_12"
Public Const REQ_INNER_BORE As String = "INNER_BORE_D40"
Public Const REQ_FIT_BORE As String = "FIT_BORE_D47_H7"
Public Const REQ_LOWER_WALL_STEP As String = "LOWER_WALL_STEP_11_5"
Public Const REQ_LONG_VERTICAL As String = "LONG_VERTICAL_REF_173_6"
Public Const REQ_LOWER_VERTICAL As String = "LOWER_VERTICAL_REF_104_8"

' R23-808 lanes. A lane is a NAME, not a coordinate. Turning a lane into a
' page position needs the full annotation envelope of the finished section,
' which is Phase 9's job; deciding it here would repeat the Phase 7 mistake
' of letting a module that cannot see the layout choose placement.
Public Const LANE_ABOVE As String = "SectionAbove"
Public Const LANE_BELOW As String = "SectionBelow"
Public Const LANE_BORE_SIDE_A As String = "BoreSideA"
Public Const LANE_BORE_SIDE_B As String = "BoreSideB"
Public Const LANE_EXTERIOR_VERTICAL_OUTER As String = _
    "ExteriorVerticalOuter"
Public Const LANE_EXTERIOR_VERTICAL_INNER As String = _
    "ExteriorVerticalInner"

' R23-806. The single place the reference fit is stated.
Public Const REFERENCE_HOLE_FIT As String = "H7"
Public Const REFERENCE_FIT_MIN_M As Double = 0#
Public Const REFERENCE_FIT_MAX_M As Double = 0.000025
Public Const REFERENCE_FIT_AUTHORITY As String = _
    "TargetSpecReferenceAuthority.NotModelData"

' R23-810. The legacy free-text bore callout, still emitted by
' Module7_TitleBlockEngine. Detected here, removed when the pipeline
' switches over.
Private Const LEGACY_BORE_CALLOUT_MARK As String = "47 H7"

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

' Dimension names carry "@" freely and can carry the delimiters this
' evidence format uses. An unsanitized name splits one field into several
' and silently corrupts every field after it on the line.
Public Function SectionToken(ByVal value As String) As String
    Dim result As String
    result = value
    result = Replace$(result, "|", "/")
    result = Replace$(result, "=", ":")
    result = Replace$(result, vbCr, " ")
    result = Replace$(result, vbLf, " ")
    If Len(Trim$(result)) = 0 Then result = "Empty"
    SectionToken = result
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

' R23-801. The seven P-0251 section requirements, in the order the reference
' states them.
Public Function BuildSectionRequirements() As Collection
    Dim result As Collection
    Set result = New Collection
    Set BuildSectionRequirements = result

    ' Linear requirements accept generic, horizontal and vertical linear
    ' records. Which of the three SOLIDWORKS assigns depends on how the
    ' section is oriented on the sheet, and the nominal already separates
    ' these seven from each other by millimetres.
    Dim linearTypes As String
    linearTypes = CStr(DIM_TYPE_LINEAR) & ";" & _
        CStr(DIM_TYPE_HOR_LINEAR) & ";" & CStr(DIM_TYPE_VERT_LINEAR)

    ' R23-804. Diameters accept the live-proven type 6; 15 is accepted too
    ' where the build happens to produce it, but is never required.
    Dim diameterTypes As String
    diameterTypes = CStr(DIM_TYPE_DIAMETER) & ";" & _
        CStr(DIM_TYPE_DIAMETRIC_LINEAR)

    result.Add MakeRequirement(REQ_OVERALL_THICKNESS, _
        "OverallThickness", 0.018, linearTypes, LANE_ABOVE)
    result.Add MakeRequirement(REQ_BORE_STEP_DEPTH, _
        "BoreStepDepth", 0.012, linearTypes, LANE_ABOVE)
    result.Add MakeRequirement(REQ_LOWER_WALL_STEP, _
        "LowerWallStep", 0.0115, linearTypes, LANE_BELOW)
    result.Add MakeRequirement(REQ_INNER_BORE, _
        "InnerBoreDiameter", 0.04, diameterTypes, LANE_BORE_SIDE_A)

    Dim fitBore As CSectionRequirement
    Set fitBore = MakeRequirement(REQ_FIT_BORE, _
        "FitBoreDiameter", 0.047, diameterTypes, LANE_BORE_SIDE_B)

    ' R23-806. The only requirement carrying a tolerance, and its authority
    ' is stated in the same breath as the numbers.
    fitBore.ToleranceRequired = True
    fitBore.RequiredHoleFit = REFERENCE_HOLE_FIT
    fitBore.RequiredToleranceMinM = REFERENCE_FIT_MIN_M
    fitBore.RequiredToleranceMaxM = REFERENCE_FIT_MAX_M
    fitBore.ToleranceAuthority = REFERENCE_FIT_AUTHORITY
    result.Add fitBore

    result.Add MakeRequirement(REQ_LONG_VERTICAL, _
        "LongVerticalReference", 0.1736, linearTypes, _
        LANE_EXTERIOR_VERTICAL_OUTER)
    result.Add MakeRequirement(REQ_LOWER_VERTICAL, _
        "LowerVerticalReference", 0.1048, linearTypes, _
        LANE_EXTERIOR_VERTICAL_INNER)
End Function

Private Function MakeRequirement( _
    ByVal key As String, _
    ByVal role As String, _
    ByVal nominalM As Double, _
    ByVal acceptedTypes As String, _
    ByVal lane As String) As CSectionRequirement

    Dim requirement As CSectionRequirement
    Set requirement = New CSectionRequirement

    requirement.Key = key
    requirement.Role = role
    requirement.NominalM = nominalM
    requirement.AcceptedTypes = acceptedTypes
    requirement.Lane = lane
    requirement.ToleranceRequired = False
    requirement.ToleranceAuthority = "None"

    Set MakeRequirement = requirement
End Function

Public Function IsSectionView( _
    ByRef swView As SldWorks.View) As Boolean

    On Error GoTo Failed
    If swView Is Nothing Then Exit Function
    IsSectionView = (swView.Type = VIEW_TYPE_SECTION)
    Exit Function

Failed:
    IsSectionView = False
End Function

' R23-809. A section view is excluded from generic auto-arrangement. Its
' dimensions sit in named lanes chosen for the cut, and a generic arranger
' that knows nothing about those lanes will move them out of the lanes it
' cannot see.
'
' This is the predicate; Module9_LayoutEngine does not consult it yet,
' because the R23 modules are not wired into the production pipeline. Same
' deferral as R23-609 and R23-704.
Public Function IsExcludedFromGenericArrangement( _
    ByRef swView As SldWorks.View) As Boolean

    IsExcludedFromGenericArrangement = IsSectionView(swView)
End Function

Public Function CollectSectionViews( _
    ByRef views As Variant) As Collection

    Dim result As Collection
    Set result = New Collection
    Set CollectSectionViews = result

    On Error GoTo Failed

    If IsEmpty(views) Or Not IsArray(views) Then Exit Function

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim candidate As SldWorks.View
        Set candidate = views(i)
        If IsSectionView(candidate) Then result.Add candidate
    Next i

    Exit Function

Failed:
    Set CollectSectionViews = result
End Function

' R23-804 and R23-805. Every display dimension owned by ONE view, read back
' field by field.
'
' IView.GetDisplayDimensions is view-scoped. Its obsolete predecessor
' GetFirstDisplayDimension5 walks the whole sheet by its own Remarks, which
' would attribute another view's dimensions to the section.
Public Function InventorySectionDimensions( _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence) As Collection

    Dim result As Collection
    Set result = New Collection
    Set InventorySectionDimensions = result

    On Error GoTo Failed

    If swView Is Nothing Then Exit Function

    Dim dimensions As Variant
    dimensions = swView.GetDisplayDimensions

    If IsEmpty(dimensions) Or Not IsArray(dimensions) Then Exit Function

    Dim viewName As String
    viewName = SafeViewName(swView)

    Dim i As Long
    For i = LBound(dimensions) To UBound(dimensions)
        Dim displayDimension As SldWorks.DisplayDimension
        Set displayDimension = dimensions(i)

        If displayDimension Is Nothing Then GoTo ContinueDimension

        result.Add displayDimension

        ' Every field is read fresh for this dimension. VBA block-scoped
        ' locals live for the whole procedure, so a value left over from
        ' the previous iteration is reported against this one - the exact
        ' labelling defect that corrupted the Phase 0 section inventory.
        Dim dimension As SldWorks.Dimension
        Set dimension = displayDimension.GetDimension2(0)

        Dim nominalM As Double
        Dim nominalAvailable As Boolean
        nominalM = 0#
        nominalAvailable = TryReadNominal(dimension, nominalM)

        Dim toleranceType As Long
        Dim fitType As Long
        Dim holeFit As String
        Dim shaftFit As String
        Dim minimumM As Double
        Dim maximumM As Double
        Dim minimumStatus As Long
        Dim maximumStatus As Long
        Dim toleranceProof As String

        toleranceProof = ReadDimensionTolerance(dimension, _
            toleranceType, fitType, holeFit, shaftFit, _
            minimumM, maximumM, minimumStatus, maximumStatus)

        Dim attachedCount As Long
        Dim attachedTypes As String
        attachedCount = 0
        attachedTypes = ReadAttachment(displayDimension, attachedCount)

        EmitInfo evidence, "SECTION_DIM|view=" & SectionToken(viewName) & _
            "|index=" & CStr(result.Count) & _
            "|name=" & SectionToken( _
                SafeDisplayDimensionName(displayDimension)) & _
            "|fullName=" & SectionToken( _
                SafeDimensionFullName(dimension)) & _
            "|type2=" & CStr(SafeTypeCode(displayDimension)) & _
            "|nominalAvailable=" & CStr(nominalAvailable) & _
            "|nominalM=" & FormatMetres(nominalM) & _
            "|attachedEntities=" & CStr(attachedCount) & _
            "|attachedTypes=" & attachedTypes & _
            "|" & toleranceProof

ContinueDimension:
    Next i

    Exit Function

Failed:
    EmitFailure evidence, "SECTION_DIM_INVENTORY_ERROR|view=" & _
        SectionToken(SafeViewName(swView)) & _
        "|error=" & CStr(Err.Number)
    Set InventorySectionDimensions = result
End Function

Private Function SafeTypeCode( _
    ByRef displayDimension As SldWorks.DisplayDimension) As Long

    On Error GoTo Failed
    SafeTypeCode = -1
    SafeTypeCode = displayDimension.Type2
    Exit Function

Failed:
    SafeTypeCode = -1
End Function

Private Function SafeDisplayDimensionName( _
    ByRef displayDimension As SldWorks.DisplayDimension) As String

    On Error GoTo Failed
    SafeDisplayDimensionName = "Unavailable"
    SafeDisplayDimensionName = displayDimension.GetNameForSelection
    Exit Function

Failed:
    SafeDisplayDimensionName = "Unavailable:" & CStr(Err.Number)
End Function

Private Function SafeDimensionFullName( _
    ByRef dimension As SldWorks.Dimension) As String

    On Error GoTo Failed
    If dimension Is Nothing Then
        SafeDimensionFullName = "NoDimension"
        Exit Function
    End If

    SafeDimensionFullName = dimension.FullName
    Exit Function

Failed:
    SafeDimensionFullName = "Unavailable:" & CStr(Err.Number)
End Function

' IDimension.GetSystemValue3 returns the value in system units - metres.
' Config_names matters only for swSpecifyConfiguration, which this is not.
Private Function TryReadNominal( _
    ByRef dimension As SldWorks.Dimension, _
    ByRef nominalM As Double) As Boolean

    If dimension Is Nothing Then Exit Function
    On Error GoTo Failed

    Dim configurationNames As Variant
    configurationNames = Empty

    Dim values As Variant
    values = dimension.GetSystemValue3( _
        swThisConfiguration, configurationNames)

    If IsArray(values) Then
        If (UBound(values) - LBound(values) + 1) < 1 Then Exit Function
        nominalM = CDbl(values(LBound(values)))
    ElseIf Not IsEmpty(values) And Not IsNull(values) Then
        nominalM = CDbl(values)
    Else
        Exit Function
    End If

    TryReadNominal = True
    Exit Function

Failed:
    TryReadNominal = False
End Function

' R23-805. Tolerance readback through IDimension.Tolerance.
'
' IDimension.GetToleranceValues, SetToleranceValues, GetToleranceFitValues
' and SetToleranceFitValues are all marked obsolete by the 2025 Help, each
' superseded by an IDimensionTolerance member. This reads the supported
' route only. GetMinValue2 and GetMaxValue2 return a STATUS and hand the
' value back through a ByRef argument, so the status is reported next to
' the value it qualifies - a zero value with a failed status is not a zero
' tolerance.
Private Function ReadDimensionTolerance( _
    ByRef dimension As SldWorks.Dimension, _
    ByRef toleranceType As Long, _
    ByRef fitType As Long, _
    ByRef holeFit As String, _
    ByRef shaftFit As String, _
    ByRef minimumM As Double, _
    ByRef maximumM As Double, _
    ByRef minimumStatus As Long, _
    ByRef maximumStatus As Long) As String

    toleranceType = -1
    fitType = -1
    holeFit = vbNullString
    shaftFit = vbNullString
    minimumM = 0#
    maximumM = 0#
    minimumStatus = -1
    maximumStatus = -1

    If dimension Is Nothing Then
        ReadDimensionTolerance = "tolerance=DimensionNothing"
        Exit Function
    End If

    On Error GoTo Failed

    Dim tolerance As SldWorks.DimensionTolerance
    Set tolerance = dimension.Tolerance

    If tolerance Is Nothing Then
        ReadDimensionTolerance = "tolerance=Nothing"
        Exit Function
    End If

    toleranceType = CLng(tolerance.Type)
    fitType = CLng(tolerance.FitType)
    holeFit = tolerance.GetHoleFitValue
    shaftFit = tolerance.GetShaftFitValue
    minimumStatus = CLng(tolerance.GetMinValue2(minimumM))
    maximumStatus = CLng(tolerance.GetMaxValue2(maximumM))

    ReadDimensionTolerance = "toleranceType=" & CStr(toleranceType) & _
        "|fitType=" & CStr(fitType) & _
        "|holeFit=" & SectionToken(holeFit) & _
        "|shaftFit=" & SectionToken(shaftFit) & _
        "|minimumStatus=" & CStr(minimumStatus) & _
        "|minimumM=" & FormatMetres(minimumM) & _
        "|maximumStatus=" & CStr(maximumStatus) & _
        "|maximumM=" & FormatMetres(maximumM)
    Exit Function

Failed:
    ReadDimensionTolerance = "toleranceReadError=" & CStr(Err.Number)
End Function

' R23-802, attached-geometry half. An associative dimension is attached to
' real drawing entities; a dangling one reports swSelNOTHING in the type
' array with a Nothing in the matching entity slot. Counting attachments is
' how "associative" is proved rather than assumed.
Private Function ReadAttachment( _
    ByRef displayDimension As SldWorks.DisplayDimension, _
    ByRef attachedCount As Long) As String

    attachedCount = 0

    On Error GoTo Failed

    Dim annotation As SldWorks.Annotation
    Set annotation = displayDimension.GetAnnotation

    If annotation Is Nothing Then
        ReadAttachment = "NoAnnotation"
        Exit Function
    End If

    Dim types As Variant
    types = annotation.GetAttachedEntityTypes

    If IsEmpty(types) Or Not IsArray(types) Then
        ReadAttachment = "NoAttachedEntities"
        Exit Function
    End If

    Dim summary As String

    Dim i As Long
    For i = LBound(types) To UBound(types)
        attachedCount = attachedCount + 1
        If Len(summary) = 0 Then
            summary = CStr(CLng(types(i)))
        Else
            summary = summary & "/" & CStr(CLng(types(i)))
        End If
    Next i

    If Len(summary) = 0 Then summary = "None"
    ReadAttachment = summary
    Exit Function

Failed:
    ReadAttachment = "AttachmentReadError:" & CStr(Err.Number)
End Function

' R23-802. Reconciles the imported dimensions against the requirements,
' BEFORE anything is created.
'
' A requirement matches a dimension on nominal and accepted type together.
' The other four observations - source dimension identity, attached
' geometry, semantic role, tolerance - are recorded on the requirement so
' the match can be audited, and the tolerance additionally decides whether
' a matched requirement is fully satisfied.
'
' Every match is counted, not just the first, because R23-811 has to be
' able to fail on a duplicate rather than quietly dimension the same thing
' twice.
Public Sub ReconcileSectionDimensions( _
    ByRef requirements As Collection, _
    ByRef dimensions As Collection, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed

    Dim r As Long
    For r = 1 To requirements.Count
        Dim requirement As CSectionRequirement
        Set requirement = requirements(r)

        requirement.MatchCount = 0
        requirement.Matched = False
        requirement.Origin = "Missing"
        requirement.Failures = vbNullString

        Dim d As Long
        For d = 1 To dimensions.Count
            Dim displayDimension As SldWorks.DisplayDimension
            Set displayDimension = dimensions(d)

            Dim dimension As SldWorks.Dimension
            Set dimension = displayDimension.GetDimension2(0)

            Dim nominalM As Double
            Dim nominalAvailable As Boolean
            nominalM = 0#
            nominalAvailable = TryReadNominal(dimension, nominalM)

            If Not nominalAvailable Then GoTo ContinueDimension

            Dim typeCode As Long
            typeCode = SafeTypeCode(displayDimension)

            If Not requirement.AcceptsTypeCode(typeCode) Then
                GoTo ContinueDimension
            End If

            If Abs(nominalM - requirement.NominalM) > _
                NOMINAL_TOLERANCE_M Then

                GoTo ContinueDimension
            End If

            requirement.MatchCount = requirement.MatchCount + 1

            ' The first match populates the observation; a second one is
            ' counted and reported rather than overwriting it.
            If requirement.MatchCount = 1 Then
                RecordObservation requirement, displayDimension, _
                    dimension, typeCode, nominalAvailable, nominalM
            End If

ContinueDimension:
        Next d

        If requirement.MatchCount > 0 Then
            requirement.Matched = True
            requirement.Origin = "ImportedSectionDimension"
        Else
            requirement.Failures = AppendFailure( _
                requirement.Failures, "NoImportedDimension")
        End If

        EvaluateTolerance requirement

        If requirement.MatchCount > 1 Then
            requirement.Failures = AppendFailure( _
                requirement.Failures, _
                "DuplicateDimensions:" & CStr(requirement.MatchCount))
        End If

        If Len(requirement.Failures) = 0 Then
            requirement.Failures = "None"
        End If

        EmitInfo evidence, "SECTION_REQUIREMENT|" & requirement.Summary()
    Next r

    Exit Sub

Failed:
    EmitFailure evidence, "SECTION_RECONCILE_ERROR|error=" & _
        CStr(Err.Number)
End Sub

Private Sub RecordObservation( _
    ByRef requirement As CSectionRequirement, _
    ByRef displayDimension As SldWorks.DisplayDimension, _
    ByRef dimension As SldWorks.Dimension, _
    ByVal typeCode As Long, _
    ByVal nominalAvailable As Boolean, _
    ByVal nominalM As Double)

    requirement.MatchedName = _
        SectionToken(SafeDisplayDimensionName(displayDimension))
    requirement.MatchedFullName = _
        SectionToken(SafeDimensionFullName(dimension))
    requirement.MatchedTypeCode = typeCode
    requirement.MatchedNominalAvailable = nominalAvailable
    requirement.MatchedNominalM = nominalM

    Dim toleranceType As Long
    Dim fitType As Long
    Dim holeFit As String
    Dim shaftFit As String
    Dim minimumM As Double
    Dim maximumM As Double
    Dim minimumStatus As Long
    Dim maximumStatus As Long

    requirement.ToleranceProof = ReadDimensionTolerance(dimension, _
        toleranceType, fitType, holeFit, shaftFit, _
        minimumM, maximumM, minimumStatus, maximumStatus)

    requirement.MatchedToleranceType = toleranceType
    requirement.MatchedFitType = fitType
    requirement.MatchedHoleFit = SectionToken(holeFit)
    requirement.MatchedShaftFit = SectionToken(shaftFit)
    requirement.MatchedToleranceMinM = minimumM
    requirement.MatchedToleranceMaxM = maximumM
    requirement.MatchedToleranceMinStatus = minimumStatus
    requirement.MatchedToleranceMaxStatus = maximumStatus

    Dim attachedCount As Long
    attachedCount = 0
    requirement.AttachedEntityTypes = _
        ReadAttachment(displayDimension, attachedCount)
    requirement.AttachedEntityCount = attachedCount

    If attachedCount = 0 Then
        requirement.Failures = AppendFailure( _
            requirement.Failures, "DimensionNotAttached")
    End If
End Sub

' R23-806. Decides whether a matched requirement's tolerance is already
' satisfied, and states where the answer came from.
'
' A requirement with no tolerance is satisfied by having none: adding one
' would be inventing a manufacturing constraint, which is exactly what the
' standing tolerance policy forbids until the designer's rule is known.
Private Sub EvaluateTolerance( _
    ByRef requirement As CSectionRequirement)

    If Not requirement.ToleranceRequired Then
        requirement.ToleranceSatisfied = True
        requirement.ToleranceProvenance = "NoToleranceRequired"
        Exit Sub
    End If

    If Not requirement.Matched Then
        requirement.ToleranceSatisfied = False
        requirement.ToleranceProvenance = "NoDimensionToCarryTolerance"
        Exit Sub
    End If

    Dim fitPresent As Boolean
    fitPresent = (InStr(1, requirement.MatchedHoleFit, _
        requirement.RequiredHoleFit, vbTextCompare) > 0)

    Dim deviationsPresent As Boolean
    deviationsPresent = _
        (Abs(requirement.MatchedToleranceMinM - _
            requirement.RequiredToleranceMinM) <= NOMINAL_TOLERANCE_M) And _
        (Abs(requirement.MatchedToleranceMaxM - _
            requirement.RequiredToleranceMaxM) <= NOMINAL_TOLERANCE_M) And _
        (requirement.MatchedToleranceMinStatus = 0) And _
        (requirement.MatchedToleranceMaxStatus = 0)

    If fitPresent And deviationsPresent Then
        requirement.ToleranceSatisfied = True

        ' Present on the drawing does not mean it came from the model.
        ' Phase 0 proved the part-source dimension carries no tolerance at
        ' all, so whatever is on the drawing was authored on the drawing.
        requirement.ToleranceProvenance = _
            "PresentOnDrawing." & REFERENCE_FIT_AUTHORITY
        Exit Sub
    End If

    requirement.ToleranceSatisfied = False
    requirement.ToleranceProvenance = "AbsentRequiresReferenceAuthority"
    requirement.Failures = AppendFailure(requirement.Failures, _
        "ToleranceMissing:fit=" & CStr(fitPresent) & _
        ".deviations=" & CStr(deviationsPresent))
End Sub

' R23-811. Exactly one dimension per requirement key, and no ordinate
' anywhere in the section.
'
' Ordinates belong to the face views, where Phase 5 places them against a
' proved datum. An ordinate in a section shares no datum with those groups
' and reads as a coordinate from an origin the section does not have.
Public Function VerifySectionDimensions( _
    ByRef requirements As Collection, _
    ByRef dimensions As Collection) As String

    On Error GoTo Failed

    Dim failures As String
    Dim satisfied As Long
    Dim missing As Long
    Dim duplicated As Long

    Dim r As Long
    For r = 1 To requirements.Count
        Dim requirement As CSectionRequirement
        Set requirement = requirements(r)

        If requirement.MatchCount = 0 Then
            missing = missing + 1
            failures = AppendFailure(failures, _
                "Missing:" & requirement.Key)
        ElseIf requirement.MatchCount > 1 Then
            duplicated = duplicated + 1
            failures = AppendFailure(failures, _
                "Duplicate:" & requirement.Key)
        ElseIf Not requirement.ToleranceSatisfied Then
            failures = AppendFailure(failures, _
                "ToleranceUnsatisfied:" & requirement.Key)
        ElseIf requirement.AttachedEntityCount = 0 Then
            failures = AppendFailure(failures, _
                "Unattached:" & requirement.Key)
        Else
            satisfied = satisfied + 1
        End If
    Next r

    Dim ordinateCount As Long
    ordinateCount = CountSectionOrdinates(dimensions)

    If ordinateCount > 0 Then
        failures = AppendFailure(failures, _
            "SectionOrdinatePresent:" & CStr(ordinateCount))
    End If

    If Len(failures) = 0 Then failures = "None"

    VerifySectionDimensions = "requirements=" & _
        CStr(requirements.Count) & _
        "|satisfied=" & CStr(satisfied) & _
        "|missing=" & CStr(missing) & _
        "|duplicated=" & CStr(duplicated) & _
        "|sectionDimensions=" & CStr(dimensions.Count) & _
        "|sectionOrdinates=" & CStr(ordinateCount) & _
        "|requirementFailures=" & failures
    Exit Function

Failed:
    VerifySectionDimensions = _
        "requirementFailures=Error:" & CStr(Err.Number)
End Function

Public Function CountSectionOrdinates( _
    ByRef dimensions As Collection) As Long

    On Error GoTo Failed

    Dim total As Long

    Dim i As Long
    For i = 1 To dimensions.Count
        Dim displayDimension As SldWorks.DisplayDimension
        Set displayDimension = dimensions(i)

        Dim typeCode As Long
        typeCode = SafeTypeCode(displayDimension)

        If typeCode = DIM_TYPE_ORDINATE_BASE Or _
           typeCode = DIM_TYPE_HOR_ORDINATE Or _
           typeCode = DIM_TYPE_VERT_ORDINATE Or _
           typeCode = DIM_TYPE_ANGULAR_ORDINATE Then

            total = total + 1
        End If
    Next i

    CountSectionOrdinates = total
    Exit Function

Failed:
    CountSectionOrdinates = total
End Function

' R23-810. Finds the legacy free-text bore callout in the section view.
'
' Detection only. Removing it before the real dimensions exist would leave
' the bore undefined, and Module7_TitleBlockEngine still authors it on the
' reachable production path.
Public Function DetectLegacyBoreCallout( _
    ByRef swView As SldWorks.View) As String

    On Error GoTo Failed

    If swView Is Nothing Then
        DetectLegacyBoreCallout = "legacyBoreCallout=NoView"
        Exit Function
    End If

    Dim notes As Variant
    notes = swView.GetNotes

    If IsEmpty(notes) Or Not IsArray(notes) Then
        DetectLegacyBoreCallout = "legacyBoreCallout=NoNotes"
        Exit Function
    End If

    Dim found As Long

    Dim i As Long
    For i = LBound(notes) To UBound(notes)
        Dim note As SldWorks.Note
        Set note = notes(i)
        If note Is Nothing Then GoTo ContinueNote

        If InStr(1, note.GetText, LEGACY_BORE_CALLOUT_MARK, _
            vbTextCompare) > 0 Then

            found = found + 1
        End If

ContinueNote:
    Next i

    DetectLegacyBoreCallout = "legacyBoreCallout=" & CStr(found) & _
        "|removalBlockedBy=Module7_TitleBlockEngine.PipelineNotSwitched"
    Exit Function

Failed:
    DetectLegacyBoreCallout = "legacyBoreCallout=Error:" & CStr(Err.Number)
End Function

' R23-806. MUTATES THE DRAWING.
'
' Applies the approved reference fit to an existing associative dimension.
' Refuses without allowMutation. The provenance it writes back says
' plainly that the value is target-spec reference authority and not model
' data, because the part-source dimension carries no tolerance and QA has
' to be able to state that.
Public Function ApplyReferenceFit( _
    ByRef dimension As SldWorks.Dimension, _
    ByRef requirement As CSectionRequirement, _
    ByVal allowMutation As Boolean, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If Not allowMutation Then
        EmitWarning evidence, "SECTION_FIT_REFUSED" & _
            "|reason=MutationNotAuthorized|key=" & requirement.Key
        Exit Function
    End If

    If dimension Is Nothing Then
        EmitFailure evidence, "SECTION_FIT_REFUSED" & _
            "|reason=NoDimension|key=" & requirement.Key
        Exit Function
    End If

    If Not requirement.ToleranceRequired Then
        EmitFailure evidence, "SECTION_FIT_REFUSED" & _
            "|reason=RequirementCarriesNoTolerance" & _
            "|key=" & requirement.Key
        Exit Function
    End If

    Dim tolerance As SldWorks.DimensionTolerance
    Set tolerance = dimension.Tolerance

    If tolerance Is Nothing Then
        EmitFailure evidence, "SECTION_FIT_REFUSED" & _
            "|reason=NoToleranceObject|key=" & requirement.Key
        Exit Function
    End If

    ' Order matters. SetValues2 refuses while the type is swTolNONE by its
    ' own Remarks, and FitType/SetFitValues are only available for the fit
    ' types, so the type is set first.
    tolerance.Type = TOL_TYPE_FIT_WITH_TOL

    Dim fitApplied As Boolean
    fitApplied = Module11_GeometryIdentity.NormalizeSwBoolean( _
        tolerance.SetFitValues(requirement.RequiredHoleFit, ""))

    Dim valuesApplied As Boolean
    valuesApplied = Module11_GeometryIdentity.NormalizeSwBoolean( _
        tolerance.SetValues2( _
            requirement.RequiredToleranceMinM, _
            requirement.RequiredToleranceMaxM, _
            SET_VALUE_NO_CONFIGURATION, Empty))

    If Not evidence Is Nothing Then
        evidence.RecordSolidWorksMutation _
            "SectionReferenceFit:" & requirement.Key
    End If

    ' Read back rather than trusting the two return values.
    Dim toleranceType As Long
    Dim fitType As Long
    Dim holeFit As String
    Dim shaftFit As String
    Dim minimumM As Double
    Dim maximumM As Double
    Dim minimumStatus As Long
    Dim maximumStatus As Long

    requirement.ToleranceProof = ReadDimensionTolerance(dimension, _
        toleranceType, fitType, holeFit, shaftFit, _
        minimumM, maximumM, minimumStatus, maximumStatus)

    requirement.MatchedToleranceType = toleranceType
    requirement.MatchedFitType = fitType
    requirement.MatchedHoleFit = SectionToken(holeFit)
    requirement.MatchedShaftFit = SectionToken(shaftFit)
    requirement.MatchedToleranceMinM = minimumM
    requirement.MatchedToleranceMaxM = maximumM
    requirement.MatchedToleranceMinStatus = minimumStatus
    requirement.MatchedToleranceMaxStatus = maximumStatus

    Dim proven As Boolean
    proven = fitApplied And valuesApplied And _
        (InStr(1, holeFit, requirement.RequiredHoleFit, _
            vbTextCompare) > 0) And _
        (Abs(maximumM - requirement.RequiredToleranceMaxM) <= _
            NOMINAL_TOLERANCE_M)

    requirement.ToleranceSatisfied = proven
    requirement.ToleranceProvenance = _
        "AppliedFrom." & REFERENCE_FIT_AUTHORITY

    ApplyReferenceFit = proven

    EmitInfo evidence, "SECTION_FIT_APPLIED|key=" & requirement.Key & _
        "|fitApplied=" & CStr(fitApplied) & _
        "|valuesApplied=" & CStr(valuesApplied) & _
        "|proven=" & CStr(proven) & _
        "|provenance=" & requirement.ToleranceProvenance & _
        "|" & requirement.ToleranceProof
    Exit Function

Failed:
    EmitFailure evidence, "SECTION_FIT_ERROR|key=" & requirement.Key & _
        "|error=" & CStr(Err.Number) & "|description=" & Err.Description
    ApplyReferenceFit = False
End Function

' R23-803 and R23-807. MUTATES THE DRAWING.
'
' Creates the one associative dimension a requirement needs, and only when
' reconciliation proved nothing already satisfies it.
'
' The entities must already be selected. IModelDoc2.AddDimension2's Remarks
' are explicit that selections are made by LOCATION and never by name -
' passing a name makes the dimensioning routines pick an endpoint at random
' - so choosing and selecting them is the caller's responsibility and is
' verified here rather than assumed.
'
' textX and textY are the dimension text position on the sheet. Like the
' section view's placement in Phase 7, they are caller arguments: the lane
' on the requirement names WHERE the dimension belongs, and turning a lane
' into coordinates needs the finished section's annotation envelope.
'
' R23-807. When this fails, it reports the failure. It never writes a note
' saying "47 H7" instead: a note is not a dimension, does not move with the
' geometry, and cannot be inspected as one.
Public Function CreateSectionDimension( _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef requirement As CSectionRequirement, _
    ByVal textX As Double, _
    ByVal textY As Double, _
    ByVal allowMutation As Boolean, _
    ByRef evidence As CRunEvidence) As SldWorks.DisplayDimension

    On Error GoTo Failed

    If Not allowMutation Then
        EmitWarning evidence, "SECTION_DIM_CREATE_REFUSED" & _
            "|reason=MutationNotAuthorized|key=" & requirement.Key
        Exit Function
    End If

    If swDraw Is Nothing Or requirement Is Nothing Then Exit Function

    ' R23-803. Reconciliation first, creation only for what is missing.
    If requirement.Matched Then
        EmitWarning evidence, "SECTION_DIM_CREATE_SKIPPED" & _
            "|reason=AlreadySatisfiedByImportedDimension" & _
            "|key=" & requirement.Key & _
            "|existing=" & requirement.MatchedFullName
        Exit Function
    End If

    Dim selectionMgr As SldWorks.SelectionMgr
    Set selectionMgr = swDraw.SelectionManager

    Dim selectedCount As Long
    selectedCount = selectionMgr.GetSelectedObjectCount2(-1)

    If selectedCount < 1 Then
        EmitFailure evidence, "SECTION_DIM_CREATE_REFUSED" & _
            "|reason=NoEntitiesSelected|key=" & requirement.Key & _
            "|policy=NoFreeTextSubstitute"
        Exit Function
    End If

    Dim created As SldWorks.DisplayDimension
    Set created = swDraw.AddDimension2(textX, textY, 0#)

    If Not evidence Is Nothing Then
        evidence.RecordSolidWorksMutation _
            "SectionDimension:" & requirement.Key
    End If

    If created Is Nothing Then
        EmitFailure evidence, "SECTION_DIM_CREATE_FAILED" & _
            "|reason=AddDimension2ReturnedNothing" & _
            "|key=" & requirement.Key & _
            "|selectedCount=" & CStr(selectedCount) & _
            "|policy=NoFreeTextSubstitute"
        Exit Function
    End If

    Dim dimension As SldWorks.Dimension
    Set dimension = created.GetDimension2(0)

    Dim typeCode As Long
    typeCode = SafeTypeCode(created)

    Dim nominalM As Double
    Dim nominalAvailable As Boolean
    nominalM = 0#
    nominalAvailable = TryReadNominal(dimension, nominalM)

    RecordObservation requirement, created, dimension, typeCode, _
        nominalAvailable, nominalM

    Dim nominalCorrect As Boolean
    nominalCorrect = nominalAvailable And _
        (Abs(nominalM - requirement.NominalM) <= NOMINAL_TOLERANCE_M)

    Dim typeAccepted As Boolean
    typeAccepted = requirement.AcceptsTypeCode(typeCode)

    If Not nominalCorrect Or Not typeAccepted Then
        ' The dimension exists but measures something other than the
        ' requirement. Reporting it as satisfied would be worse than
        ' reporting nothing, because a wrong dimension looks finished.
        requirement.Matched = False
        requirement.Origin = "CreatedButRejected"
        requirement.Failures = AppendFailure(requirement.Failures, _
            "CreatedDimensionMismatch:nominal=" & _
            CStr(nominalCorrect) & ".type=" & CStr(typeAccepted))

        EmitFailure evidence, "SECTION_DIM_CREATE_REJECTED" & _
            "|key=" & requirement.Key & _
            "|expectedNominalM=" & FormatMetres(requirement.NominalM) & _
            "|actualNominalM=" & FormatMetres(nominalM) & _
            "|type2=" & CStr(typeCode) & _
            "|acceptedTypes=" & requirement.AcceptedTypes & _
            "|policy=NoFreeTextSubstitute"
        Set CreateSectionDimension = created
        Exit Function
    End If

    requirement.Matched = True
    requirement.MatchCount = 1
    requirement.Origin = "CreatedAssociativeDimension"

    If requirement.ToleranceRequired Then
        ApplyReferenceFit dimension, requirement, allowMutation, evidence
    Else
        EvaluateTolerance requirement
    End If

    Set CreateSectionDimension = created

    EmitInfo evidence, "SECTION_DIM_CREATED|key=" & requirement.Key & _
        "|lane=" & requirement.Lane & _
        "|" & requirement.ObservationSummary()
    Exit Function

Failed:
    EmitFailure evidence, "SECTION_DIM_CREATE_ERROR|key=" & _
        requirement.Key & "|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description & _
        "|policy=NoFreeTextSubstitute"
    Set CreateSectionDimension = Nothing
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

' R23-800 to R23-811 evidence entry point. STRICTLY READ-ONLY: no
' AddDimension2, no SetFitValues, no SetValues2, no CreateSectionDimension
' and no ApplyReferenceFit call appears below.
Public Sub R23_ProbeSectionDimensions()
    On Error GoTo Failed

    mEmitDiagnostics = False

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Debug.Print "R23_SECTIONDIM_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swDraw As SldWorks.ModelDoc2
    Set swDraw = swApp.ActiveDoc

    If swDraw Is Nothing Then
        Debug.Print "R23_SECTIONDIM_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swDraw.GetType <> swDocDRAWING Then
        Debug.Print "R23_SECTIONDIM_FATAL" & _
            "|reason=ActiveDocumentNotDrawing"
        Exit Sub
    End If

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet

    Dim views As Variant
    views = swSheet.GetViews

    If IsEmpty(views) Or Not IsArray(views) Then
        Debug.Print "R23_SECTIONDIM_FATAL|reason=NoViewsOnSheet"
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
        Debug.Print "R23_SECTIONDIM_FATAL|reason=NoReferencedDocument"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Debug.Print "R23_SECTIONDIM_FATAL|reason=UnauthorizedFixture" & _
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

    Debug.Print "R23_SECTIONDIM_BEGIN|drawing=" & swDraw.GetPathName & _
        "|part=" & partPath & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath) & _
        "|mode=ReadOnly|creations=0|mutations=0"

    Dim sectionViews As Collection
    Set sectionViews = CollectSectionViews(views)

    If sectionViews.Count = 0 Then
        Debug.Print "R23_SECTIONDIM_FATAL|reason=NoSectionViewOnSheet"
        Exit Sub
    End If

    Dim bestVerdict As String
    bestVerdict = "requirementFailures=NoSectionViewAttempted"

    Dim satisfiedViews As Long

    For i = 1 To sectionViews.Count
        Dim swView As SldWorks.View
        Set swView = sectionViews(i)

        Dim viewName As String
        viewName = SafeViewName(swView)

        Debug.Print "QA INFO: SECTION_VIEW|name=" & _
            SectionToken(viewName) & _
            "|type=" & CStr(swView.Type) & _
            "|excludedFromGenericArrangement=" & _
            CStr(IsExcludedFromGenericArrangement(swView))

        Dim dimensions As Collection
        Set dimensions = InventorySectionDimensions(swView, evidence)

        Dim requirements As Collection
        Set requirements = BuildSectionRequirements()

        ReconcileSectionDimensions requirements, dimensions, evidence

        Dim r As Long
        For r = 1 To requirements.Count
            Dim requirement As CSectionRequirement
            Set requirement = requirements(r)
            Debug.Print "QA INFO: SECTION_REQUIREMENT|view=" & _
                SectionToken(viewName) & "|" & requirement.Summary()
        Next r

        Dim verdict As String
        verdict = VerifySectionDimensions(requirements, dimensions)

        Debug.Print "QA INFO: SECTION_DIMENSIONS|view=" & _
            SectionToken(viewName) & "|" & verdict

        Debug.Print "QA INFO: SECTION_LEGACY|view=" & _
            SectionToken(viewName) & _
            "|" & DetectLegacyBoreCallout(swView)

        If InStr(1, verdict, "requirementFailures=None", _
            vbBinaryCompare) > 0 Then

            satisfiedViews = satisfiedViews + 1
            bestVerdict = verdict
        ElseIf i = 1 Then
            bestVerdict = verdict
        End If
    Next i

    Dim finalSelectionCount As Long
    finalSelectionCount = _
        swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim drawingSaveAfter As Boolean
    drawingSaveAfter = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Debug.Print "R23_SECTIONDIM_END|sectionViews=" & _
        CStr(sectionViews.Count) & _
        "|satisfiedViews=" & CStr(satisfiedViews) & _
        "|" & bestVerdict & _
        "|creations=0|mutations=0" & _
        "|initialSelectionCount=" & CStr(initialSelectionCount) & _
        "|finalSelectionCount=" & CStr(finalSelectionCount) & _
        "|drawingUnchanged=" & _
        CStr(drawingSaveBefore = drawingSaveAfter)
    Exit Sub

Failed:
    Debug.Print "R23_SECTIONDIM_FATAL|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description

    On Error Resume Next
    If Not swDraw Is Nothing Then
        swDraw.SetPickMode
        swDraw.ClearSelection2 True
    End If
End Sub
