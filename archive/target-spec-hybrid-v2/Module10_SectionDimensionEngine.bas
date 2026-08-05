Option Explicit

' R23 Phase 8. Semantic section-dimension engine.
' SAFETY BOUNDARY. Exactly four procedures change a drawing:
' CreateSectionDimension, ApplyReferenceFit, CreateResolvedSectionDimensions
' (R23-821, which creates nothing itself and calls the first) and
' ClampSectionAnnotationsIntoUsableArea (R23-823, which creates nothing and
' only moves annotation text). All four refuse unless passed an explicit
' allowMutation argument. R23_ProbeSectionDimensions calls none of them and
' contains no AddDimension2, SetFitValues, SetValues2 or SetPosition2 call.
' R23-802. Reconciliation runs BEFORE creation, always. The section already
' carries imported dimensions - Phase 0 counted seventeen - and creating a
' second dimension for a requirement that is already satisfied is the defect
' this ordering exists to prevent. Each requirement records six independent
' observations about whatever it matched: source dimension identity,
' attached geometry, semantic role, nominal, type, and tolerance.
' R23-804. Phase 0 read type-6 D1@Sketch4 and D1@Sketch6 records from this
' fixture on 2026-07-31. The 2026-08-01 run of THIS module found a different
' drawing state: seven dimensions in Section View J-J, every one typed
' swLinearDimension = 2 and named RD1..RD7 - drawing-authored reference
' dimensions, not imported model dimensions. Both states are real, so the
' diameter requirements accept type 6, type 15 AND the linear types, and
' IDisplayDimension.Diametric is recorded to say which of them is displayed
' as a diameter. Requiring any single type would reject the real drawing,
' which is what R23-804 exists to prevent.
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

' R23-812 geometry inventory. MCP corpus, IView::GetPolylines7:
' CrossHatchOption 1 = exclude crosshatch lines. Crosshatch is excluded
' because a section is full of it and none of it is dimensionable geometry.
Private Const CROSSHATCH_EXCLUDE As Long = 1

' Documented underlying geometry types in the polyline record.
Private Const GEOM_TYPE_POLYLINE As Long = 0
Private Const GEOM_TYPE_ARC As Long = 1

' GeomData for an arc or circle is
' [cx,cy,cz, sx,sy,sz, ex,ey,ez, nx,ny,nz].
Private Const GEOM_ARC_DATA_SIZE As Long = 12

' Six scalars sit between GeomData and NumPolyPoints: LineColor, LineStyle,
' LineFont, LineWeight, LayerID, LayerOverride.
Private Const GEOM_STYLE_FIELD_COUNT As Long = 6

' Same quantum Phase 5 and Phase 7 use, so a coordinate means the same
' thing in all three.
Private Const GEOM_COORDINATE_QUANTUM_M As Double = 0.000001

' A segment counts as axis-parallel only when its off-axis delta is below
' this. Tighter than the coordinate quantum on purpose: the question being
' measured is which curves ARE axis-parallel, not which are nearly so.
Private Const GEOM_AXIS_TOLERANCE_M As Double = 0.0000001

' A candidate span matches a requirement nominal within 0.01 mm. The seven
' nominals are at least 5.5 mm apart, so this cannot confuse two of them.
Private Const GEOM_NOMINAL_TOLERANCE_M As Double = 0.00001

' Output bound. The inventory is diagnostic evidence, not a data dump.
Private Const GEOM_MAX_REPORTED As Long = 32

' R23-816. Half-width of the raw window printed when the decode loses
' alignment. Two of these either side of the index is enough to see three
' record headers and their fields; the r53 array held 2799 doubles, so a
' whole dump is not an option.
Private Const GEOM_WINDOW_SPAN As Long = 18

' MCP corpus value for swInConfigurationOpts_e.
Private Const CONFIG_THIS As Long = 1

' MCP corpus values for swDimensionTextParts_e. The prefix is where a
' diameter symbol lives when the dimension is not a diametric record.
Private Const TEXT_PREFIX As Long = 1
Private Const TEXT_PREFIX_DEFINITION As Long = 5

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

' R23-810. This marker detects stale pre-R23 bore notes on an existing
' drawing. Module7 no longer authors such a note on the production route.
Private Const LEGACY_BORE_CALLOUT_MARK As String = "47 H7"

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

    If Len(existing) = 0 Or _
        StrComp(existing, "None", vbBinaryCompare) = 0 Then
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

    ' R23-804. Diameters accept type 6, type 15 and the linear types. The
    ' live drawing carries its bore dimensions as linear records, so
    ' demanding 6 or 15 would reject the real thing. Whether a match is
    ' actually DISPLAYED as a diameter is answered by
    ' IDisplayDimension.Diametric and recorded on the match - the nominals
    ' here are 5.5 mm apart at the closest, so type corroborates rather than
    ' discriminates.
    Dim diameterTypes As String
    diameterTypes = CStr(DIM_TYPE_DIAMETER) & ";" & _
        CStr(DIM_TYPE_DIAMETRIC_LINEAR) & ";" & linearTypes

    result.Add MakeRequirement(REQ_OVERALL_THICKNESS, _
        "OverallThickness", 0.018, linearTypes, LANE_ABOVE)
    result.Add MakeRequirement(REQ_BORE_STEP_DEPTH, _
        "BoreStepDepth", 0.012, linearTypes, LANE_ABOVE)
    result.Add MakeRequirement(REQ_LOWER_WALL_STEP, _
        "LowerWallStep", 0.0115, linearTypes, LANE_BELOW)
    Dim innerBore As CSectionRequirement
    Set innerBore = MakeRequirement(REQ_INNER_BORE, _
        "InnerBoreDiameter", 0.04, diameterTypes, LANE_BORE_SIDE_A)
    innerBore.RequiresDiameterDisplay = True
    result.Add innerBore

    Dim fitBore As CSectionRequirement
    Set fitBore = MakeRequirement(REQ_FIT_BORE, _
        "FitBoreDiameter", 0.047, diameterTypes, LANE_BORE_SIDE_B)
    fitBore.RequiresDiameterDisplay = True

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
        Dim nominalRoute As String
        nominalM = 0#
        nominalAvailable = TryReadNominal(dimension, nominalM, nominalRoute)

        Dim diametricKnown As Boolean
        Dim diametric As Boolean
        diametric = SafeDiametric(displayDimension, diametricKnown)

        Dim prefixText As String
        Dim prefixDefinition As String
        Dim diameterSymbol As Boolean
        diameterSymbol = ReadDiameterPrefix( _
            displayDimension, prefixText, prefixDefinition)

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
            "|" & nominalRoute & _
            "|diametricKnown=" & CStr(diametricKnown) & _
            "|diametric=" & CStr(diametric) & _
            "|diameterSymbol=" & CStr(diameterSymbol) & _
            "|prefix=" & SectionToken(prefixText) & _
            "|prefixDefinition=" & SectionToken(prefixDefinition) & _
            "|attachedEntities=" & CStr(attachedCount) & _
            "|attachedTypes=" & attachedTypes & _
            "|" & toleranceProof

ContinueDimension:
    Next i

    Exit Function

Failed:
    ' Capture first: SafeViewName contains On Error Resume Next, which
    ' resets the global Err, and VBA evaluates the concatenation left to
    ' right. Reading Err.Number after it reports 0 for a real raise.
    Dim inventoryErrorNumber As Long
    inventoryErrorNumber = Err.Number

    EmitFailure evidence, "SECTION_DIM_INVENTORY_ERROR|view=" & _
        SectionToken(SafeViewName(swView)) & _
        "|error=" & CStr(inventoryErrorNumber)
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
' The configuration route is not enough on its own. The first live run
' returned no nominal for any of the seven section dimensions, all of them
' drawing-authored reference dimensions named RD1..RD7 rather than model
' dimensions imported from a sketch. A drawing reference dimension has no
' configuration to ask about, so the supported route can legitimately
' decline where it succeeds for D1@Sketch4.
' The 2026-08-01 second run settled it: all seven answered
' nominalRoute=Obsolete.GetSystemValue2, with exact nominals. The
' swAllConfiguration attempt was reached and declined on every one, so it is
' gone - a route with live evidence against it is not kept "just in case".
' Two routes remain. GetSystemValue3 with swThisConfiguration is the
' supported call and is what answers for imported model dimensions like
' D1@Sketch4. GetSystemValue2 is obsolete, is labelled obsolete in the route
' name, and is the only thing that answers for a drawing-authored reference
' dimension on this build. When both decline, the raw shape of the
' GetSystemValue3 result is reported, because "no nominal" and "an empty
' SafeArray" are different problems.
Private Function TryReadNominal( _
    ByRef dimension As SldWorks.Dimension, _
    ByRef nominalM As Double, _
    ByRef route As String) As Boolean

    route = "nominalRoute=NoDimension"
    If dimension Is Nothing Then Exit Function

    If TryNominalInConfiguration(dimension, CONFIG_THIS, nominalM) Then
        route = "nominalRoute=GetSystemValue3.ThisConfiguration"
        TryReadNominal = True
        Exit Function
    End If

    If TryNominalFromObsoleteMembers(dimension, nominalM, route) Then
        TryReadNominal = True
        Exit Function
    End If

    route = "nominalRoute=Unavailable" & _
        "|nominalShape=" & DescribeNominalShape(dimension)
End Function

Private Function TryNominalInConfiguration( _
    ByRef dimension As SldWorks.Dimension, _
    ByVal whichConfigurations As Long, _
    ByRef nominalM As Double) As Boolean

    On Error GoTo Failed

    Dim configurationNames As Variant
    configurationNames = Empty

    Dim values As Variant
    values = dimension.GetSystemValue3( _
        whichConfigurations, configurationNames)

    If IsArray(values) Then
        If (UBound(values) - LBound(values) + 1) < 1 Then Exit Function
        nominalM = CDbl(values(LBound(values)))
    ElseIf Not IsEmpty(values) And Not IsNull(values) Then
        nominalM = CDbl(values)
    Else
        Exit Function
    End If

    ' Every R23 section requirement is a positive length or diameter. A
    ' zero with no COM error is this build's decline value, not a usable
    ' nominal; continue to the documented fallback chain.
    If nominalM <= 0# Then Exit Function

    TryNominalInConfiguration = True
    Exit Function

Failed:
    TryNominalInConfiguration = False
End Function

' IDimension.GetSystemValue2 and IDimension.SystemValue are both marked
' obsolete by the 2025 Help. They run ONLY after the supported route has
' declined, and the route name says so, so no reader can mistake this for
' the preferred call. GetSystemValue2 is the one that answers on this build;
' SystemValue has never been reached and is kept as the final fallback
' rather than removed on no evidence either way.
Private Function TryNominalFromObsoleteMembers( _
    ByRef dimension As SldWorks.Dimension, _
    ByRef nominalM As Double, _
    ByRef route As String) As Boolean

    Dim candidate As Double

    On Error Resume Next

    candidate = 0#
    candidate = CDbl(dimension.GetSystemValue2(""))
    If Err.Number = 0 And candidate > 0# Then
        nominalM = candidate
        route = "nominalRoute=Obsolete.GetSystemValue2"
        TryNominalFromObsoleteMembers = True
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Err.Clear
    candidate = 0#
    candidate = CDbl(dimension.SystemValue)
    If Err.Number = 0 And candidate > 0# Then
        nominalM = candidate
        route = "nominalRoute=Obsolete.SystemValue"
        TryNominalFromObsoleteMembers = True
    End If

    Err.Clear
    On Error GoTo 0
End Function

' Says WHY no nominal came back rather than only that none did.
Private Function DescribeNominalShape( _
    ByRef dimension As SldWorks.Dimension) As String

    On Error GoTo Failed

    Dim configurationNames As Variant
    configurationNames = Empty

    Dim values As Variant
    values = dimension.GetSystemValue3(CONFIG_THIS, configurationNames)

    If IsEmpty(values) Then
        DescribeNominalShape = "Empty"
    ElseIf IsNull(values) Then
        DescribeNominalShape = "Null"
    ElseIf IsArray(values) Then
        DescribeNominalShape = "Array:" & _
            CStr(UBound(values) - LBound(values) + 1)
    Else
        DescribeNominalShape = "VarType:" & CStr(VarType(values))
    End If
    Exit Function

Failed:
    DescribeNominalShape = "Error:" & CStr(Err.Number)
End Function

' R23-804, third reading. The second live run returned diametric=False for
' every section dimension INCLUDING the 47 and the 40, with
' diametricKnown=True - a real answer, not a read failure. So the bore
' dimensions are plain linear records measuring across the bore.
' That is not yet enough to call them non-diameters. A drawing can carry the
' diameter symbol in the dimension's text PREFIX while the diametric flag
' stays false, and then the sheet reads correctly even though the record
' does not. Both the rendered prefix and its definition are read:
' IDisplayDimension.GetText(swDimensionTextPrefix) gives what is drawn, and
' swDimensionTextPrefixDefinition gives the authored form, where SOLIDWORKS
' writes the <MOD-DIAM> token.
' Unicode 216 is the diameter sign. It is written as ChrW$(216) rather than
' as a literal because every byte of this source must stay below 0x80.
Private Function ReadDiameterPrefix( _
    ByRef displayDimension As SldWorks.DisplayDimension, _
    ByRef prefixText As String, _
    ByRef definitionText As String) As Boolean

    prefixText = vbNullString
    definitionText = vbNullString

    On Error GoTo Failed

    prefixText = displayDimension.GetText(TEXT_PREFIX)
    definitionText = displayDimension.GetText(TEXT_PREFIX_DEFINITION)

    ReadDiameterPrefix = _
        (InStr(1, definitionText, "MOD-DIAM", vbTextCompare) > 0) Or _
        (InStr(1, prefixText, ChrW$(216), vbBinaryCompare) > 0) Or _
        (InStr(1, definitionText, ChrW$(216), vbBinaryCompare) > 0)
    Exit Function

Failed:
    prefixText = "Error:" & CStr(Err.Number)
    ReadDiameterPrefix = False
End Function

' The first live run found every section dimension typed
' swLinearDimension=2, including the one carrying H7 - not the
' swDiameterDimension=6 records Phase 0 saw in an earlier state of this
' drawing. IDisplayDimension.Diametric is what distinguishes a linear
' dimension DISPLAYED as a diameter from a plain one, so it is read for
' every dimension and recorded whether or not it decides anything.
Private Function SafeDiametric( _
    ByRef displayDimension As SldWorks.DisplayDimension, _
    ByRef known As Boolean) As Boolean

    known = False
    On Error GoTo Failed

    SafeDiametric = Module11_GeometryIdentity.NormalizeSwBoolean( _
        displayDimension.Diametric)
    known = True
    Exit Function

Failed:
    known = False
    SafeDiametric = False
End Function

' R23-805. Tolerance readback through IDimension.Tolerance.
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
' A requirement matches a dimension on nominal and accepted type together.
' The other four observations - source dimension identity, attached
' geometry, semantic role, tolerance - are recorded on the requirement so
' the match can be audited, and the tolerance additionally decides whether
' a matched requirement is fully satisfied.
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
            Dim nominalRoute As String
            nominalM = 0#
            nominalAvailable = _
                TryReadNominal(dimension, nominalM, nominalRoute)

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
                    dimension, typeCode, nominalAvailable, nominalM, _
                    nominalRoute
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
    ByVal nominalM As Double, _
    ByVal nominalRoute As String)

    requirement.NominalRoute = nominalRoute
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

    Dim diametricKnown As Boolean
    requirement.MatchedDiametric = _
        SafeDiametric(displayDimension, diametricKnown)
    requirement.MatchedDiametricKnown = diametricKnown

    Dim prefixText As String
    Dim prefixDefinition As String
    requirement.MatchedDiameterSymbol = ReadDiameterPrefix( _
        displayDimension, prefixText, prefixDefinition)
    requirement.MatchedDiameterPrefix = _
        SectionToken(prefixText) & "/" & SectionToken(prefixDefinition)

    If attachedCount = 0 Then
        requirement.Failures = AppendFailure( _
            requirement.Failures, "DimensionNotAttached")
    End If

    ' A diameter requirement matched against a linear record has to say so.
    ' The nominal is what identified it; whether the drawing DISPLAYS it as
    ' a diameter is a separate fact, and an unproved one is reported rather
    ' than assumed in either direction.
    If requirement.RequiresDiameterDisplay Then
        If typeCode = DIM_TYPE_DIAMETER Or _
            typeCode = DIM_TYPE_DIAMETRIC_LINEAR Then

            requirement.DiameterDisplaySource = "DiametricRecord"
        ElseIf requirement.MatchedDiameterSymbol Then
            ' The record is linear but the sheet carries the diameter
            ' symbol, so what a machinist reads is a diameter. Recorded as
            ' the weaker source it is, not failed.
            requirement.DiameterDisplaySource = "TextPrefix"
        ElseIf requirement.MatchedDiametric Then
            requirement.DiameterDisplaySource = "DiametricFlag"
        ElseIf Not diametricKnown Then
            requirement.DiameterDisplaySource = "Unreadable"
            requirement.Failures = AppendFailure( _
                requirement.Failures, _
                "DiameterDisplayUnreadable:" & CStr(typeCode))
        Else
            requirement.DiameterDisplaySource = "None"
            requirement.Failures = AppendFailure( _
                requirement.Failures, _
                "NotDisplayedAsDiameter:" & CStr(typeCode))
        End If
    End If
End Sub

' R23-806. Decides whether a matched requirement's tolerance is already
' satisfied, and states where the answer came from.
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
        ElseIf Len(requirement.Failures) > 0 And _
            StrComp(requirement.Failures, "None", vbBinaryCompare) <> 0 Then

            ' Anything reconciliation recorded against this requirement
            ' keeps it out of the satisfied count. A requirement whose own
            ' failure list is non-empty is not satisfied, whatever the
            ' counts say.
            failures = AppendFailure(failures, _
                "RequirementFlagged:" & requirement.Key & _
                ":" & requirement.Failures)
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

' R23-810. Finds stale pre-R23 free-text bore callouts for evidence only.
' Any removal remains a separately authorized drawing mutation.
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
        "|removalStatus=PipelineSwitched|removalRequiresAuthorization=True"
    Exit Function

Failed:
    DetectLegacyBoreCallout = "legacyBoreCallout=Error:" & CStr(Err.Number)
End Function

' R23-806. MUTATES THE DRAWING.
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
' Creates the one associative dimension a requirement needs, and only when
' reconciliation proved nothing already satisfies it.
' The entities must already be selected. IModelDoc2.AddDimension2's Remarks
' are explicit that selections are made by LOCATION and never by name -
' passing a name makes the dimensioning routines pick an endpoint at random
' - so choosing and selecting them is the caller's responsibility and is
' verified here rather than assumed.
' textX and textY are the dimension text position on the sheet. Like the
' section view's placement in Phase 7, they are caller arguments: the lane
' on the requirement names WHERE the dimension belongs, and turning a lane
' into coordinates needs the finished section's annotation envelope.
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
    Dim nominalRoute As String
    nominalM = 0#
    nominalAvailable = TryReadNominal(dimension, nominalM, nominalRoute)

    RecordObservation requirement, created, dimension, typeCode, _
        nominalAvailable, nominalM, nominalRoute

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

' R23-812. READ-ONLY geometry inventory for one section view.
'
' Section View J-J did not exist before r49, so nothing has ever measured
' what it offers. SECTION_DIMENSIONS reports 7 missing / 0 satisfied and
' CreateSectionDimension refuses without a selection
' (reason=NoEntitiesSelected), so the open question is exactly: which
' drawing curves does this view expose, and does any pair of them span a
' required nominal.
'
' Record layout is the documented one, SOLIDWORKS 2025 Help,
' IView::GetPolylines7:
'   [Type, GeomDataSize, GeomData[], LineColor, LineStyle, LineFont,
'    LineWeight, LayerID, LayerOverride, NumPolyPoints, [x,y,z]...]
' repeated once per polyline, with Type 1 carrying centre, start, end and
' normal points.
'
' The decode is self-checking. It stops on any out-of-range field and
' reports Desynchronized rather than emitting invented geometry, and the
' number of records it decodes must equal the number of entries in the
' returned entity array - the Help states the two arrays are positionally
' paired, with Null entity entries for silhouette edges. That equality is
' the control; without it the numbers below would be unfalsifiable.
'
' Nothing here selects, creates or modifies anything.
Public Sub InventorySectionGeometry( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef evidence As CRunEvidence, _
    ByRef resolvedRequirements As Collection, _
    ByRef resolvedEntities As Variant)

    On Error GoTo Failed

    If swView Is Nothing Then Exit Sub

    Dim viewName As String
    viewName = SafeViewName(swView)

    Dim scaleDecimal As Double
    Dim useSheetScale As Long
    scaleDecimal = SafeScaleDecimal(swView, useSheetScale)

    Dim polylineData As Variant
    Dim entities As Variant
    Dim readError As Long
    On Error Resume Next
    entities = swView.GetPolylines7(CROSSHATCH_EXCLUDE, polylineData)
    readError = Err.Number
    Err.Clear
    On Error GoTo Failed

    Dim entityCount As Long
    entityCount = SafeItemCount(entities)

    Dim dataCount As Long
    dataCount = SafeItemCount(polylineData)

    If dataCount = 0 Then
        EmitWarning evidence, "SECTION_GEOM_SUMMARY|view=" & viewName & _
            "|status=NoPolylineData|entities=" & CStr(entityCount) & _
            "|error=" & CStr(readError)
        Exit Sub
    End If

    Dim horizontalYs As Collection
    Dim verticalXs As Collection
    Dim radii As Collection
    Set horizontalYs = New Collection
    Set verticalXs = New Collection
    Set radii = New Collection

    ' R23-819. Parallel to the three above: the polyline record that
    ' produced each distinct coordinate, which is also its index in the
    ' entity array.
    Dim horizontalYRecords As Collection
    Dim verticalXRecords As Collection
    Dim radiusRecords As Collection
    Set horizontalYRecords = New Collection
    Set verticalXRecords = New Collection
    Set radiusRecords = New Collection

    Dim decoded As Long
    Dim arcCount As Long
    Dim polylineCount As Long
    Dim pointTotal As Long
    Dim horizontalSegments As Long
    Dim verticalSegments As Long
    Dim otherSegments As Long
    Dim arcTessellationSegments As Long
    Dim decodeStatus As String
    decodeStatus = "Complete"

    Dim minX As Double
    Dim minY As Double
    Dim maxX As Double
    Dim maxY As Double
    Dim boundsSeeded As Boolean

    Dim i As Long
    Dim lower As Long
    Dim upper As Long
    lower = LBound(polylineData)
    upper = UBound(polylineData)
    i = lower

    ' R23-816. The walk is bounded by the entity count and remembers where
    ' the last three records began.
    '
    ' The r53 run desynchronized: decodeStatus=Desynchronized:StyleAt2801
    ' with records=135 against entities=79. The Help states the polyline
    ' data and the returned entity array are positionally paired, so 135
    ' records from 79 entities means the walk lost alignment and kept
    ' going - it could not have been right, and nothing in the output said
    ' where it went wrong.
    '
    ' The layout itself is not in doubt. Against the r50 and r52 data it
    ' closes to the double: 38 records * 9 fixed fields + 6 arcs * 12
    ' GeomData + 3 * 214 points = 1056, exactly the array length. Something
    ' in the richer r53 section deviates from it, and a window of raw values
    ' around the failure is what identifies that, not another guess at the
    ' layout.
    Dim recordStart As Long
    Dim priorStart1 As Long
    Dim priorStart2 As Long
    Dim priorStart3 As Long
    priorStart1 = -1
    priorStart2 = -1
    priorStart3 = -1

    Do While i <= upper
        ' R23-818. The entity count IS the record count, and the array is
        ' zero-padded beyond it.
        '
        ' r54 bounded the walk here; r56 removed the bound after the r55
        ' run stopped at 79 records with 510 doubles still to go, which I
        ' read as proof that more records existed. That was wrong, and the
        ' r56 window shows why: every double from index 2289 to the end of
        ' the array is 0.000000000. The unbounded walk parsed 510 zeros as
        ' 56 phantom records of stride 9 - type 0, GeomDataSize 0, six zero
        ' style fields, NumPolyPoints 0 - reaching the reported 135, with 6
        ' doubles left over that cannot complete a record. That is the
        ' whole of Desynchronized:StyleAt2801.
        '
        ' So 79 real records consume exactly 2289 doubles
        ' (9*79 + 12*18 + 3*454) and the Help's positional pairing holds
        ' after all. The bound is restored, and the padding is now VERIFIED
        ' to be zeros rather than assumed - a non-zero tail would mean real
        ' records the bound is discarding, which is the failure this
        ' iteration exists to make visible.
        If entityCount > 0 And decoded >= entityCount Then Exit Do

        recordStart = i

        If i + 1 > upper Then
            decodeStatus = "Desynchronized:HeaderAt" & CStr(i)
            Exit Do
        End If

        Dim recordType As Long
        Dim geomSize As Long
        recordType = CLng(polylineData(i))
        geomSize = CLng(polylineData(i + 1))
        i = i + 2

        If (recordType <> GEOM_TYPE_POLYLINE And _
            recordType <> GEOM_TYPE_ARC) Or geomSize < 0 Or _
            i + geomSize > upper + 1 Then

            decodeStatus = "Desynchronized:TypeAt" & CStr(i - 2) & _
                ":type" & CStr(recordType) & ":size" & CStr(geomSize)
            Exit Do
        End If

        If recordType = GEOM_TYPE_ARC And _
            geomSize >= GEOM_ARC_DATA_SIZE Then

            Dim centreX As Double
            Dim centreY As Double
            Dim centreZ As Double
            Dim startX As Double
            Dim startY As Double
            Dim startZ As Double
            centreX = CDbl(polylineData(i))
            centreY = CDbl(polylineData(i + 1))
            centreZ = CDbl(polylineData(i + 2))
            startX = CDbl(polylineData(i + 3))
            startY = CDbl(polylineData(i + 4))
            startZ = CDbl(polylineData(i + 5))

            Dim radius As Double
            radius = Sqr(((startX - centreX) ^ 2) + _
                ((startY - centreY) ^ 2) + ((startZ - centreZ) ^ 2))

            AddDistinctValue radii, radiusRecords, radius, _
                GEOM_COORDINATE_QUANTUM_M, decoded
            arcCount = arcCount + 1
        ElseIf recordType = GEOM_TYPE_POLYLINE Then
            polylineCount = polylineCount + 1
        End If

        i = i + geomSize + GEOM_STYLE_FIELD_COUNT

        If i > upper Then
            decodeStatus = "Desynchronized:StyleAt" & CStr(i)
            Exit Do
        End If

        Dim pointCount As Long
        pointCount = CLng(polylineData(i))
        i = i + 1

        If pointCount < 0 Or i + (3 * pointCount) > upper + 1 Then
            decodeStatus = "Desynchronized:PointsAt" & CStr(i) & _
                ":count" & CStr(pointCount)
            Exit Do
        End If

        pointTotal = pointTotal + pointCount

        Dim pointIndex As Long
        For pointIndex = 0 To pointCount - 1
            Dim px As Double
            Dim py As Double
            px = CDbl(polylineData(i + (3 * pointIndex)))
            py = CDbl(polylineData(i + (3 * pointIndex) + 1))

            If Not boundsSeeded Then
                minX = px
                maxX = px
                minY = py
                maxY = py
                boundsSeeded = True
            Else
                If px < minX Then minX = px
                If px > maxX Then maxX = px
                If py < minY Then minY = py
                If py > maxY Then maxY = py
            End If

            ' R23-820. Only a Type 0 record's points describe straight
            ' edges. An arc record also carries a tessellated point array,
            ' and walking it produced a false straight edge in r58: record
            ' 25 supplied BOTH the 0.023500 radius for FIT_BORE_D47_H7 and
            ' the x=0.008500 side of LOWER_WALL_STEP_11_5, because a chord
            ' of the bore's tessellation happens to run parallel to an
            ' axis. A linear dimension attached there would have measured
            ' the bore and looked correct on the sheet.
            If pointIndex > 0 And recordType = GEOM_TYPE_POLYLINE Then
                Dim previousX As Double
                Dim previousY As Double
                previousX = CDbl(polylineData(i + (3 * pointIndex) - 3))
                previousY = CDbl(polylineData(i + (3 * pointIndex) - 2))

                Dim deltaX As Double
                Dim deltaY As Double
                deltaX = Abs(px - previousX)
                deltaY = Abs(py - previousY)

                If deltaY <= GEOM_AXIS_TOLERANCE_M And _
                    deltaX > GEOM_AXIS_TOLERANCE_M Then

                    horizontalSegments = horizontalSegments + 1
                    AddDistinctValue horizontalYs, horizontalYRecords, _
                        py, GEOM_COORDINATE_QUANTUM_M, decoded
                ElseIf deltaX <= GEOM_AXIS_TOLERANCE_M And _
                    deltaY > GEOM_AXIS_TOLERANCE_M Then

                    verticalSegments = verticalSegments + 1
                    AddDistinctValue verticalXs, verticalXRecords, _
                        px, GEOM_COORDINATE_QUANTUM_M, decoded
                Else
                    otherSegments = otherSegments + 1
                End If
            ElseIf pointIndex > 0 Then
                arcTessellationSegments = arcTessellationSegments + 1
            End If
        Next pointIndex

        i = i + (3 * pointCount)
        decoded = decoded + 1

        priorStart3 = priorStart2
        priorStart2 = priorStart1
        priorStart1 = recordStart
    Loop

    Dim trailingDoubles As Long
    trailingDoubles = upper - i + 1
    If trailingDoubles < 0 Then trailingDoubles = 0

    ' R23-818. Everything after the last real record must be zero padding.
    ' Verified, not assumed: a non-zero tail would mean the entity-count
    ' bound is discarding real records, and that has to be visible rather
    ' than silently trimmed away.
    Dim trailingAllZero As Boolean
    trailingAllZero = True

    Dim tailIndex As Long
    For tailIndex = i To upper
        If CDbl(polylineData(tailIndex)) <> 0# Then
            trailingAllZero = False
            Exit For
        End If
    Next tailIndex

    If StrComp(decodeStatus, "Complete", vbBinaryCompare) = 0 Then
        If decoded <> entityCount Then
            decodeStatus = "RecordCountMismatch"
        ElseIf Not trailingAllZero Then
            decodeStatus = "TrailingDataAfterEntities"
        End If
    End If

    EmitInfo evidence, "SECTION_GEOM_SUMMARY|view=" & viewName & _
        "|decodeStatus=" & decodeStatus & _
        "|records=" & CStr(decoded) & _
        "|entities=" & CStr(entityCount) & _
        "|recordsMatchEntities=" & CStr(decoded = entityCount) & _
        "|doubles=" & CStr(dataCount) & _
        "|consumed=" & CStr(i - lower) & _
        "|trailing=" & CStr(trailingDoubles) & _
        "|trailingAllZero=" & CStr(trailingAllZero) & _
        "|arcs=" & CStr(arcCount) & _
        "|polylines=" & CStr(polylineCount) & _
        "|points=" & CStr(pointTotal) & _
        "|error=" & CStr(readError)

    ' R23-816. On anything but a clean decode, print the raw values around
    ' the point where alignment was lost and around the last three record
    ' headers. 2799 doubles is far too many to dump whole; a window is what
    ' identifies the deviation, and without it the next iteration would be
    ' guessing at the layout again.
    If StrComp(decodeStatus, "Complete", vbBinaryCompare) <> 0 Then
        EmitWarning evidence, "SECTION_GEOM_DESYNC|view=" & viewName & _
            "|status=" & decodeStatus & _
            "|stoppedAt=" & CStr(i - lower) & _
            "|recordsDecoded=" & CStr(decoded) & _
            "|lastRecordStarts=" & CStr(priorStart1 - lower) & _
            "," & CStr(priorStart2 - lower) & _
            "," & CStr(priorStart3 - lower)

        EmitRawWindow evidence, viewName, polylineData, lower, upper, _
            priorStart3, "lastRecords"
        EmitRawWindow evidence, viewName, polylineData, lower, upper, _
            i - GEOM_WINDOW_SPAN, "stopPoint"
    End If

    EmitInfo evidence, "SECTION_GEOM_SEGMENTS|view=" & viewName & _
        "|horizontal=" & CStr(horizontalSegments) & _
        "|vertical=" & CStr(verticalSegments) & _
        "|other=" & CStr(otherSegments) & _
        "|arcTessellation=" & CStr(arcTessellationSegments) & _
        "|distinctY=" & CStr(horizontalYs.Count) & _
        "|distinctX=" & CStr(verticalXs.Count) & _
        "|distinctRadii=" & CStr(radii.Count)

    ' The frame these coordinates live in has never been proved for a
    ' section view. Reporting the box next to the sheet-space outline makes
    ' it decidable from evidence instead of assumed.
    Dim outline As Variant
    Dim outlineToken As String
    outlineToken = "Unavailable"
    On Error Resume Next
    outline = swView.GetOutline
    If IsArray(outline) Then
        outlineToken = FormatMetres(CDbl(outline(0))) & "," & _
            FormatMetres(CDbl(outline(1))) & "," & _
            FormatMetres(CDbl(outline(2))) & "," & _
            FormatMetres(CDbl(outline(3)))
    End If
    Err.Clear
    On Error GoTo Failed

    EmitInfo evidence, "SECTION_GEOM_FRAME|view=" & viewName & _
        "|polylineBox=" & FormatMetres(minX) & "," & _
        FormatMetres(minY) & "," & FormatMetres(maxX) & "," & _
        FormatMetres(maxY) & _
        "|sheetOutline=" & outlineToken & _
        "|useSheetScale=" & CStr(useSheetScale) & _
        "|scaleDecimal=" & Format$(scaleDecimal, "0.000000")

    EmitInfo evidence, "SECTION_GEOM_Y|view=" & viewName & _
        "|decodeStatus=" & decodeStatus & "|values=" & _
        FormatDistinctList(horizontalYs)
    EmitInfo evidence, "SECTION_GEOM_X|view=" & viewName & _
        "|decodeStatus=" & decodeStatus & "|values=" & _
        FormatDistinctList(verticalXs)
    EmitInfo evidence, "SECTION_GEOM_R|view=" & viewName & _
        "|decodeStatus=" & decodeStatus & "|values=" & _
        FormatDistinctList(radii)

    ' R23-816. Coordinates from a walk that lost alignment are not evidence
    ' of anything. The r53 report carried seven radii and six found=True
    ' verdicts from a decode that had already desynchronized, and only the
    ' summary line said so. Every consumer now carries the verdict with it.
    ' R23-821. The resolved candidates and the entity array they index into
    ' are handed back so the creation pass can select exactly what the
    ' inventory measured, rather than searching for it a second time.
    resolvedEntities = entities
    Set resolvedRequirements = ReportRequirementCandidates( _
        swDrawModel, swView, entities, viewName, _
        horizontalYs, horizontalYRecords, verticalXs, verticalXRecords, _
        radii, radiusRecords, scaleDecimal, decodeStatus, evidence)
    Exit Sub

Failed:
    ' Capture first: SafeViewName contains On Error Resume Next, which
    ' resets the global Err before the concatenation reaches Err.Number.
    Dim inventoryErrorNumber As Long
    inventoryErrorNumber = Err.Number

    EmitWarning evidence, "SECTION_GEOM_SUMMARY|view=" & _
        SafeViewName(swView) & "|status=Error:" & _
        CStr(inventoryErrorNumber)
End Sub

' For each of the seven requirements, whether the geometry that would
' measure it is present in this view at all. A diameter requirement wants a
' curve of half its nominal radius; a linear requirement wants two parallel
' curves its nominal apart. This says nothing about whether such a
' selection would produce the right dimension - only whether the candidate
' exists, which is the thing currently unknown.
Private Function ReportRequirementCandidates( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef entities As Variant, _
    ByVal viewName As String, _
    ByRef horizontalYs As Collection, _
    ByRef horizontalYRecords As Collection, _
    ByRef verticalXs As Collection, _
    ByRef verticalXRecords As Collection, _
    ByRef radii As Collection, _
    ByRef radiusRecords As Collection, _
    ByVal scaleDecimal As Double, _
    ByVal decodeStatus As String, _
    ByRef evidence As CRunEvidence) As Collection

    Dim requirements As Collection
    Set requirements = BuildSectionRequirements()
    Set ReportRequirementCandidates = requirements

    On Error GoTo Failed

    ' A found=True from a desynchronized walk is not a finding.
    Dim geometryTrusted As Boolean
    geometryTrusted = _
        (StrComp(decodeStatus, "Complete", vbBinaryCompare) = 0)

    Dim i As Long
    For i = 1 To requirements.Count
        Dim requirement As CSectionRequirement
        Set requirement = requirements(i)

        Dim kind As String
        Dim found As Boolean
        Dim first As Double
        Dim second As Double
        Dim firstRecord As Long
        Dim secondRecord As Long

        ' Reset per requirement. The r50 run printed the PREVIOUS
        ' requirement's coordinates on every found=False row, which reads
        ' as a near miss when nothing was found at all.
        first = 0#
        second = 0#
        firstRecord = -1
        secondRecord = -1
        kind = "None"
        found = False

        ' A bore that has been cut open is not an arc in the section - it
        ' is the gap between two cut walls. r50 measured exactly that: the
        ' section held no 0.020 or 0.023500 radius anywhere. So a diameter
        ' requirement is searched as a radius first and then as a linear
        ' span, and a linear requirement is searched on both axes.
        If requirement.RequiresDiameterDisplay Then
            kind = "Radius"
            found = FindValueNear(radii, radiusRecords, _
                requirement.NominalM / 2#, first, firstRecord)
            If found Then
                second = first
                secondRecord = firstRecord
            End If
        End If

        If Not found Then
            kind = "YPair"
            found = FindPairWithSpan(horizontalYs, horizontalYRecords, _
                requirement.NominalM, first, second, _
                firstRecord, secondRecord)
        End If

        If Not found Then
            kind = "XPair"
            found = FindPairWithSpan(verticalXs, verticalXRecords, _
                requirement.NominalM, first, second, _
                firstRecord, secondRecord)
        End If

        If Not found Then
            kind = "NoneOfRadiusYPairXPair"
            first = 0#
            second = 0#
            firstRecord = -1
            secondRecord = -1
        End If

        EmitInfo evidence, "SECTION_REQ_CANDIDATE|view=" & viewName & _
            "|key=" & requirement.Key & _
            "|nominalM=" & FormatMetres(requirement.NominalM) & _
            "|kind=" & kind & _
            "|found=" & CStr(found) & _
            "|geometryTrusted=" & CStr(geometryTrusted) & _
            "|decodeStatus=" & decodeStatus & _
            "|a=" & FormatMetres(first) & _
            "|b=" & FormatMetres(second) & _
            "|recordA=" & CStr(firstRecord) & _
            "|recordB=" & CStr(secondRecord) & _
            "|scaleDecimal=" & Format$(scaleDecimal, "0.000000")

        ' R23-819. Whether the measured geometry is SELECTABLE is the thing
        ' CreateSectionDimension actually needs; it refuses with
        ' reason=NoEntitiesSelected otherwise. Only attempted for a
        ' requirement whose geometry was found by a trusted decode, so at
        ' most twelve selections are made and each is cleared.
        requirement.CandidateFound = (found And geometryTrusted)
        requirement.CandidateKind = kind
        requirement.CandidateRecordA = firstRecord
        requirement.CandidateRecordB = secondRecord

        If found And geometryTrusted Then
            requirement.CandidateSelectableA = _
                ReportCandidateSelectability(swDrawModel, swView, _
                    entities, viewName, requirement.Key, "A", _
                    firstRecord, evidence)

            If secondRecord <> firstRecord Then
                requirement.CandidateSelectableB = _
                    ReportCandidateSelectability(swDrawModel, swView, _
                        entities, viewName, requirement.Key, "B", _
                        secondRecord, evidence)
            Else
                ' A radius requirement is one entity; side B is the same
                ' record and is not selected twice.
                requirement.CandidateSelectableB = _
                    requirement.CandidateSelectableA
            End If
        End If
    Next i

    Exit Function

Failed:
    EmitWarning evidence, "SECTION_REQ_CANDIDATE|view=" & viewName & _
        "|status=Error:" & CStr(Err.Number)
End Function

' R23-816. Raw values around one index, six per line, so a layout
' deviation can be read out of the log instead of theorised about.
Private Sub EmitRawWindow( _
    ByRef evidence As CRunEvidence, _
    ByVal viewName As String, _
    ByRef values As Variant, _
    ByVal lower As Long, _
    ByVal upper As Long, _
    ByVal centreIndex As Long, _
    ByVal label As String)

    On Error GoTo Failed

    Dim first As Long
    Dim last As Long
    first = centreIndex
    If first < lower Then first = lower
    last = first + (2 * GEOM_WINDOW_SPAN) - 1
    If last > upper Then last = upper
    If first > last Then Exit Sub

    Dim chunk As String
    Dim chunkFrom As Long
    Dim emitted As Long
    chunkFrom = first
    chunk = vbNullString

    Dim i As Long
    For i = first To last
        If Len(chunk) > 0 Then chunk = chunk & ";"
        chunk = chunk & Format$(CDbl(values(i)), "0.000000000")
        emitted = emitted + 1

        If (emitted Mod 6 = 0) Or i = last Then
            EmitWarning evidence, "SECTION_GEOM_WINDOW|view=" & viewName & _
                "|label=" & label & _
                "|from=" & CStr(chunkFrom - lower) & _
                "|values=" & chunk
            chunkFrom = i + 1
            chunk = vbNullString
        End If
    Next i

    Exit Sub

Failed:
    EmitWarning evidence, "SECTION_GEOM_WINDOW|view=" & viewName & _
        "|label=" & label & "|status=Error:" & CStr(Err.Number)
End Sub

' R23-819. Each distinct coordinate remembers the polyline record that
' produced it, because the record index is also the index of the paired
' entry in the entity array GetPolylines7 returns. That pairing is what
' turns a measured coordinate back into a selectable drawing entity, and
' without it the inventory can say a requirement's geometry exists but
' never point at the thing a dimension would attach to.
Private Sub AddDistinctValue( _
    ByRef values As Collection, _
    ByRef records As Collection, _
    ByVal value As Double, _
    ByVal tolerance As Double, _
    ByVal recordIndex As Long)

    Dim i As Long
    For i = 1 To values.Count
        If Abs(CDbl(values(i)) - value) <= tolerance Then Exit Sub
    Next i

    values.Add value
    records.Add recordIndex
End Sub

Private Function FormatDistinctList( _
    ByRef values As Collection) As String

    Dim result As String
    Dim reported As Long

    Dim i As Long
    For i = 1 To values.Count
        If reported >= GEOM_MAX_REPORTED Then
            result = result & ";+" & CStr(values.Count - reported) & "more"
            Exit For
        End If

        If Len(result) > 0 Then result = result & ";"
        result = result & FormatMetres(CDbl(values(i)))
        reported = reported + 1
    Next i

    If Len(result) = 0 Then result = "None"
    FormatDistinctList = result
End Function

Private Function FindValueNear( _
    ByRef values As Collection, _
    ByRef records As Collection, _
    ByVal target As Double, _
    ByRef foundValue As Double, _
    ByRef foundRecord As Long) As Boolean

    Dim i As Long
    For i = 1 To values.Count
        If Abs(CDbl(values(i)) - target) <= GEOM_NOMINAL_TOLERANCE_M Then
            foundValue = CDbl(values(i))
            foundRecord = CLng(records(i))
            FindValueNear = True
            Exit Function
        End If
    Next i
End Function

Private Function FindPairWithSpan( _
    ByRef values As Collection, _
    ByRef records As Collection, _
    ByVal span As Double, _
    ByRef firstValue As Double, _
    ByRef secondValue As Double, _
    ByRef firstRecord As Long, _
    ByRef secondRecord As Long) As Boolean

    Dim i As Long
    For i = 1 To values.Count
        Dim j As Long
        For j = i + 1 To values.Count
            If Abs(Abs(CDbl(values(j)) - CDbl(values(i))) - span) <= _
                GEOM_NOMINAL_TOLERANCE_M Then

                firstValue = CDbl(values(i))
                secondValue = CDbl(values(j))
                firstRecord = CLng(records(i))
                secondRecord = CLng(records(j))
                FindPairWithSpan = True
                Exit Function
            End If
        Next j
    Next i
End Function

' R23-821. Creates the section dimensions whose measuring geometry the
' inventory found AND proved selectable, and nothing else.
'
' This is the third mutating procedure in the module and it refuses without
' an explicit allowMutation, like the other two. It creates nothing for a
' requirement that reconciliation already matched (R23-802: reconcile
' first, create only what is missing), nothing for a requirement whose
' candidate was not found, and nothing for one whose entities did not prove
' selectable.
'
' The created dimension is still verified by CreateSectionDimension: it
' reads the nominal back and rejects a mismatch as CreatedButRejected. That
' readback is what makes a wrong selection or a wrong coordinate frame
' visible instead of producing a dimension that merely looks right.
Public Function CreateResolvedSectionDimensions( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef entities As Variant, _
    ByRef requirements As Collection, _
    ByVal allowMutation As Boolean, _
    ByRef evidence As CRunEvidence) As Long

    On Error GoTo Failed

    If Not allowMutation Then
        EmitWarning evidence, "SECTION_DIM_PASS_REFUSED" & _
            "|reason=MutationNotAuthorized"
        Exit Function
    End If

    If requirements Is Nothing Or swDrawModel Is Nothing Then Exit Function
    If swView Is Nothing Then Exit Function

    Dim viewName As String
    viewName = SafeViewName(swView)

    Dim created As Long

    ' R23-822. One ordinal per lane. A global counter made the fifth
    ' dimension start 60 mm off its view whichever lane it belonged to.
    Dim laneOrdinals As Collection
    Set laneOrdinals = New Collection
    Dim laneOrdinal As Long

    Dim i As Long
    For i = 1 To requirements.Count
        Dim requirement As CSectionRequirement
        Set requirement = requirements(i)

        Dim skipReason As String
        skipReason = vbNullString

        If requirement.Matched Then
            skipReason = "AlreadySatisfiedByImportedDimension"
        ElseIf Not requirement.CandidateFound Then
            skipReason = "NoCandidateGeometry:" & requirement.CandidateKind
        ElseIf Not requirement.CandidateSelectableA Then
            skipReason = "CandidateSideANotSelectable"
        ElseIf Not requirement.CandidateSelectableB Then
            skipReason = "CandidateSideBNotSelectable"
        End If

        If Len(skipReason) > 0 Then
            EmitInfo evidence, "SECTION_DIM_SKIPPED|view=" & viewName & _
                "|key=" & requirement.Key & "|reason=" & skipReason
            GoTo ContinueRequirement
        End If

        If Not SelectCandidateEntities( _
            swDrawModel, swView, entities, requirement, evidence) Then

            GoTo ContinueRequirement
        End If

        laneOrdinal = NextLaneOrdinal(laneOrdinals, requirement.Lane)

        Dim textX As Double
        Dim textY As Double
        Dim placementProof As String
        LaneTextPoint swView, requirement.Lane, laneOrdinal, evidence, _
            textX, textY, placementProof

        EmitInfo evidence, "SECTION_DIM_PLACEMENT|view=" & viewName & _
            "|key=" & requirement.Key & _
            "|lane=" & requirement.Lane & _
            "|textX=" & FormatMetres(textX) & _
            "|textY=" & FormatMetres(textY) & _
            "|" & placementProof

        Dim createdDimension As SldWorks.DisplayDimension
        Set createdDimension = CreateSectionDimension( _
            swDrawModel, requirement, textX, textY, allowMutation, _
            evidence)

        swDrawModel.ClearSelection2 True

        If requirement.Matched Then created = created + 1

ContinueRequirement:
    Next i

    EmitInfo evidence, "SECTION_DIM_PASS|view=" & viewName & _
        "|requirements=" & CStr(requirements.Count) & _
        "|created=" & CStr(created)

    CreateResolvedSectionDimensions = created
    Exit Function

Failed:
    Dim passErrorNumber As Long
    passErrorNumber = Err.Number
    EmitFailure evidence, "SECTION_DIM_PASS_ERROR|view=" & _
        SafeViewName(swView) & "|error=" & CStr(passErrorNumber)
    On Error Resume Next
    swDrawModel.ClearSelection2 True
End Function

' Selects the one or two entities the requirement measures, appending the
' second so both are in the selection list when AddDimension2 runs.
' SOLIDWORKS reads the selection, so an unverified count produces a
' dimension between the wrong things.
Private Function SelectCandidateEntities( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef entities As Variant, _
    ByRef requirement As CSectionRequirement, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    swDrawModel.ClearSelection2 True

    Dim expected As Long
    expected = 1
    If requirement.CandidateRecordB <> requirement.CandidateRecordA Then
        expected = 2
    End If

    Dim selectedA As Boolean
    selectedA = SelectOneCandidate( _
        swView, entities, requirement.CandidateRecordA, False)

    Dim selectedB As Boolean
    selectedB = True
    If expected = 2 Then
        selectedB = SelectOneCandidate( _
            swView, entities, requirement.CandidateRecordB, True)
    End If

    Dim actual As Long
    actual = swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)

    If Not selectedA Or Not selectedB Or actual <> expected Then
        EmitFailure evidence, "SECTION_DIM_SELECT_FAILED" & _
            "|key=" & requirement.Key & _
            "|recordA=" & CStr(requirement.CandidateRecordA) & _
            "|recordB=" & CStr(requirement.CandidateRecordB) & _
            "|selectedA=" & CStr(selectedA) & _
            "|selectedB=" & CStr(selectedB) & _
            "|expected=" & CStr(expected) & _
            "|actual=" & CStr(actual)
        swDrawModel.ClearSelection2 True
        Exit Function
    End If

    EmitInfo evidence, "SECTION_DIM_SELECTED|key=" & requirement.Key & _
        "|recordA=" & CStr(requirement.CandidateRecordA) & _
        "|recordB=" & CStr(requirement.CandidateRecordB) & _
        "|count=" & CStr(actual)

    SelectCandidateEntities = True
    Exit Function

Failed:
    Dim selectErrorNumber As Long
    selectErrorNumber = Err.Number
    EmitFailure evidence, "SECTION_DIM_SELECT_ERROR|key=" & _
        requirement.Key & "|error=" & CStr(selectErrorNumber)
    On Error Resume Next
    swDrawModel.ClearSelection2 True
End Function

Private Function SelectOneCandidate( _
    ByRef swView As SldWorks.View, _
    ByRef entities As Variant, _
    ByVal recordIndex As Long, _
    ByVal appendToSelection As Boolean) As Boolean

    On Error GoTo Failed

    If recordIndex < 0 Or Not IsArray(entities) Then Exit Function

    Dim entityIndex As Long
    entityIndex = LBound(entities) + recordIndex
    If entityIndex > UBound(entities) Then Exit Function

    Dim candidate As Object
    Set candidate = Nothing
    On Error Resume Next
    Set candidate = entities(entityIndex)
    Err.Clear
    On Error GoTo Failed

    If candidate Is Nothing Then Exit Function

    SelectOneCandidate = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swView.SelectEntity(candidate, appendToSelection))
    Exit Function

Failed:
    SelectOneCandidate = False
End Function

' R23-822. How many entries this lane has already taken, plus one. Keyed by
' the lane name so two lanes never share a counter. A Collection keyed by
' string is the only associative store available here without a reference to
' the Scripting runtime, and re-Adding an existing key raises 457, so the
' previous entry is removed first.
Private Function NextLaneOrdinal( _
    ByRef laneOrdinals As Collection, _
    ByVal lane As String) As Long

    ' Collection.Add rejects an empty key, and an unset Lane must not take
    ' the whole pass down over placement bookkeeping.
    Dim key As String
    key = lane
    If Len(key) = 0 Then key = "UnnamedLane"

    Dim current As Long
    current = 0

    On Error Resume Next
    current = CLng(laneOrdinals.Item(key))
    Err.Clear
    laneOrdinals.Remove key
    Err.Clear
    On Error GoTo 0

    current = current + 1
    laneOrdinals.Add current, key
    NextLaneOrdinal = current
End Function

' R23-808. A lane is a NAME, and this is the single place it becomes a
' sheet coordinate. The point is derived from the section view's own
' measured outline plus the existing layout margin - no new literal, and no
' fraction of the sheet. Auto-Arrange is enabled in the fixture profile and
' runs after this, so this is a starting point, not final placement; the
' proof string records exactly what it was derived from.
'
' R23-822. The ordinal is per lane, and the result is clamped into the
' proved usable area. r60 stacked a GLOBAL ordinal, so the second entry in
' LANE_ABOVE started 24 mm above the view top, outside the usable area, and
' auto-arrange finished the job: RD2 landed in the zoned border and
' regressed ANNOTATION_EXTENTS from PROVED to FAILED. An unbounded starting
' point is the defect - a lane is allowed to stack, but not off the drawing.
' If the usable bounds are not proved the ordinal stops stacking at 1 rather
' than growing against an unknown boundary.
Private Sub LaneTextPoint( _
    ByRef swView As SldWorks.View, _
    ByVal lane As String, _
    ByVal ordinal As Long, _
    ByRef evidence As CRunEvidence, _
    ByRef textX As Double, _
    ByRef textY As Double, _
    ByRef placementProof As String)

    On Error GoTo Failed

    placementProof = "placement=Unavailable"
    textX = 0#
    textY = 0#

    Dim outline As Variant
    outline = swView.GetOutline
    If Not IsArray(outline) Then Exit Sub

    Dim viewLeft As Double
    Dim viewBottom As Double
    Dim viewRight As Double
    Dim viewTop As Double
    viewLeft = CDbl(outline(0))
    viewBottom = CDbl(outline(1))
    viewRight = CDbl(outline(2))
    viewTop = CDbl(outline(3))

    Dim centreX As Double
    Dim centreY As Double
    centreX = (viewLeft + viewRight) / 2#
    centreY = (viewBottom + viewTop) / 2#

    Dim boundsProven As Boolean
    If Not evidence Is Nothing Then
        boundsProven = evidence.LayoutBoundariesProven
    End If

    Dim effectiveOrdinal As Long
    effectiveOrdinal = ordinal
    If effectiveOrdinal < 1 Then effectiveOrdinal = 1
    If Not boundsProven Then effectiveOrdinal = 1

    Dim gap As Double
    gap = Module8_RuntimeSupport.LAYOUT_MARGIN_M * CDbl(effectiveOrdinal)

    Select Case lane
        Case LANE_ABOVE
            textX = centreX
            textY = viewTop + gap
        Case LANE_BELOW
            textX = centreX
            textY = viewBottom - gap
        Case LANE_BORE_SIDE_A, LANE_EXTERIOR_VERTICAL_INNER
            textX = viewLeft - gap
            textY = centreY
        Case Else
            textX = viewRight + gap
            textY = centreY
    End Select

    Dim requestedX As Double
    Dim requestedY As Double
    requestedX = textX
    requestedY = textY

    Dim usableProof As String
    usableProof = "usable=Unproved"

    If boundsProven Then
        If textX < evidence.UsableLeft Then textX = evidence.UsableLeft
        If textX > evidence.UsableRight Then textX = evidence.UsableRight
        If textY < evidence.UsableBottom Then
            textY = evidence.UsableBottom
        End If
        If textY > evidence.UsableTop Then textY = evidence.UsableTop

        usableProof = "usable=" & FormatMetres(evidence.UsableLeft) & _
            "," & FormatMetres(evidence.UsableBottom) & "," & _
            FormatMetres(evidence.UsableRight) & "," & _
            FormatMetres(evidence.UsableTop)
    End If

    placementProof = "placement=ViewOutlinePlusMargin" & _
        "|outline=" & FormatMetres(viewLeft) & "," & _
        FormatMetres(viewBottom) & "," & FormatMetres(viewRight) & "," & _
        FormatMetres(viewTop) & _
        "|laneOrdinal=" & CStr(ordinal) & _
        "|appliedOrdinal=" & CStr(effectiveOrdinal) & _
        "|gapM=" & FormatMetres(gap) & _
        "|" & usableProof & _
        "|clamped=" & _
        CStr(requestedX <> textX Or requestedY <> textY) & _
        "|autoArrangeRunsAfter=True"
    Exit Sub

Failed:
    placementProof = "placement=Error:" & CStr(Err.Number)
End Sub

' One selectability verdict, named by requirement and by which side of the
' pair it came from.
Private Function ReportCandidateSelectability( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef entities As Variant, _
    ByVal viewName As String, _
    ByVal requirementKey As String, _
    ByVal side As String, _
    ByVal recordIndex As Long, _
    ByRef evidence As CRunEvidence) As Boolean

    If recordIndex < 0 Then Exit Function

    Dim route As String
    Dim selectable As Boolean
    selectable = ProveSectionEntitySelection( _
        swDrawModel, swView, entities, recordIndex, route)

    EmitInfo evidence, "SECTION_ENTITY_SELECT|view=" & viewName & _
        "|key=" & requirementKey & _
        "|side=" & side & _
        "|record=" & CStr(recordIndex) & _
        "|selectable=" & CStr(selectable) & _
        "|route=" & route & _
        "|source=IView.SelectEntity+GetSelectedObjectsDrawingView2"

    ReportCandidateSelectability = selectable
End Function

' R23-819. Route D, applied to a section view.
'
' Module13.SelectModelEntityInView proved this pattern for orthographic
' views: select through IView.SelectEntity, then prove the owning view
' through ISelectionMgr.GetSelectedObjectsDrawingView2 before trusting the
' object. Here the entity is already in hand - GetPolylines7's return value
' is the entity array, positionally paired with the polyline records, which
' r57 confirmed exactly (records=79, entities=79). The open question is
' whether those entries are live and selectable in a CUT view, where many
' curves are section faces with no model edge behind them and the Help says
' the array carries Null in their place.
'
' Read-only. Nothing is created; the selection is cleared on every path.
Private Function ProveSectionEntitySelection( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef entities As Variant, _
    ByVal recordIndex As Long, _
    ByRef route As String) As Boolean

    On Error GoTo Failed

    route = "NotAttempted"

    If swDrawModel Is Nothing Then
        route = "NoDrawingDocument"
        Exit Function
    End If
    If swView Is Nothing Then
        route = "NoView"
        Exit Function
    End If
    If Not IsArray(entities) Then
        route = "NoEntityArray"
        Exit Function
    End If

    Dim entityIndex As Long
    entityIndex = LBound(entities) + recordIndex
    If entityIndex > UBound(entities) Then
        route = "RecordIndexBeyondEntityArray"
        Exit Function
    End If

    Dim candidate As Object
    Set candidate = Nothing
    On Error Resume Next
    Set candidate = entities(entityIndex)
    Err.Clear
    On Error GoTo Failed

    ' The Help states the entity array carries Null where a polyline renders
    ' something no edge backs. In a section view that is expected for cut
    ' faces, so it is a measurement, not an error.
    If candidate Is Nothing Then
        route = "EntityIsNothing"
        Exit Function
    End If

    Dim selectionMgr As SldWorks.SelectionMgr
    Set selectionMgr = swDrawModel.SelectionManager
    If selectionMgr Is Nothing Then
        route = "NoSelectionManager"
        Exit Function
    End If

    ' Same refusal as Module13: SelectEntity with AppendFlag=False replaces
    ' whatever is selected, and a caller's selection is not ours to discard.
    Dim initialCount As Long
    initialCount = selectionMgr.GetSelectedObjectCount2(-1)
    If initialCount <> 0 Then
        route = "RefusedPreexistingSelection:count" & CStr(initialCount)
        Exit Function
    End If

    Dim selected As Boolean
    Dim selectError As Long
    On Error Resume Next
    selected = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swView.SelectEntity(candidate, False))
    selectError = Err.Number
    Err.Clear
    On Error GoTo Failed

    Dim selectedCount As Long
    selectedCount = selectionMgr.GetSelectedObjectCount2(-1)
    If Not selected Or selectedCount <> 1 Then
        route = "SelectEntityFailed:result" & CStr(selected) & _
            ":count" & CStr(selectedCount) & ":err" & CStr(selectError)
        GoTo CleanUp
    End If

    Dim owningView As SldWorks.View
    Set owningView = selectionMgr.GetSelectedObjectsDrawingView2(1, -1)
    If owningView Is Nothing Then
        route = "OwnerUnavailable"
        GoTo CleanUp
    End If

    If StrComp(Module11_GeometryIdentity.IdentityToken( _
        SafeViewName(owningView)), _
        Module11_GeometryIdentity.IdentityToken(SafeViewName(swView)), _
        vbBinaryCompare) <> 0 Then

        route = "WrongOwner:" & SafeViewName(owningView)
        GoTo CleanUp
    End If

    route = "Selectable|owner=" & SafeViewName(owningView)
    ProveSectionEntitySelection = True

CleanUp:
    swDrawModel.ClearSelection2 True
    Exit Function

Failed:
    Dim selectionErrorNumber As Long
    selectionErrorNumber = Err.Number
    route = "Error:" & CStr(selectionErrorNumber)
    On Error Resume Next
    swDrawModel.ClearSelection2 True
End Function

Private Function SafeItemCount(ByVal value As Variant) As Long
    On Error Resume Next
    If IsEmpty(value) Then Exit Function
    If Not IsArray(value) Then Exit Function
    SafeItemCount = UBound(value) - LBound(value) + 1
End Function

' IView.UseSheetScale returns 1 only when the view is tied to the SHEET
' scale. SOLIDWORKS 2025 Help states plainly: "If the property is 0, then
' it is possible that the view scale is the same as the sheet scale" - a
' section view inherits its PARENT view's scale (IView::UseParentScale),
' which reads 0 here at an identical rendered ratio. Both values are
' reported so the flag and the actual ratio can be told apart.
Private Function SafeScaleDecimal( _
    ByRef swView As SldWorks.View, _
    ByRef useSheetScale As Long) As Double

    On Error Resume Next
    useSheetScale = -1
    SafeScaleDecimal = 0#
    If swView Is Nothing Then Exit Function
    useSheetScale = swView.UseSheetScale
    SafeScaleDecimal = swView.ScaleDecimal
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
        Module21_EvidenceSink.LogLine _
            "R23_SECTIONDIM_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swDraw As SldWorks.ModelDoc2
    Set swDraw = swApp.ActiveDoc

    If swDraw Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_SECTIONDIM_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swDraw.GetType <> swDocDRAWING Then
        Module21_EvidenceSink.LogLine _
            "R23_SECTIONDIM_FATAL" & _
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
        Module21_EvidenceSink.LogLine _
            "R23_SECTIONDIM_FATAL|reason=NoViewsOnSheet"
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
            "R23_SECTIONDIM_FATAL|reason=NoReferencedDocument"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Module21_EvidenceSink.LogLine _
            "R23_SECTIONDIM_FATAL|reason=UnauthorizedFixture" & _
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

    Module21_EvidenceSink.LogLine _
        "R23_SECTIONDIM_BEGIN|drawing=" & swDraw.GetPathName & _
        "|part=" & partPath & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath) & _
        "|mode=ReadOnly|creations=0|mutations=0"

    Dim sectionViews As Collection
    Set sectionViews = CollectSectionViews(views)

    If sectionViews.Count = 0 Then
        Module21_EvidenceSink.LogLine _
            "R23_SECTIONDIM_FATAL|reason=NoSectionViewOnSheet"
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

        Module21_EvidenceSink.LogLine _
            "QA INFO: SECTION_VIEW|name=" & _
            SectionToken(viewName) & _
            "|type=" & CStr(swView.Type) & _
            "|excludedFromGenericArrangement=" & _
            CStr(IsExcludedFromGenericArrangement(swView))

        Dim dimensions As Collection
        Set dimensions = InventorySectionDimensions(swView, evidence)

        Dim requirements As Collection
        Set requirements = BuildSectionRequirements()

        ReconcileSectionDimensions requirements, dimensions, evidence

        ' ReconcileSectionDimensions already emitted one line per
        ' requirement through the evidence ledger, and CRunEvidence.AddInfo
        ' prints what it records. Printing them again here duplicated every
        ' requirement line in the first run's log.

        Dim verdict As String
        verdict = VerifySectionDimensions(requirements, dimensions)

        Module21_EvidenceSink.LogLine _
            "QA INFO: SECTION_DIMENSIONS|view=" & _
            SectionToken(viewName) & "|" & verdict

        Module21_EvidenceSink.LogLine _
            "QA INFO: SECTION_LEGACY|view=" & _
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

    Module21_EvidenceSink.LogLine _
        "R23_SECTIONDIM_END|sectionViews=" & _
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
    Module21_EvidenceSink.LogLine _
        "R23_SECTIONDIM_FATAL|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description

    On Error Resume Next
    If Not swDraw Is Nothing Then
        swDraw.SetPickMode
        swDraw.ClearSelection2 True
    End If
End Sub

' R23-823. Pulls the section view's annotation text back inside the proved
' usable area, AFTER everything that moves it has run.
'
' r61 proved that clamping at creation time cannot hold. The clamp was
' correct when it was applied - RD2 was placed at exactly UsableTop - but
' ArrangeViewsInMeasuredGrid moves the section view afterwards, and a
' dimension travels with its view. The move was
' (-0.097308, +0.014235) m, which carried RD2 from (0.304000, 0.275000) to
' (0.206692, 0.289235): past ContentBorderTop, region=ZonedBorder, and
' ANNOTATION_EXTENTS failed on that one line. The r60 diagnosis blamed
' auto-arrange; the r61 evidence shows auto-arrange never touched this view
' (ACTIVATE_VIEW|operation=Dimension arrange lists Drawing View1, Drawing
' View2 and Drawing View4 only). A box measured before the view is placed
' is stale by construction, so the correction belongs after placement.
'
' This is the fourth and last mutating procedure in the module. It creates
' nothing, deletes nothing, changes no dimension value, and moves no view -
' it sets annotation origins only, so the content-envelope repositioning and
' rescaling the 2026-08-04 user decision retired stay uncalled.
'
' IAnnotation::SetPosition2 takes sheet coordinates from the lower-left
' corner, the same frame IAnnotation::GetPosition returns, so the two can be
' compared directly. The Help says radial and diametric dimensions cannot be
' positioned this way because they already hang off a leader, and that a
' constrained annotation is placed "as near as possible" instead. Neither
' case is assumed away: the setter result is recorded, the position is read
' back, and the readback - not the request - decides whether the annotation
' ended up inside.
Public Function ClampSectionAnnotationsIntoUsableArea( _
    ByRef swView As SldWorks.View, _
    ByVal allowMutation As Boolean, _
    ByRef evidence As CRunEvidence) As Long

    On Error GoTo Failed

    If Not allowMutation Then
        EmitWarning evidence, "SECTION_ANNOTATION_CLAMP_REFUSED" & _
            "|reason=MutationNotAuthorized"
        Exit Function
    End If

    If swView Is Nothing Or evidence Is Nothing Then Exit Function

    Dim viewName As String
    viewName = SafeViewName(swView)

    ' Fail closed. Clamping to bounds nothing measured would move the text
    ' somewhere arbitrary, which is worse than leaving it where it is.
    If Not evidence.LayoutBoundariesProven Then
        EmitWarning evidence, "SECTION_ANNOTATION_CLAMP_REFUSED|view=" & _
            viewName & "|reason=LayoutBoundariesUnproved"
        Exit Function
    End If

    Dim annotations As Variant
    annotations = swView.GetAnnotations
    If Not IsArray(annotations) Then
        EmitInfo evidence, "SECTION_ANNOTATION_CLAMP_PASS|view=" & _
            viewName & "|annotations=0|moved=0|stillOutside=0"
        Exit Function
    End If

    Dim inspected As Long
    Dim moved As Long
    Dim stillOutside As Long

    Dim i As Long
    For i = LBound(annotations) To UBound(annotations)
        Dim annotation As SldWorks.Annotation
        Set annotation = Nothing
        On Error Resume Next
        Set annotation = annotations(i)
        Err.Clear
        On Error GoTo Failed

        If annotation Is Nothing Then GoTo ContinueAnnotation

        Dim position As Variant
        position = annotation.GetPosition
        If Not IsArray(position) Then GoTo ContinueAnnotation

        Dim lower As Long
        lower = LBound(position)
        If UBound(position) - lower + 1 < 3 Then GoTo ContinueAnnotation

        inspected = inspected + 1

        Dim currentX As Double
        Dim currentY As Double
        Dim currentZ As Double
        currentX = CDbl(position(lower))
        currentY = CDbl(position(lower + 1))
        currentZ = CDbl(position(lower + 2))

        Dim targetX As Double
        Dim targetY As Double
        targetX = ClampToRange( _
            currentX, evidence.UsableLeft, evidence.UsableRight)
        targetY = ClampToRange( _
            currentY, evidence.UsableBottom, evidence.UsableTop)

        Dim annotationName As String
        annotationName = SafeAnnotationName(annotation)

        If targetX = currentX And targetY = currentY Then
            EmitInfo evidence, "SECTION_ANNOTATION_CLAMP|view=" & _
                viewName & "|name=" & annotationName & _
                "|x=" & FormatMetres(currentX) & _
                "|y=" & FormatMetres(currentY) & _
                "|action=AlreadyInside"
            GoTo ContinueAnnotation
        End If

        Dim setterResult As Boolean
        setterResult = False
        On Error Resume Next
        setterResult = Module11_GeometryIdentity.NormalizeSwBoolean( _
            annotation.SetPosition2(targetX, targetY, currentZ))
        Err.Clear
        On Error GoTo Failed

        ' The readback is the verdict. "As near as possible" and the
        ' radial/diametric refusal both look like a successful call.
        Dim readbackX As Double
        Dim readbackY As Double
        readbackX = currentX
        readbackY = currentY

        Dim readback As Variant
        readback = annotation.GetPosition
        If IsArray(readback) Then
            Dim readbackLower As Long
            readbackLower = LBound(readback)
            If UBound(readback) - readbackLower + 1 >= 2 Then
                readbackX = CDbl(readback(readbackLower))
                readbackY = CDbl(readback(readbackLower + 1))
            End If
        End If

        Dim nowInside As Boolean
        nowInside = _
            (readbackX >= evidence.UsableLeft) And _
            (readbackX <= evidence.UsableRight) And _
            (readbackY >= evidence.UsableBottom) And _
            (readbackY <= evidence.UsableTop)

        If nowInside Then
            moved = moved + 1
        Else
            stillOutside = stillOutside + 1
        End If

        EmitInfo evidence, "SECTION_ANNOTATION_CLAMP|view=" & viewName & _
            "|name=" & annotationName & _
            "|fromX=" & FormatMetres(currentX) & _
            "|fromY=" & FormatMetres(currentY) & _
            "|requestX=" & FormatMetres(targetX) & _
            "|requestY=" & FormatMetres(targetY) & _
            "|setterResult=" & CStr(setterResult) & _
            "|readbackX=" & FormatMetres(readbackX) & _
            "|readbackY=" & FormatMetres(readbackY) & _
            "|nowInside=" & CStr(nowInside)

ContinueAnnotation:
    Next i

    EmitInfo evidence, "SECTION_ANNOTATION_CLAMP_PASS|view=" & viewName & _
        "|annotations=" & CStr(inspected) & _
        "|moved=" & CStr(moved) & _
        "|stillOutside=" & CStr(stillOutside) & _
        "|usable=" & FormatMetres(evidence.UsableLeft) & "," & _
        FormatMetres(evidence.UsableBottom) & "," & _
        FormatMetres(evidence.UsableRight) & "," & _
        FormatMetres(evidence.UsableTop)

    ClampSectionAnnotationsIntoUsableArea = moved
    Exit Function

Failed:
    ' Captured BEFORE SafeViewName, whose On Error Resume Next resets Err
    ' and would report every clamp failure as error 0.
    Dim clampErrorNumber As Long
    clampErrorNumber = Err.Number
    EmitFailure evidence, "SECTION_ANNOTATION_CLAMP_ERROR|view=" & _
        SafeViewName(swView) & "|error=" & CStr(clampErrorNumber)
End Function

Private Function ClampToRange( _
    ByVal value As Double, _
    ByVal lowerBound As Double, _
    ByVal upperBound As Double) As Double

    ClampToRange = value
    If ClampToRange < lowerBound Then ClampToRange = lowerBound
    If ClampToRange > upperBound Then ClampToRange = upperBound
End Function

Private Function SafeAnnotationName( _
    ByRef annotation As SldWorks.Annotation) As String

    On Error Resume Next
    SafeAnnotationName = annotation.GetName
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    If Len(SafeAnnotationName) = 0 Then
        SafeAnnotationName = "Unnamed"
    End If
End Function

' R23 probe-runner compile-failure localisation. A no-op; VBA compiles
' at module granularity, so a module that loads this has compiled.
Public Sub R23_CompileTouch()
End Sub
