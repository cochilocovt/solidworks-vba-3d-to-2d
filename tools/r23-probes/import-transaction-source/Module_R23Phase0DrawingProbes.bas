Option Explicit

' R23 Phase 0 drawing-contract probes.
' These procedures create and mutate only a fresh, unsaved drawing.
' They do not modify, rebuild, or save the referenced fixture model.

Private Const R23_PROBE_BUILD As String = _
    "20260731.2-mapping-frame-h7-contracts"
Private Const R23_LOG_DIRECTORY As String = _
    "C:\Users\V.T\Documents\VBA 3D TO 2D\test_assets\iteration_evidence\r23\20260730-075811\live-probes"

' Verified against installed SOLIDWORKS 2025 SP1.2 interop 33.1.2.4.
Private Const swDocPART As Long = 1
Private Const swDocDRAWING As Long = 3
Private Const swDrawingSectionView As Long = 2
Private Const swDrawingNamedView As Long = 7
Private Const swVerticalOrdinate As Long = 2
Private Const swHorizontalOrdinate As Long = 3
Private Const swCreateOrdDimErr_Success As Long = 0
Private Const swSolidBody As Long = 0
Private Const swThisConfiguration As Long = 1
Private Const swDiameterDimension As Long = 6
Private Const swDiametricLinearDimension As Long = 15
Private Const swLinearDimension As Long = 2
Private Const swViewEntityType_Edge As Long = 1
Private Const swObjectSameValue As Long = 1

Private Const LOCATION_TOLERANCE_M As Double = 0.00005
Private Const NOMINAL_TOLERANCE_M As Double = 0.0000001
Private Const PAGE_MATCH_TOLERANCE_M As Double = 0.0000005

Private mR23DrawingProbeLogPath As String

' Captured while parsing GetSectionLineInfo2 so the page-frame conversion and
' clearance checks read exactly the parsed payload values.
Private mParsedSegmentPoints As Collection
Private mParsedArrowPoints As Collection
Private mParsedLabelPoints As Collection
Private mParsedLabelTextHeight As Double

' Page-frame reserved regions gathered for the clearance comparison.
Private mContentRegionValid As Boolean
Private mContentLeft As Double
Private mContentBottom As Double
Private mContentRight As Double
Private mContentTop As Double
Private mTitleRegionValid As Boolean
Private mTitleLeft As Double
Private mTitleBottom As Double
Private mTitleRight As Double
Private mTitleTop As Double
Private mPartIdRegionValid As Boolean
Private mPartIdLeft As Double
Private mPartIdBottom As Double
Private mPartIdRight As Double
Private mPartIdTop As Double

Public Sub R23_ProbeDatumFirstXYOrdinates()
    RunDatumFirstOrdinateProbe
End Sub

Public Sub R23_ProbeSectionDimensionsAndJJGeometry()
    RunSectionDimensionProbe
End Sub

Private Sub RunDatumFirstOrdinateProbe()
    On Error GoTo Failed

    Dim swApp As SldWorks.SldWorks
    Dim swPart As SldWorks.ModelDoc2
    Dim swDrawModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Dim primaryView As SldWorks.View
    Dim sectionView As SldWorks.View
    Dim modelSaveFlagBefore As Boolean

    If Not PrepareDisposableProbeDrawing( _
        "ORDINATE", swApp, swPart, swDrawModel, swDraw, _
        primaryView, sectionView, modelSaveFlagBefore) Then

        GoTo SafeExit
    End If

    Dim component As SldWorks.Component2
    Set component = GetSingleVisibleComponent(primaryView)

    If component Is Nothing Then
        ProbeLog "R23_ORDINATE_FATAL|reason=SingleVisibleComponentNotProved"
        GoTo SafeExit
    End If

    ProbeLog "R23_ORDINATE_COMPONENT" & _
        "|view=" & ProbeToken(SafeViewName(primaryView)) & _
        "|component=" & ProbeToken(component.Name2) & _
        "|source=IView.GetVisibleComponents"

    ' Route C inventory: drawing edges visible for the limited component.
    Dim visibleEdges As Variant
    visibleEdges = GetVisibleEdgeInventory(swApp, primaryView, component)

    Dim featureEntities As New Collection
    Dim featureXValues As New Collection
    Dim featureYValues As New Collection

    If Not CollectFeatureCylinderLocations( _
        swApp, swPart, primaryView, component, visibleEdges, _
        "CBORE for M6 Socket Head Cap Screw1", _
        featureEntities, featureXValues, featureYValues) Then

        ProbeLog "R23_ORDINATE_FATAL" & _
            "|reason=CounterboreLocationsUnavailable"
        GoTo SafeExit
    End If

    ProbeLog "R23_ORDINATE_LOCATION_SUMMARY" & _
        "|family=CBORE_M6" & _
        "|uniqueProjectedLocations=" & CStr(featureEntities.Count)

    If featureEntities.Count <> 6 Then
        ProbeLog "R23_ORDINATE_FATAL" & _
            "|reason=ExpectedSixCounterboreLocations" & _
            "|actual=" & CStr(featureEntities.Count)
        GoTo SafeExit
    End If

    Dim centerEntities As New Collection
    Dim centerXValues As New Collection
    Dim centerYValues As New Collection

    If Not CollectFeatureCylinderLocations( _
        swApp, swPart, primaryView, component, visibleEdges, _
        "Cut-Extrude1", _
        centerEntities, centerXValues, centerYValues) Then

        ProbeLog "R23_ORDINATE_FATAL" & _
            "|reason=CenterDatumCylinderUnavailable"
        GoTo SafeExit
    End If

    If centerEntities.Count < 1 Then
        ProbeLog "R23_ORDINATE_FATAL" & _
            "|reason=CenterDatumEntityUnavailable"
        GoTo SafeExit
    End If

    Dim centerDatum As SldWorks.Entity
    Set centerDatum = centerEntities(1)

    Dim bottomDatum As SldWorks.Entity
    Dim bottomX As Double
    Dim bottomY As Double

    Set bottomDatum = FindBottomLeftMappedVertex( _
        swApp, swPart, primaryView, component, bottomX, bottomY)

    If bottomDatum Is Nothing Then
        ProbeLog "R23_ORDINATE_FATAL" & _
            "|reason=BottomLeftDatumUnavailable"
        GoTo SafeExit
    End If

    ProbeLog "R23_ORDINATE_DATUM" & _
        "|direction=X" & _
        "|source=Cut-Extrude1Cylinder" & _
        "|viewX=" & FormatProbeNumber(CDbl(centerXValues(1))) & _
        "|viewY=" & FormatProbeNumber(CDbl(centerYValues(1)))

    ProbeLog "R23_ORDINATE_DATUM" & _
        "|direction=Y" & _
        "|source=BottomLeftMappedModelVertex" & _
        "|viewX=" & FormatProbeNumber(bottomX) & _
        "|viewY=" & FormatProbeNumber(bottomY)

    Dim xEntities As Collection
    Dim xCoordinates As Collection
    Set xEntities = New Collection
    Set xCoordinates = New Collection

    BuildUniqueDirectionSelection _
        featureEntities, featureXValues, True, xEntities, xCoordinates

    Dim yEntities As Collection
    Dim yCoordinates As Collection
    Set yEntities = New Collection
    Set yCoordinates = New Collection

    BuildUniqueDirectionSelection _
        featureEntities, featureYValues, False, yEntities, yCoordinates

    ProbeLog "R23_ORDINATE_DIRECTION_SUMMARY" & _
        "|xUnique=" & CStr(xEntities.Count) & _
        "|yUnique=" & CStr(yEntities.Count)

    If xEntities.Count <> 2 Or yEntities.Count <> 3 Then
        ProbeLog "R23_ORDINATE_FATAL" & _
            "|reason=UnexpectedCanonicalDirectionCounts" & _
            "|x=" & CStr(xEntities.Count) & _
            "|y=" & CStr(yEntities.Count)
        GoTo SafeExit
    End If

    Dim xCreated As Boolean
    Dim yCreated As Boolean

    xCreated = CreateDatumFirstOrdinateGroup( _
        swDrawModel, swDraw, primaryView, centerDatum, _
        xEntities, xCoordinates, True, "P0251-FACE-X")

    yCreated = CreateDatumFirstOrdinateGroup( _
        swDrawModel, swDraw, primaryView, bottomDatum, _
        yEntities, yCoordinates, False, "P0251-FACE-Y")

    DumpOrdinateReadback primaryView

    ProbeLog "R23_ORDINATE_RESULT" & _
        "|xCreated=" & CStr(xCreated) & _
        "|yCreated=" & CStr(yCreated) & _
        "|status=" & IIf(xCreated And yCreated, "SUCCESS", "FAILED")

SafeExit:
    FinishDrawingProbe _
        "ORDINATE", swPart, swDrawModel, swDraw, modelSaveFlagBefore
    Exit Sub

Failed:
    ProbeLog "R23_ORDINATE_FATAL|reason=UnhandledError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
    Resume SafeExit
End Sub

Private Sub RunSectionDimensionProbe()
    On Error GoTo Failed

    Dim swApp As SldWorks.SldWorks
    Dim swPart As SldWorks.ModelDoc2
    Dim swDrawModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Dim primaryView As SldWorks.View
    Dim sectionView As SldWorks.View
    Dim modelSaveFlagBefore As Boolean

    If Not PrepareDisposableProbeDrawing( _
        "SECTION", swApp, swPart, swDrawModel, swDraw, _
        primaryView, sectionView, modelSaveFlagBefore) Then

        GoTo SafeExit
    End If

    If sectionView Is Nothing Then
        ProbeLog "R23_SECTION_FATAL|reason=SectionViewUnavailable"
        GoTo SafeExit
    End If

    ' Direct part-source tolerance readback for the H7 authority decision.
    ' This reads the fixture model; it never modifies it.
    Dim sourceH7Proven As Boolean
    sourceH7Proven = DumpPartSourceDimensionAuthority(swPart)

    Dim diameter47Found As Boolean
    Dim diameter40Found As Boolean
    Dim h7Proven As Boolean
    Dim linearFound As Boolean
    Dim diametricLinearFound As Boolean
    Dim exactTargetCounts As Boolean

    DumpSectionDimensionContracts _
        sectionView, diameter47Found, diameter40Found, h7Proven, _
        linearFound, diametricLinearFound, exactTargetCounts

    Dim sectionGeometryParsed As Boolean
    sectionGeometryParsed = _
        DumpSectionLineGeometry(swDraw, primaryView, sectionView)

    ProbeLog "R23_SECTION_RESULT" & _
        "|diameter47Found=" & CStr(diameter47Found) & _
        "|diameter40Found=" & CStr(diameter40Found) & _
        "|exactTargetCounts=" & CStr(exactTargetCounts) & _
        "|h7ProvenInDrawing=" & CStr(h7Proven) & _
        "|h7ProvenInPartSource=" & CStr(sourceH7Proven) & _
        "|linearFound=" & CStr(linearFound) & _
        "|diametricLinearFound=" & CStr(diametricLinearFound) & _
        "|sectionGeometryParsed=" & CStr(sectionGeometryParsed) & _
        "|evidenceStatus=" & _
            IIf(diameter47Found And diameter40Found And _
                exactTargetCounts And _
                linearFound And sectionGeometryParsed, _
                "CAPTURED", "INCOMPLETE")

SafeExit:
    FinishDrawingProbe _
        "SECTION", swPart, swDrawModel, swDraw, modelSaveFlagBefore
    Exit Sub

Failed:
    ProbeLog "R23_SECTION_FATAL|reason=UnhandledError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
    Resume SafeExit
End Sub

Private Function PrepareDisposableProbeDrawing( _
    ByVal probeName As String, _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef primaryView As SldWorks.View, _
    ByRef sectionView As SldWorks.View, _
    ByRef modelSaveFlagBefore As Boolean) As Boolean

    On Error GoTo Failed

    Dim preparationErrorNumber As Long
    Dim preparationErrorDescription As String

    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        Debug.Print "R23_DRAWING_PROBE_FATAL|reason=SolidWorksUnavailable"
        Exit Function
    End If

    Set swPart = swApp.ActiveDoc
    If swPart Is Nothing Then
        Debug.Print "R23_DRAWING_PROBE_FATAL|reason=NoActiveDocument"
        Exit Function
    End If

    If swPart.GetType <> swDocPART Then
        Debug.Print "R23_DRAWING_PROBE_FATAL|reason=ActiveDocumentNotPart"
        Exit Function
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Module1_Main.GetFixtureKey(partPath) <> "P-0251-14A-001" Then
        Debug.Print "R23_DRAWING_PROBE_FATAL" & _
            "|reason=AuthorizedP0251Required" & _
            "|path=" & ProbeToken(partPath)
        Exit Function
    End If

    StartProbeLog probeName
    modelSaveFlagBefore = NormalizeSwBoolean(swPart.GetSaveFlag)

    ProbeLog "R23_DRAWING_PROBE_BEGIN" & _
        "|build=" & R23_PROBE_BUILD & _
        "|name=" & ProbeToken(probeName) & _
        "|part=" & ProbeToken(partPath) & _
        "|configuration=" & ProbeToken( _
            swPart.ConfigurationManager.ActiveConfiguration.Name) & _
        "|solidWorksRevision=" & ProbeToken(swApp.RevisionNumber) & _
        "|importVariant=SelectedViewsFalse" & _
        "|drawingDisposable=True" & _
        "|modelSaveFlagBefore=" & CStr(modelSaveFlagBefore) & _
        "|logPath=" & ProbeToken(mR23DrawingProbeLogPath)

    Dim templatePath As String
    templatePath = Module1_Main.GetValidDrawingTemplatePath(swApp)

    If Len(templatePath) = 0 Then
        ProbeLog "R23_DRAWING_PROBE_FATAL" & _
            "|reason=ControlledTemplateMissing"
        Exit Function
    End If

    Set Module1_Main.GlobalEvidence = New CRunEvidence
    Module1_Main.ResetGlobalConfig
    Module1_Main.ApplyFixtureAcceptanceProfile partPath

    With Module1_Main.GlobalConfig
        .AutoArrange = False
        .PopulateTitle = False
        .InsertBarcode = False
        .InsertNotes = False
        .GenerateQAReport = False
        .ShowLayoutPreview = False
        .Cancelled = False
    End With

    Module1_Main.R23_InitializeProbeEvidence _
        swApp, swPart, partPath, templatePath
    Module1_Main.GlobalEvidence.TemplatePath = templatePath

    Module4_ModelItemImporter.R23_ConfigureImportProbe _
        "SelectedViewsFalse"

    Module2_DrawingPipeline.CreateDrawing _
        swApp, swPart, partPath, templatePath, _
        Module1_Main.GlobalEvidence

    Module4_ModelItemImporter.R23_ClearImportProbe

    Set swDrawModel = swApp.ActiveDoc
    If swDrawModel Is Nothing Then
        ProbeLog "R23_DRAWING_PROBE_FATAL" & _
            "|reason=NoActiveDrawingAfterCreation"
        Exit Function
    End If

    If swDrawModel.GetType <> swDocDRAWING Then
        ProbeLog "R23_DRAWING_PROBE_FATAL" & _
            "|reason=ActiveDocumentNotGeneratedDrawing"
        Exit Function
    End If

    Set swDraw = swDrawModel
    If swDraw Is Nothing Then
        ProbeLog "R23_DRAWING_PROBE_FATAL" & _
            "|reason=DrawingInterfaceUnavailable"
        Exit Function
    End If

    Set primaryView = FindUniqueView(swDraw, "PRIMARY")
    Set sectionView = FindUniqueView(swDraw, "SECTION")

    If primaryView Is Nothing Then
        ProbeLog "R23_DRAWING_PROBE_FATAL" & _
            "|reason=PrimaryViewUnavailable"
        Exit Function
    End If

    ProbeLog "R23_DRAWING_PROBE_SCAFFOLD" & _
        "|drawing=" & ProbeToken(swDrawModel.GetTitle) & _
        "|primary=" & ProbeToken(SafeViewName(primaryView)) & _
        "|section=" & ProbeToken(SafeViewName(sectionView)) & _
        "|displayDimensions=" & CStr( _
            Module4_ModelItemImporter.CountAllDisplayDimensions(swDraw))

    PrepareDisposableProbeDrawing = True
    Exit Function

Failed:
    preparationErrorNumber = Err.Number
    preparationErrorDescription = Err.Description

    On Error Resume Next
    Module4_ModelItemImporter.R23_ClearImportProbe
    ProbeLog "R23_DRAWING_PROBE_FATAL|reason=PreparationError" & _
        "|error=" & CStr(preparationErrorNumber) & _
        "|description=" & ProbeToken(preparationErrorDescription)
    On Error GoTo 0
End Function

Private Function FindUniqueView( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal roleName As String) As SldWorks.View

    Dim matchCount As Long
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        Dim matches As Boolean

        If roleName = "SECTION" Then
            matches = (swView.Type = swDrawingSectionView)
        ElseIf roleName = "PRIMARY" Then
            On Error Resume Next
            matches = (swView.Type = swDrawingNamedView And _
                StrComp(swView.GetOrientationName, "*Front", _
                    vbTextCompare) = 0)
            On Error GoTo 0
        End If

        If matches Then
            matchCount = matchCount + 1
            If matchCount = 1 Then Set FindUniqueView = swView
        End If

        Set swView = swView.GetNextView
    Loop

    If matchCount <> 1 Then Set FindUniqueView = Nothing
End Function

Private Function GetSingleVisibleComponent( _
    ByRef swView As SldWorks.View) As SldWorks.Component2

    On Error GoTo Failed

    Dim visible As Variant
    visible = swView.GetVisibleComponents

    If IsEmpty(visible) Or Not IsArray(visible) Then Exit Function

    Dim count As Long
    Dim i As Long

    For i = LBound(visible) To UBound(visible)
        Dim component As SldWorks.Component2
        Set component = visible(i)

        If Not component Is Nothing Then
            count = count + 1
            If count = 1 Then Set GetSingleVisibleComponent = component
        End If
    Next i

    If count <> 1 Then Set GetSingleVisibleComponent = Nothing
    Exit Function

Failed:
    Set GetSingleVisibleComponent = Nothing
End Function

Private Function CollectFeatureCylinderLocations( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef component As SldWorks.Component2, _
    ByVal visibleEdges As Variant, _
    ByVal featureName As String, _
    ByRef drawingEntities As Collection, _
    ByRef xValues As Collection, _
    ByRef yValues As Collection) As Boolean

    On Error GoTo Failed

    Dim swFeature As SldWorks.Feature
    Set swFeature = FindFeatureByExactName(swPart, featureName)

    If swFeature Is Nothing Then
        ProbeLog "R23_ORDINATE_GEOMETRY_REJECT" & _
            "|feature=" & ProbeToken(featureName) & _
            "|reason=FeatureNotFound"
        Exit Function
    End If

    Dim faces As Variant
    faces = swFeature.GetFaces

    If IsEmpty(faces) Or Not IsArray(faces) Then
        ProbeLog "R23_ORDINATE_GEOMETRY_REJECT" & _
            "|feature=" & ProbeToken(featureName) & _
            "|reason=FacesUnavailable"
        Exit Function
    End If

    ProbeLog "R23_ORDINATE_FEATURE" & _
        "|feature=" & ProbeToken(featureName) & _
        "|ownedFaces=" & CStr(VariantItemCount(faces)) & _
        "|ownershipSource=IFeature.GetFaces"

    Dim faceIndex As Long
    For faceIndex = LBound(faces) To UBound(faces)
        Dim face As SldWorks.Face2
        Set face = faces(faceIndex)

        If face Is Nothing Then
            ProbeLog "R23_ORDINATE_FACE" & _
                "|feature=" & ProbeToken(featureName) & _
                "|faceIndex=" & CStr(faceIndex) & _
                "|verdict=Reject|reason=FaceIsNothing"
        Else
            InspectFeatureFace _
                swApp, swView, component, visibleEdges, _
                featureName, faceIndex, face, _
                drawingEntities, xValues, yValues
        End If
    Next faceIndex

    ProbeLog "R23_ORDINATE_ROUTE_SUMMARY" & _
        "|feature=" & ProbeToken(featureName) & _
        "|acceptedUniqueLocations=" & CStr(drawingEntities.Count)

    CollectFeatureCylinderLocations = (drawingEntities.Count > 0)
    Exit Function

Failed:
    ProbeLog "R23_ORDINATE_GEOMETRY_REJECT" & _
        "|feature=" & ProbeToken(featureName) & _
        "|reason=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Function

Private Sub InspectFeatureFace( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef component As SldWorks.Component2, _
    ByVal visibleEdges As Variant, _
    ByVal featureName As String, _
    ByVal faceIndex As Long, _
    ByRef face As SldWorks.Face2, _
    ByRef drawingEntities As Collection, _
    ByRef xValues As Collection, _
    ByRef yValues As Collection)

    On Error GoTo Failed

    Dim surface As SldWorks.Surface
    Set surface = face.GetSurface

    If surface Is Nothing Then
        ProbeLog "R23_ORDINATE_FACE" & _
            "|feature=" & ProbeToken(featureName) & _
            "|faceIndex=" & CStr(faceIndex) & _
            "|verdict=Reject|reason=SurfaceIsNothing"
        Exit Sub
    End If

    ' SOLIDWORKS COM Booleans must be normalized with CBool before negation.
    ' A raw VARIANT_BOOL of 1 makes "Not value" evaluate to -2, which VBA
    ' treats as True, so the 2026-07-31 run rejected all 18 owned faces as
    ' NotCylindrical even though the accepted feature probe reads the same
    ' faces as cylinders through CBool(ISurface.IsCylinder).  The raw value
    ' is logged so this defect class cannot hide in a future transcript.
    Dim rawIsCylinder As Variant
    rawIsCylinder = surface.IsCylinder

    Dim isCylindrical As Boolean
    isCylindrical = NormalizeSwBoolean(rawIsCylinder)

    If Not isCylindrical Then
        ProbeLog "R23_ORDINATE_FACE" & _
            "|feature=" & ProbeToken(featureName) & _
            "|faceIndex=" & CStr(faceIndex) & _
            "|isCylinderRaw=" & ProbeToken(CStr(rawIsCylinder)) & _
            "|isCylinder=" & CStr(isCylindrical) & _
            "|verdict=Reject|reason=NotCylindrical"
        Exit Sub
    End If

    Dim cylinderData As Variant
    cylinderData = surface.CylinderParams

    If VariantItemCount(cylinderData) <> 7 Then
        ProbeLog "R23_ORDINATE_FACE" & _
            "|feature=" & ProbeToken(featureName) & _
            "|faceIndex=" & CStr(faceIndex) & _
            "|isCylinderRaw=" & ProbeToken(CStr(rawIsCylinder)) & _
            "|isCylinder=True|verdict=Reject" & _
            "|reason=CylinderParamsItemCount:" & _
                CStr(VariantItemCount(cylinderData))
        Exit Sub
    End If

    Dim baseIndex As Long
    baseIndex = LBound(cylinderData)

    Dim radiusM As Double
    radiusM = Abs(CDbl(cylinderData(baseIndex + 6)))

    Dim edges As Variant
    edges = face.GetEdges

    ProbeLog "R23_ORDINATE_FACE" & _
        "|feature=" & ProbeToken(featureName) & _
        "|faceIndex=" & CStr(faceIndex) & _
        "|isCylinderRaw=" & ProbeToken(CStr(rawIsCylinder)) & _
        "|isCylinder=True" & _
        "|radiusM=" & FormatProbeNumber(radiusM) & _
        "|edgeCount=" & CStr(VariantItemCount(edges)) & _
        "|verdict=Candidate"

    Dim pageX As Double
    Dim pageY As Double
    Dim pageZ As Double
    Dim transformProof As String
    Dim transformOk As Boolean

    transformOk = Module8_RuntimeSupport.TransformPointToView( _
        swApp, swView, _
        CDbl(cylinderData(baseIndex)), _
        CDbl(cylinderData(baseIndex + 1)), _
        CDbl(cylinderData(baseIndex + 2)), _
        pageX, pageY, pageZ, transformProof, True)

    ProbeLog "R23_ORDINATE_FACE_TRANSFORM" & _
        "|feature=" & ProbeToken(featureName) & _
        "|faceIndex=" & CStr(faceIndex) & _
        "|frame=Page" & _
        "|result=" & CStr(transformOk) & _
        "|pageX=" & FormatProbeNumber(pageX) & _
        "|pageY=" & FormatProbeNumber(pageY) & _
        "|proof=" & ProbeToken(transformProof)

    If Not transformOk Then Exit Sub

    If IsEmpty(edges) Or Not IsArray(edges) Then
        ProbeLog "R23_ORDINATE_FACE" & _
            "|feature=" & ProbeToken(featureName) & _
            "|faceIndex=" & CStr(faceIndex) & _
            "|verdict=Reject|reason=EdgesUnavailable"
        Exit Sub
    End If

    Dim edgeIndex As Long
    For edgeIndex = LBound(edges) To UBound(edges)
        Dim modelEdge As SldWorks.Edge
        Set modelEdge = edges(edgeIndex)

        Dim mappedEntity As SldWorks.Entity
        Set mappedEntity = EvaluateModelEdgeMapping( _
            swApp, swView, component, visibleEdges, _
            featureName, faceIndex, edgeIndex, modelEdge)

        If Not mappedEntity Is Nothing Then
            AddUniqueLocation _
                mappedEntity, pageX, pageY, _
                drawingEntities, xValues, yValues

            ProbeLog "R23_ORDINATE_GEOMETRY" & _
                "|feature=" & ProbeToken(featureName) & _
                "|faceIndex=" & CStr(faceIndex) & _
                "|edgeIndex=" & CStr(edgeIndex) & _
                "|radiusM=" & FormatProbeNumber(radiusM) & _
                "|frame=Page" & _
                "|pageX=" & FormatProbeNumber(pageX) & _
                "|pageY=" & FormatProbeNumber(pageY) & _
                "|circleProof=IsCircleAndClosedCurveParams"
            Exit Sub
        End If
    Next edgeIndex
    Exit Sub

Failed:
    ProbeLog "R23_ORDINATE_FACE" & _
        "|feature=" & ProbeToken(featureName) & _
        "|faceIndex=" & CStr(faceIndex) & _
        "|verdict=Reject|reason=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

Private Function FindFeatureByExactName( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal featureName As String) As SldWorks.Feature

    Dim swFeature As SldWorks.Feature
    Set swFeature = swPart.FirstFeature

    Do While Not swFeature Is Nothing
        If StrComp(swFeature.Name, featureName, vbTextCompare) = 0 Then
            Set FindFeatureByExactName = swFeature
            Exit Function
        End If

        Set swFeature = swFeature.GetNextFeature
    Loop
End Function

' Compares the three candidate model-edge to drawing-entity mapping routes and
' logs one complete evidence record per attempted edge.  Qualification is not
' weakened: only a feature-owned, complete circular edge with an
' ownership-backed mapped drawing entity (route A or route B) is accepted.
Private Function EvaluateModelEdgeMapping( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef component As SldWorks.Component2, _
    ByVal visibleEdges As Variant, _
    ByVal featureName As String, _
    ByVal faceIndex As Long, _
    ByVal edgeIndex As Long, _
    ByRef modelEdge As SldWorks.Edge) As SldWorks.Entity

    On Error GoTo Failed

    Dim edgePrefix As String
    edgePrefix = "R23_ORDINATE_EDGE" & _
        "|feature=" & ProbeToken(featureName) & _
        "|faceIndex=" & CStr(faceIndex) & _
        "|edgeIndex=" & CStr(edgeIndex)

    If modelEdge Is Nothing Then
        ProbeLog edgePrefix & "|verdict=Reject|reason=EdgeIsNothing"
        Exit Function
    End If

    Dim isCircleResult As String
    Dim paramsAvailable As Boolean
    Dim uMinValue As Double
    Dim uMaxValue As Double
    Dim endpointGapM As Double
    Dim rejectGate As String
    Dim closureToleranceM As Double
    Dim circleParamsSummary As String
    Dim completeCircle As Boolean

    completeCircle = ReadEdgeCircleEvidence( _
        modelEdge, isCircleResult, paramsAvailable, _
        uMinValue, uMaxValue, endpointGapM, _
        rejectGate, closureToleranceM, circleParamsSummary)

    ProbeLog edgePrefix & _
        "|edgeType=" & ProbeToken(TypeName(modelEdge)) & _
        "|isCircle=" & ProbeToken(isCircleResult) & _
        "|curveParams3Available=" & CStr(paramsAvailable) & _
        "|uMin=" & FormatProbeNumber(uMinValue) & _
        "|uMax=" & FormatProbeNumber(uMaxValue) & _
        "|endpointGapM=" & FormatProbeNumber(endpointGapM) & _
        "|closureToleranceM=" & FormatProbeNumber(closureToleranceM) & _
        "|circleParams=" & ProbeToken(circleParamsSummary) & _
        "|rejectGate=" & ProbeToken(rejectGate) & _
        "|completeCircle=" & CStr(completeCircle)

    If Not completeCircle Then
        ProbeLog edgePrefix & _
            "|verdict=Reject|reason=NotCompleteCircle" & _
            "|gate=" & ProbeToken(rejectGate)
        Exit Function
    End If

    ' Route A: active-part model edge straight into the drawing view.
    Dim directError As Long
    Dim directMapped As SldWorks.Entity
    Set directMapped = TryGetCorrespondingViewEntity( _
        swView, modelEdge, directError)

    ProbeLog edgePrefix & _
        "|route=A:IView.GetCorrespondingEntity(modelEdge)" & _
        "|result=" & DescribeMappedObject(directMapped) & _
        "|error=" & CStr(directError)

    ' Route B: model edge into the component context, then into the view.
    Dim componentError As Long
    Dim componentEntity As SldWorks.Entity
    Set componentEntity = TryGetComponentEntity( _
        component, modelEdge, componentError)

    ProbeLog edgePrefix & _
        "|route=B1:IComponent2.GetCorrespondingEntity(modelEdge)" & _
        "|result=" & DescribeMappedObject(componentEntity) & _
        "|error=" & CStr(componentError)

    Dim viaComponentError As Long
    Dim viaComponentMapped As SldWorks.Entity
    If Not componentEntity Is Nothing Then
        Set viaComponentMapped = TryGetCorrespondingViewEntity( _
            swView, componentEntity, viaComponentError)

        ProbeLog edgePrefix & _
            "|route=B2:IView.GetCorrespondingEntity(componentEntity)" & _
            "|result=" & DescribeMappedObject(viaComponentMapped) & _
            "|error=" & CStr(viaComponentError)
    End If

    If Not directMapped Is Nothing And _
       Not viaComponentMapped Is Nothing Then

        ProbeLog edgePrefix & _
            "|routeIdentity=AvsB" & _
            "|swObjectEquality=" & _
                CStr(SafeIsSame(swApp, directMapped, viaComponentMapped))
    End If

    ' Route C: relate the raw model edge and any mapped drawing entity to the
    ' visible drawing-edge inventory by COM identity.
    ProbeLog edgePrefix & _
        "|route=C:GetVisibleEntities2(component,Edge)" & _
        "|modelEdgeVisibleIndex=" & _
            CStr(FindVisibleEdgeIndex(swApp, visibleEdges, modelEdge))

    Dim chosenEntity As SldWorks.Entity
    Dim chosenRoute As String

    If Not directMapped Is Nothing Then
        Set chosenEntity = directMapped
        chosenRoute = "A:Direct"
    ElseIf Not viaComponentMapped Is Nothing Then
        Set chosenEntity = viaComponentMapped
        chosenRoute = "B:ComponentMediated"
    End If

    If chosenEntity Is Nothing Then
        ProbeLog edgePrefix & _
            "|verdict=Reject|reason=NoOwnershipBackedDrawingEntity"
        Exit Function
    End If

    ProbeLog edgePrefix & _
        "|chosenRoute=" & ProbeToken(chosenRoute) & _
        "|chosenType=" & ProbeToken(TypeName(chosenEntity)) & _
        "|chosenVisibleIndex=" & _
            CStr(FindVisibleEdgeIndex(swApp, visibleEdges, chosenEntity)) & _
        "|verdict=Accept"

    Set EvaluateModelEdgeMapping = chosenEntity
    Exit Function

Failed:
    ProbeLog "R23_ORDINATE_EDGE" & _
        "|feature=" & ProbeToken(featureName) & _
        "|faceIndex=" & CStr(faceIndex) & _
        "|edgeIndex=" & CStr(edgeIndex) & _
        "|verdict=Reject|reason=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
    Set EvaluateModelEdgeMapping = Nothing
End Function

Private Function TryGetCorrespondingViewEntity( _
    ByRef swView As SldWorks.View, _
    ByVal modelEntity As Object, _
    ByRef errorNumber As Long) As SldWorks.Entity

    errorNumber = 0
    On Error GoTo Failed

    Dim mapped As Object
    Set mapped = swView.GetCorrespondingEntity(modelEntity)
    If mapped Is Nothing Then Exit Function

    Set TryGetCorrespondingViewEntity = mapped
    Exit Function

Failed:
    errorNumber = Err.Number
    Set TryGetCorrespondingViewEntity = Nothing
End Function

Private Function TryGetComponentEntity( _
    ByRef component As SldWorks.Component2, _
    ByVal modelEntity As Object, _
    ByRef errorNumber As Long) As SldWorks.Entity

    errorNumber = 0
    On Error GoTo Failed

    If component Is Nothing Then Exit Function

    Dim mapped As Object
    Set mapped = component.GetCorrespondingEntity(modelEntity)
    If mapped Is Nothing Then Exit Function

    Set TryGetComponentEntity = mapped
    Exit Function

Failed:
    errorNumber = Err.Number
    Set TryGetComponentEntity = Nothing
End Function

Private Function DescribeMappedObject(ByVal mapped As Object) As String
    On Error GoTo Failed

    If mapped Is Nothing Then
        DescribeMappedObject = "Nothing"
    Else
        DescribeMappedObject = ProbeToken(TypeName(mapped))
    End If
    Exit Function

Failed:
    DescribeMappedObject = "Unreadable:" & CStr(Err.Number)
End Function

' Returns swObjectEquality: 0 not same, 1 same, 2 unsupported, -1 call failed.
Private Function SafeIsSame( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal firstObject As Object, _
    ByVal secondObject As Object) As Long

    On Error GoTo Failed
    SafeIsSame = swApp.IsSame(firstObject, secondObject)
    Exit Function

Failed:
    SafeIsSame = -1
End Function

Private Function FindVisibleEdgeIndex( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal visibleEdges As Variant, _
    ByVal candidate As Object) As Long

    FindVisibleEdgeIndex = -1
    On Error GoTo Failed

    If candidate Is Nothing Then Exit Function
    If IsEmpty(visibleEdges) Or Not IsArray(visibleEdges) Then Exit Function

    Dim i As Long
    For i = LBound(visibleEdges) To UBound(visibleEdges)
        Dim visibleEdge As Object
        Set visibleEdge = visibleEdges(i)

        If Not visibleEdge Is Nothing Then
            If SafeIsSame(swApp, candidate, visibleEdge) = _
               swObjectSameValue Then

                FindVisibleEdgeIndex = i
                Exit Function
            End If
        End If
    Next i
    Exit Function

Failed:
    FindVisibleEdgeIndex = -1
End Function

Private Function GetVisibleEdgeInventory( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef component As SldWorks.Component2) As Variant

    GetVisibleEdgeInventory = Empty
    On Error GoTo Failed

    Dim visibleEdges As Variant
    visibleEdges = swView.GetVisibleEntities2( _
        component, swViewEntityType_Edge)

    If IsEmpty(visibleEdges) Or Not IsArray(visibleEdges) Then
        ProbeLog "R23_ORDINATE_VISIBLE_EDGES" & _
            "|view=" & ProbeToken(SafeViewName(swView)) & _
            "|count=0|status=NoArrayReturned"
        Exit Function
    End If

    Dim totalCount As Long
    Dim circularCount As Long
    Dim i As Long

    totalCount = VariantItemCount(visibleEdges)

    For i = LBound(visibleEdges) To UBound(visibleEdges)
        Dim visibleEdge As SldWorks.Edge
        Set visibleEdge = Nothing
        On Error Resume Next
        Set visibleEdge = visibleEdges(i)
        On Error GoTo Failed

        If Not visibleEdge Is Nothing Then
            Dim isCircleResult As String
            Dim paramsAvailable As Boolean
            Dim uMinValue As Double
            Dim uMaxValue As Double
            Dim endpointGapM As Double
            Dim rejectGate As String
            Dim closureToleranceM As Double
            Dim circleParamsSummary As String
            Dim completeCircle As Boolean

            completeCircle = ReadEdgeCircleEvidence( _
                visibleEdge, isCircleResult, paramsAvailable, _
                uMinValue, uMaxValue, endpointGapM, _
                rejectGate, closureToleranceM, circleParamsSummary)

            If completeCircle Then circularCount = circularCount + 1

            ProbeLog "R23_ORDINATE_VISIBLE_EDGE" & _
                "|index=" & CStr(i) & _
                "|type=" & ProbeToken(TypeName(visibleEdge)) & _
                "|isCircle=" & ProbeToken(isCircleResult) & _
                "|curveParams3Available=" & CStr(paramsAvailable) & _
                "|uMin=" & FormatProbeNumber(uMinValue) & _
                "|uMax=" & FormatProbeNumber(uMaxValue) & _
                "|endpointGapM=" & FormatProbeNumber(endpointGapM) & _
                "|closureToleranceM=" & _
                    FormatProbeNumber(closureToleranceM) & _
                "|rejectGate=" & ProbeToken(rejectGate) & _
                "|completeCircle=" & CStr(completeCircle)
        End If
    Next i

    ProbeLog "R23_ORDINATE_VISIBLE_EDGES" & _
        "|view=" & ProbeToken(SafeViewName(swView)) & _
        "|component=" & ProbeToken(component.Name2) & _
        "|count=" & CStr(totalCount) & _
        "|completeCircles=" & CStr(circularCount) & _
        "|source=IView.GetVisibleEntities2"

    GetVisibleEdgeInventory = visibleEdges
    Exit Function

Failed:
    ProbeLog "R23_ORDINATE_VISIBLE_EDGES" & _
        "|view=" & ProbeToken(SafeViewName(swView)) & _
        "|status=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
    GetVisibleEdgeInventory = Empty
End Function

' Retains the accepted proof: ICurve.IsCircle=True plus a closed
' IEdge.GetCurveParams3 boundary.  Returns each intermediate observation so
' rejected edges leave a usable evidence trail.
Private Function ReadEdgeCircleEvidence( _
    ByRef modelEdge As SldWorks.Edge, _
    ByRef isCircleResult As String, _
    ByRef paramsAvailable As Boolean, _
    ByRef uMinValue As Double, _
    ByRef uMaxValue As Double, _
    ByRef endpointGapM As Double, _
    ByRef rejectGate As String, _
    ByRef closureToleranceM As Double, _
    ByRef circleParamsSummary As String) As Boolean

    isCircleResult = "NotRead"
    paramsAvailable = False
    uMinValue = 0#
    uMaxValue = 0#
    endpointGapM = -1#
    rejectGate = "NotEvaluated"
    closureToleranceM = Module8_RuntimeSupport.GEOMETRY_TOLERANCE_M
    circleParamsSummary = "NotRead"

    If modelEdge Is Nothing Then
        isCircleResult = "EdgeIsNothing"
        rejectGate = "EdgeIsNothing"
        Exit Function
    End If

    On Error GoTo Failed

    Dim curve As SldWorks.Curve
    Set curve = modelEdge.GetCurve
    If curve Is Nothing Then
        isCircleResult = "CurveIsNothing"
        rejectGate = "CurveIsNothing"
        Exit Function
    End If

    Dim rawIsCircle As Variant
    rawIsCircle = curve.IsCircle

    Dim circleFlag As Boolean
    circleFlag = NormalizeSwBoolean(rawIsCircle)
    isCircleResult = CStr(circleFlag) & _
        "(raw=" & ProbeToken(CStr(rawIsCircle)) & ")"

    ' ICurve.CircleParams is read for evidence only.  The historical
    ' "SkippedNotCircle" result that barred it from production came from an
    ' "If Not <SOLIDWORKS Boolean>" guard, not from the API, so its actual
    ' behaviour has never been observed on this fixture.
    circleParamsSummary = ReadCircleParamsSummary(curve, circleFlag)

    Dim curveParameters As SldWorks.CurveParamData
    Set curveParameters = modelEdge.GetCurveParams3
    If curveParameters Is Nothing Then
        rejectGate = "CurveParamsNothing"
        Exit Function
    End If

    paramsAvailable = True
    uMinValue = curveParameters.UMinValue
    uMaxValue = curveParameters.UMaxValue

    Dim startPoint As Variant
    Dim endPoint As Variant
    startPoint = curveParameters.StartPoint
    endPoint = curveParameters.EndPoint

    endpointGapM = PointDistance(startPoint, endPoint)

    ' Each gate is a block If and records the exact rejection reason.  The
    ' 2026-07-31 run returned False for every edge, including perfect
    ' circles reporting UMin=0, UMax=2pi and a zero endpoint gap, so the
    ' rejecting gate must be identified rather than inferred.  The previous
    ' revision ended this chain with a single-line "If ... Then _" line
    ' continuation immediately followed by the success assignment, which is
    ' the only construct here that can swallow that assignment.
    If Not circleFlag Then
        rejectGate = "IsCircleFalse"
        Exit Function
    End If

    If uMaxValue <= uMinValue Then
        rejectGate = "ParameterSpanInvalid"
        Exit Function
    End If

    If endpointGapM < 0# Then
        rejectGate = "EndpointGapUnreadable"
        Exit Function
    End If

    If endpointGapM > closureToleranceM Then
        rejectGate = "EndpointGapExceedsTolerance"
        Exit Function
    End If

    rejectGate = "None"
    ReadEdgeCircleEvidence = True
    Exit Function

Failed:
    isCircleResult = "ReadError:" & CStr(Err.Number)
    rejectGate = "ReadError:" & CStr(Err.Number)
    ReadEdgeCircleEvidence = False
End Function

' The only reliable way to consume a SOLIDWORKS COM Boolean in this VBA host.
'
' The installed build returns VARIANT_BOOL values whose True is not VBA's -1.
' Such a value prints as "True" through CStr and behaves correctly in
' "If value Then" and "If value = False Then", but "Not value" yields -2,
' which VBA treats as True.  CBool applied directly to the method-call
' expression does not reliably normalize it either: the 2026-07-31 runs show
' CBool(rawVariant) working for ISurface.IsCylinder while
' CBool(curve.IsCircle) still produced a value that failed "Not".
'
' Converting through CDbl and comparing with zero is representation
' independent and therefore cannot exhibit either failure mode.
Private Function NormalizeSwBoolean(ByVal rawValue As Variant) As Boolean
    On Error GoTo Failed

    If IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function
    NormalizeSwBoolean = (CDbl(rawValue) <> 0#)
    Exit Function

Failed:
    NormalizeSwBoolean = False
End Function

' Evidence-only read of ICurve.CircleParams.  Never load bearing.
Private Function ReadCircleParamsSummary( _
    ByRef curve As SldWorks.Curve, _
    ByVal circleFlag As Boolean) As String

    If Not circleFlag Then
        ReadCircleParamsSummary = "NotCircle"
        Exit Function
    End If

    On Error GoTo Failed

    Dim circleData As Variant
    circleData = curve.CircleParams

    If Not IsArray(circleData) Then
        ReadCircleParamsSummary = "NotArray"
        Exit Function
    End If

    Dim itemCount As Long
    itemCount = VariantItemCount(circleData)

    If itemCount <> 7 Then
        ReadCircleParamsSummary = "ItemCount:" & CStr(itemCount)
        Exit Function
    End If

    ReadCircleParamsSummary = "Count7:radiusM=" & _
        FormatProbeNumber(Abs(CDbl(circleData(LBound(circleData) + 6))))
    Exit Function

Failed:
    ReadCircleParamsSummary = "ReadError:" & CStr(Err.Number)
End Function

' Returns the distance between two 3-value points, or -1 when unreadable.
Private Function PointDistance( _
    ByVal firstPoint As Variant, _
    ByVal secondPoint As Variant) As Double

    PointDistance = -1#
    On Error GoTo Failed

    If Not IsArray(firstPoint) Or Not IsArray(secondPoint) Then Exit Function
    If VariantItemCount(firstPoint) < 3 Or _
       VariantItemCount(secondPoint) < 3 Then Exit Function

    Dim firstIndex As Long
    Dim secondIndex As Long
    firstIndex = LBound(firstPoint)
    secondIndex = LBound(secondPoint)

    PointDistance = Sqr( _
        (CDbl(firstPoint(firstIndex)) - _
         CDbl(secondPoint(secondIndex))) ^ 2 + _
        (CDbl(firstPoint(firstIndex + 1)) - _
         CDbl(secondPoint(secondIndex + 1))) ^ 2 + _
        (CDbl(firstPoint(firstIndex + 2)) - _
         CDbl(secondPoint(secondIndex + 2))) ^ 2)
    Exit Function

Failed:
    PointDistance = -1#
End Function

Private Sub AddUniqueLocation( _
    ByRef drawingEntity As SldWorks.Entity, _
    ByVal viewX As Double, _
    ByVal viewY As Double, _
    ByRef drawingEntities As Collection, _
    ByRef xValues As Collection, _
    ByRef yValues As Collection)

    Dim i As Long
    For i = 1 To xValues.Count
        If Abs(CDbl(xValues(i)) - viewX) <= LOCATION_TOLERANCE_M And _
           Abs(CDbl(yValues(i)) - viewY) <= LOCATION_TOLERANCE_M Then

            Exit Sub
        End If
    Next i

    drawingEntities.Add drawingEntity
    xValues.Add viewX
    yValues.Add viewY
End Sub

Private Function FindBottomLeftMappedVertex( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swView As SldWorks.View, _
    ByRef component As SldWorks.Component2, _
    ByRef resultX As Double, _
    ByRef resultY As Double) As SldWorks.Entity

    On Error GoTo Failed

    Dim swPartDoc As SldWorks.PartDoc
    Set swPartDoc = swPart

    Dim bodies As Variant
    bodies = swPartDoc.GetBodies2(swSolidBody, True)
    If IsEmpty(bodies) Or Not IsArray(bodies) Then Exit Function

    Dim found As Boolean
    Dim bestX As Double
    Dim bestY As Double
    Dim bestEntity As SldWorks.Entity
    Dim bestRoute As String

    Dim consideredCount As Long
    Dim transformRejectedCount As Long
    Dim routeAMappedCount As Long
    Dim routeBMappedCount As Long
    Dim unmappedCount As Long

    Dim bodyIndex As Long
    For bodyIndex = LBound(bodies) To UBound(bodies)
        Dim body As SldWorks.Body2
        Set body = bodies(bodyIndex)

        If Not body Is Nothing Then
            Dim edges As Variant
            edges = body.GetEdges

            If IsArray(edges) Then
                Dim edgeIndex As Long
                For edgeIndex = LBound(edges) To UBound(edges)
                    Dim modelEdge As SldWorks.Edge
                    Set modelEdge = edges(edgeIndex)

                    If Not modelEdge Is Nothing Then
                        Dim startVertex As SldWorks.Vertex
                        Dim endVertex As SldWorks.Vertex
                        Set startVertex = modelEdge.GetStartVertex
                        Set endVertex = modelEdge.GetEndVertex

                        ConsiderBottomLeftVertex _
                            swApp, swView, component, startVertex, _
                            found, bestX, bestY, bestEntity, bestRoute, _
                            consideredCount, transformRejectedCount, _
                            routeAMappedCount, routeBMappedCount, _
                            unmappedCount
                        ConsiderBottomLeftVertex _
                            swApp, swView, component, endVertex, _
                            found, bestX, bestY, bestEntity, bestRoute, _
                            consideredCount, transformRejectedCount, _
                            routeAMappedCount, routeBMappedCount, _
                            unmappedCount
                    End If
                Next edgeIndex
            End If
        End If
    Next bodyIndex

    ProbeLog "R23_ORDINATE_VERTEX_SUMMARY" & _
        "|considered=" & CStr(consideredCount) & _
        "|transformRejected=" & CStr(transformRejectedCount) & _
        "|routeAMapped=" & CStr(routeAMappedCount) & _
        "|routeBMapped=" & CStr(routeBMappedCount) & _
        "|unmapped=" & CStr(unmappedCount) & _
        "|found=" & CStr(found)

    If found Then
        ProbeLog "R23_ORDINATE_VERTEX_CHOSEN" & _
            "|frame=Page" & _
            "|pageX=" & FormatProbeNumber(bestX) & _
            "|pageY=" & FormatProbeNumber(bestY) & _
            "|route=" & ProbeToken(bestRoute) & _
            "|entityType=" & ProbeToken(TypeName(bestEntity))

        resultX = bestX
        resultY = bestY
        Set FindBottomLeftMappedVertex = bestEntity
    Else
        Set FindBottomLeftMappedVertex = Nothing
    End If
    Exit Function

Failed:
    ProbeLog "R23_ORDINATE_VERTEX_SUMMARY" & _
        "|status=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
    Set FindBottomLeftMappedVertex = Nothing
End Function

Private Sub ConsiderBottomLeftVertex( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swView As SldWorks.View, _
    ByRef component As SldWorks.Component2, _
    ByRef modelVertex As SldWorks.Vertex, _
    ByRef found As Boolean, _
    ByRef bestX As Double, _
    ByRef bestY As Double, _
    ByRef bestEntity As SldWorks.Entity, _
    ByRef bestRoute As String, _
    ByRef consideredCount As Long, _
    ByRef transformRejectedCount As Long, _
    ByRef routeAMappedCount As Long, _
    ByRef routeBMappedCount As Long, _
    ByRef unmappedCount As Long)

    If modelVertex Is Nothing Then Exit Sub
    On Error GoTo Failed

    consideredCount = consideredCount + 1

    Dim pointData As Variant
    pointData = modelVertex.GetPoint
    If VariantItemCount(pointData) < 3 Then Exit Sub

    Dim viewX As Double
    Dim viewY As Double
    Dim viewZ As Double
    Dim transformProof As String

    If Not Module8_RuntimeSupport.TransformPointToView( _
        swApp, swView, _
        CDbl(pointData(LBound(pointData))), _
        CDbl(pointData(LBound(pointData) + 1)), _
        CDbl(pointData(LBound(pointData) + 2)), _
        viewX, viewY, viewZ, transformProof, True) Then

        transformRejectedCount = transformRejectedCount + 1
        Exit Sub
    End If

    Dim mappedError As Long
    Dim vertexRoute As String
    Dim drawingEntity As SldWorks.Entity
    Set drawingEntity = TryGetCorrespondingViewEntity( _
        swView, modelVertex, mappedError)

    If Not drawingEntity Is Nothing Then
        vertexRoute = "A:Direct"
        routeAMappedCount = routeAMappedCount + 1
    Else
        Dim componentEntity As SldWorks.Entity
        Set componentEntity = TryGetComponentEntity( _
            component, modelVertex, mappedError)

        If Not componentEntity Is Nothing Then
            Set drawingEntity = TryGetCorrespondingViewEntity( _
                swView, componentEntity, mappedError)
        End If

        If Not drawingEntity Is Nothing Then
            vertexRoute = "B:ComponentMediated"
            routeBMappedCount = routeBMappedCount + 1
        End If
    End If

    If drawingEntity Is Nothing Then
        unmappedCount = unmappedCount + 1
        Exit Sub
    End If

    If Not found Or _
       viewY < bestY - LOCATION_TOLERANCE_M Or _
       (Abs(viewY - bestY) <= LOCATION_TOLERANCE_M And _
        viewX < bestX) Then

        found = True
        bestX = viewX
        bestY = viewY
        bestRoute = vertexRoute
        Set bestEntity = drawingEntity
    End If
    Exit Sub

Failed:
    ' One unmappable vertex does not invalidate other mapped vertices.
End Sub

Private Sub BuildUniqueDirectionSelection( _
    ByRef sourceEntities As Collection, _
    ByRef sourceCoordinates As Collection, _
    ByVal useX As Boolean, _
    ByRef resultEntities As Collection, _
    ByRef resultCoordinates As Collection)

    Dim i As Long
    For i = 1 To sourceEntities.Count
        Dim coordinateValue As Double
        coordinateValue = CDbl(sourceCoordinates(i))

        If Not CoordinateAlreadyPresent( _
            resultCoordinates, coordinateValue) Then

            Dim entity As SldWorks.Entity
            Set entity = sourceEntities(i)

            resultEntities.Add entity
            resultCoordinates.Add coordinateValue

            ProbeLog "R23_ORDINATE_CANONICAL_ITEM" & _
                "|direction=" & IIf(useX, "X", "Y") & _
                "|order=" & CStr(resultEntities.Count) & _
                "|coordinateM=" & FormatProbeNumber(coordinateValue)
        End If
    Next i
End Sub

Private Function CoordinateAlreadyPresent( _
    ByRef coordinates As Collection, _
    ByVal value As Double) As Boolean

    Dim i As Long
    For i = 1 To coordinates.Count
        If Abs(CDbl(coordinates(i)) - value) <= _
           LOCATION_TOLERANCE_M Then

            CoordinateAlreadyPresent = True
            Exit Function
        End If
    Next i
End Function

Private Function CreateDatumFirstOrdinateGroup( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByRef datumEntity As SldWorks.Entity, _
    ByRef featureEntities As Collection, _
    ByRef coordinates As Collection, _
    ByVal useX As Boolean, _
    ByVal groupKey As String) As Boolean

    On Error GoTo Failed

    Dim outcome As String
    Dim resultCode As Long
    Dim cleanupCount As Long
    Dim datumSelected As Boolean
    Dim appendedCount As Long
    Dim finalSelectionCount As Long
    Dim groupStep As String

    outcome = "NotStarted"
    resultCode = -9999
    cleanupCount = -1
    groupStep = "Entry"

    ' Normalized before negation for the same reason as ISurface.IsCylinder.
    groupStep = "ActivateView"
    Dim viewActivated As Boolean
    viewActivated = NormalizeSwBoolean( _
        swDraw.ActivateView(SafeViewName(swView)))

    If Not viewActivated Then
        outcome = "ViewActivationFailed"
        GoTo SafeExit
    End If

    groupStep = "SetPickModeAndClear"
    swDrawModel.SetPickMode
    swDrawModel.ClearSelection2 True

    groupStep = "AcquireSelectionManager"
    Dim selectionManager As SldWorks.SelectionMgr
    Set selectionManager = swDrawModel.SelectionManager

    If selectionManager Is Nothing Then
        outcome = "SelectionManagerNothing"
        GoTo SafeExit
    End If

    groupStep = "CreateSelectData"
    Dim selectData As SldWorks.SelectData
    Set selectData = selectionManager.CreateSelectData

    If selectData Is Nothing Then
        outcome = "CreateSelectDataNothing"
        GoTo SafeExit
    End If

    ' Assigning ISelectData.View raises runtime error 91 in this installed VBA
    ' host.  Module2_DrawingPipeline already records the same behaviour for
    ' section sketch segments and works around it by activating the source
    ' view first and proving ownership after Select4.  The assignment is
    ' attempted here and its outcome recorded, because the API documents the
    ' property as get/set; when it fails the group continues with unbound
    ' selection data and each selection's owning view is verified instead.
    groupStep = "BindSelectDataView"
    Dim viewBindingResult As String
    viewBindingResult = TryBindSelectDataView(selectData, swView)

    ProbeLog "R23_ORDINATE_SELECTDATA" & _
        "|group=" & ProbeToken(groupKey) & _
        "|viewBinding=" & ProbeToken(viewBindingResult) & _
        "|activatedView=" & ProbeToken(SafeViewName(swView)) & _
        "|contract=ISelectData.View"

    groupStep = "SelectDatum"
    datumSelected = NormalizeSwBoolean( _
        datumEntity.Select4(False, selectData))

    ProbeLog "R23_ORDINATE_SELECTION_OWNER" & _
        "|group=" & ProbeToken(groupKey) & _
        "|order=0|role=Datum" & _
        "|ownerView=" & ProbeToken( _
            SafeSelectedObjectViewName(selectionManager, 1))

    ProbeLog "R23_ORDINATE_SELECTION" & _
        "|group=" & ProbeToken(groupKey) & _
        "|order=0|role=Datum" & _
        "|selectResult=" & CStr(datumSelected) & _
        "|selectionCount=" & CStr( _
            swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)) & _
        "|selectionType=" & CStr( _
            swDrawModel.SelectionManager.GetSelectedObjectType3(1, -1))

    If Not datumSelected Then
        outcome = "DatumSelectionFailed"
        GoTo SafeExit
    End If

    groupStep = "AppendFeatureSelections"
    Dim i As Long
    For i = 1 To featureEntities.Count
        Dim entity As SldWorks.Entity
        Set entity = featureEntities(i)

        Dim appendResult As Boolean
        appendResult = NormalizeSwBoolean( _
            entity.Select4(True, selectData))

        Dim selectionCount As Long
        selectionCount = selectionManager.GetSelectedObjectCount2(-1)

        ProbeLog "R23_ORDINATE_SELECTION" & _
            "|group=" & ProbeToken(groupKey) & _
            "|order=" & CStr(i) & _
            "|role=Feature" & _
            "|coordinateM=" & _
                FormatProbeNumber(CDbl(coordinates(i))) & _
            "|selectResult=" & CStr(appendResult) & _
            "|selectionCount=" & CStr(selectionCount) & _
            "|selectionType=" & CStr( _
                selectionManager.GetSelectedObjectType3( _
                    selectionCount, -1)) & _
            "|ownerView=" & ProbeToken( _
                SafeSelectedObjectViewName( _
                    selectionManager, selectionCount))

        If Not appendResult Or selectionCount <> i + 1 Then
            outcome = "AppendSelectionFailed"
            GoTo SafeExit
        End If

        appendedCount = appendedCount + 1
    Next i

    groupStep = "ReadFinalSelectionCount"
    finalSelectionCount = selectionManager.GetSelectedObjectCount2(-1)

    groupStep = "ReadViewOutline"
    Dim outline As Variant
    outline = swView.GetOutline
    If VariantItemCount(outline) <> 4 Then
        outcome = "ViewOutlineUnavailable"
        GoTo SafeExit
    End If

    Dim outlineBase As Long
    outlineBase = LBound(outline)

    Dim ordinateType As Long
    Dim locationX As Double
    Dim locationY As Double

    If useX Then
        ordinateType = swHorizontalOrdinate
        locationX = _
            (CDbl(outline(outlineBase)) + _
             CDbl(outline(outlineBase + 2))) / 2#
        locationY = CDbl(outline(outlineBase + 1)) - 0.012
    Else
        ordinateType = swVerticalOrdinate
        locationX = CDbl(outline(outlineBase)) - 0.012
        locationY = _
            (CDbl(outline(outlineBase + 1)) + _
             CDbl(outline(outlineBase + 3))) / 2#
    End If

    Dim beforeDimensions As Long
    beforeDimensions = CountDisplayDimensions(swView)

    groupStep = "AddOrdinateDimension"
    resultCode = swDrawModel.Extension.AddOrdinateDimension( _
        ordinateType, locationX, locationY, 0#)

    Dim afterDimensions As Long
    afterDimensions = CountDisplayDimensions(swView)

    ProbeLog "R23_ORDINATE_API" & _
        "|group=" & ProbeToken(groupKey) & _
        "|direction=" & IIf(useX, "X", "Y") & _
        "|ordinateType=" & CStr(ordinateType) & _
        "|resultCode=" & CStr(resultCode) & _
        "|result=" & ProbeToken(DecodeOrdinateResult(resultCode)) & _
        "|displayDimensionsBefore=" & CStr(beforeDimensions) & _
        "|displayDimensionsAfter=" & CStr(afterDimensions) & _
        "|locationX=" & FormatProbeNumber(locationX) & _
        "|locationY=" & FormatProbeNumber(locationY)

    If resultCode <> swCreateOrdDimErr_Success Then
        outcome = "AddOrdinateFailed"
        GoTo SafeExit
    End If

    outcome = "Created"
    CreateDatumFirstOrdinateGroup = True

SafeExit:
    On Error Resume Next
    swDrawModel.SetPickMode
    swDrawModel.ClearSelection2 True
    cleanupCount = _
        swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)
    On Error GoTo 0

    ProbeLog "R23_ORDINATE_GROUP" & _
        "|group=" & ProbeToken(groupKey) & _
        "|direction=" & IIf(useX, "X", "Y") & _
        "|datumSelected=" & CStr(datumSelected) & _
        "|appended=" & CStr(appendedCount) & _
        "|expectedAppended=" & CStr(featureEntities.Count) & _
        "|finalSelection=" & CStr(finalSelectionCount) & _
        "|expectedFinal=" & CStr(featureEntities.Count + 1) & _
        "|resultCode=" & CStr(resultCode) & _
        "|cleanupSelection=" & CStr(cleanupCount) & _
        "|setPickModeCalled=True" & _
        "|lastStep=" & ProbeToken(groupStep) & _
        "|outcome=" & ProbeToken(outcome)
    Exit Function

Failed:
    outcome = "RuntimeError:" & CStr(Err.Number) & _
        "@" & groupStep
    ProbeLog "R23_ORDINATE_GROUP_ERROR" & _
        "|group=" & ProbeToken(groupKey) & _
        "|step=" & ProbeToken(groupStep) & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
    Resume SafeExit
End Function

' ISelectData.View is documented as a get/set property, but assigning it
' raises runtime error 91 in this installed VBA host.  The assignment is
' attempted and its outcome reported so the log states which contract the run
' actually used; the caller activates the view first either way and proves
' each selection's owning view afterwards.
Private Function TryBindSelectDataView( _
    ByRef selectData As SldWorks.SelectData, _
    ByRef swView As SldWorks.View) As String

    On Error GoTo Failed

    Set selectData.View = swView
    TryBindSelectDataView = "Bound"
    Exit Function

Failed:
    TryBindSelectDataView = _
        "UnboundAfterError:" & CStr(Err.Number)
End Function

Private Function SafeSelectedObjectViewName( _
    ByRef selectionManager As SldWorks.SelectionMgr, _
    ByVal selectionIndex As Long) As String

    On Error GoTo Failed

    If selectionManager Is Nothing Then
        SafeSelectedObjectViewName = "SelectionManagerNothing"
        Exit Function
    End If

    If selectionIndex < 1 Then
        SafeSelectedObjectViewName = "NoSelection"
        Exit Function
    End If

    Dim ownerView As SldWorks.View
    Set ownerView = _
        selectionManager.GetSelectedObjectsDrawingView2(selectionIndex, -1)

    If ownerView Is Nothing Then
        SafeSelectedObjectViewName = "Nothing"
        Exit Function
    End If

    SafeSelectedObjectViewName = SafeViewName(ownerView)
    Exit Function

Failed:
    SafeSelectedObjectViewName = "ReadError:" & CStr(Err.Number)
End Function

Private Sub DumpOrdinateReadback(ByRef swView As SldWorks.View)
    On Error GoTo Failed

    Dim dimensions As Variant
    dimensions = swView.GetDisplayDimensions

    If IsEmpty(dimensions) Or Not IsArray(dimensions) Then
        ProbeLog "R23_ORDINATE_READBACK|count=0"
        Exit Sub
    End If

    Dim ordinateCount As Long
    Dim i As Long

    For i = LBound(dimensions) To UBound(dimensions)
        Dim displayDimension As SldWorks.DisplayDimension
        Set displayDimension = dimensions(i)

        If Not displayDimension Is Nothing Then
            If displayDimension.Type2 = 7 Or _
               displayDimension.Type2 = 8 Or _
               displayDimension.Type2 = 1 Then

                ordinateCount = ordinateCount + 1

                Dim annotation As SldWorks.Annotation
                Set annotation = displayDimension.GetAnnotation

                ProbeLog "R23_ORDINATE_READBACK_ITEM" & _
                    "|index=" & CStr(ordinateCount) & _
                    "|type2=" & CStr(displayDimension.Type2) & _
                    "|name=" & ProbeToken( _
                        SafeDisplayDimensionName(displayDimension)) & _
                    "|position=" & ArrayToken( _
                        SafeAnnotationPosition(annotation), 8)
            End If
        End If
    Next i

    ProbeLog "R23_ORDINATE_READBACK" & _
        "|count=" & CStr(ordinateCount) & _
        "|view=" & ProbeToken(SafeViewName(swView))
    Exit Sub

Failed:
    ProbeLog "R23_ORDINATE_READBACK_ERROR" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

Private Sub DumpSectionDimensionContracts( _
    ByRef sectionView As SldWorks.View, _
    ByRef diameter47Found As Boolean, _
    ByRef diameter40Found As Boolean, _
    ByRef h7Proven As Boolean, _
    ByRef linearFound As Boolean, _
    ByRef diametricLinearFound As Boolean, _
    ByRef exactTargetCounts As Boolean)

    On Error GoTo Failed

    Dim dimensions As Variant
    dimensions = sectionView.GetDisplayDimensions

    If IsEmpty(dimensions) Or Not IsArray(dimensions) Then
        ProbeLog "R23_SECTION_DIMENSION_SUMMARY|count=0"
        Exit Sub
    End If

    Dim i As Long
    Dim dimensionCount As Long
    Dim firstLinearGeometryDumped As Boolean
    Dim diameter47Count As Long
    Dim diameter40Count As Long

    ' VBA block declarations are procedure-scoped, so every per-dimension
    ' value is explicitly reset at the top of each iteration.  Stale values
    ' from a previous iteration corrupted the 2026-07-31 run's labels.
    Dim dimension As SldWorks.Dimension
    Dim nominalM As Double
    Dim nominalAvailable As Boolean
    Dim targetName As String
    Dim dimensionH7Proven As Boolean
    Dim toleranceSummary As String

    For i = LBound(dimensions) To UBound(dimensions)
        Dim displayDimension As SldWorks.DisplayDimension
        Set displayDimension = dimensions(i)

        Set dimension = Nothing
        nominalM = 0#
        nominalAvailable = False
        targetName = vbNullString
        dimensionH7Proven = False
        toleranceSummary = vbNullString

        If Not displayDimension Is Nothing Then
            dimensionCount = dimensionCount + 1

            Set dimension = displayDimension.GetDimension2(0)

            nominalAvailable = _
                TryReadNominal(dimension, nominalM)

            If displayDimension.Type2 = swLinearDimension Then
                linearFound = True
            ElseIf displayDimension.Type2 = _
                swDiametricLinearDimension Then

                diametricLinearFound = True
            End If

            ' The live-proven imported diameters are standard
            ' swDiameterDimension=6 records; a semantic diameter target
            ' therefore requires both the nominal and type 6.  Type 15 is
            ' recorded but not required.
            If nominalAvailable And _
               displayDimension.Type2 = swDiameterDimension And _
               Abs(nominalM - 0.047) <= NOMINAL_TOLERANCE_M Then

                diameter47Count = diameter47Count + 1
                If diameter47Count = 1 Then
                    diameter47Found = True
                    targetName = "DIAMETER_47"
                Else
                    targetName = "DIAMETER_47_DUPLICATE"
                End If
            ElseIf nominalAvailable And _
                   displayDimension.Type2 = swDiameterDimension And _
                   Abs(nominalM - 0.04) <= NOMINAL_TOLERANCE_M Then

                diameter40Count = diameter40Count + 1
                If diameter40Count = 1 Then
                    diameter40Found = True
                    targetName = "DIAMETER_40"
                Else
                    targetName = "DIAMETER_40_DUPLICATE"
                End If
            ElseIf displayDimension.Type2 = swLinearDimension And _
                   Not firstLinearGeometryDumped Then

                targetName = "FIRST_LINEAR"
                firstLinearGeometryDumped = True
            End If

            toleranceSummary = _
                ReadDimensionTolerance(dimension, dimensionH7Proven)

            If targetName = "DIAMETER_47" And _
               dimensionH7Proven Then

                h7Proven = True
            End If

            ProbeLog "R23_SECTION_DIMENSION" & _
                "|index=" & CStr(dimensionCount) & _
                "|target=" & ProbeToken(targetName) & _
                "|name=" & ProbeToken( _
                    SafeDisplayDimensionName(displayDimension)) & _
                "|type2=" & CStr(displayDimension.Type2) & _
                "|nominalAvailable=" & CStr(nominalAvailable) & _
                "|nominalM=" & FormatProbeNumber(nominalM) & _
                "|dimensionH7Proven=" & CStr(dimensionH7Proven) & _
                "|" & toleranceSummary

            If targetName = "DIAMETER_47" Or _
               targetName = "DIAMETER_40" Or _
               targetName = "FIRST_LINEAR" Then

                DumpDisplayDimensionGeometry _
                    displayDimension, targetName
            End If
        End If
    Next i

    exactTargetCounts = _
        (diameter47Count = 1) And _
        (diameter40Count = 1) And _
        firstLinearGeometryDumped

    ProbeLog "R23_SECTION_DIMENSION_SUMMARY" & _
        "|count=" & CStr(dimensionCount) & _
        "|diameter47Found=" & CStr(diameter47Found) & _
        "|diameter47Count=" & CStr(diameter47Count) & _
        "|diameter40Found=" & CStr(diameter40Found) & _
        "|diameter40Count=" & CStr(diameter40Count) & _
        "|linearTargetSelected=" & CStr(firstLinearGeometryDumped) & _
        "|exactTargetCounts=" & CStr(exactTargetCounts) & _
        "|h7Proven=" & CStr(h7Proven) & _
        "|linearFound=" & CStr(linearFound) & _
        "|diametricLinearFound=" & CStr(diametricLinearFound)
    Exit Sub

Failed:
    ProbeLog "R23_SECTION_DIMENSION_ERROR" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

' Reads the original part-source dimensions behind the imported section
' diameters directly through IModelDoc2.Parameter (MCP contract: the fully
' qualified dimension name returns the IDimension).  Read-only access; the
' fixture is never modified.  Returns True when the Diameter-47 source
' dimension itself proves an H7 fit.
Private Function DumpPartSourceDimensionAuthority( _
    ByRef swPart As SldWorks.ModelDoc2) As Boolean

    On Error GoTo Failed

    Dim activeConfiguration As String
    activeConfiguration = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    Dim sourceH747 As Boolean
    Dim sourceH740 As Boolean

    sourceH747 = DumpOnePartSourceDimension( _
        swPart, "D1@Sketch4", "DIAMETER_47_SOURCE", _
        activeConfiguration)
    sourceH740 = DumpOnePartSourceDimension( _
        swPart, "D1@Sketch6", "DIAMETER_40_SOURCE", _
        activeConfiguration)

    ProbeLog "R23_H7_AUTHORITY" & _
        "|sourceDimension=D1@Sketch4" & _
        "|h7PresentInSource=" & CStr(sourceH747) & _
        "|configuration=" & ProbeToken(activeConfiguration) & _
        "|decisionRequired=" & IIf(sourceH747, "None", _
            "ModelAuthoritativeFailClosed_Or_ControlledTargetSpecAuthority") & _
        "|policy=NoInventedTolerance"

    DumpPartSourceDimensionAuthority = sourceH747
    Exit Function

Failed:
    ProbeLog "R23_SECTION_SOURCE_ERROR" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
    DumpPartSourceDimensionAuthority = False
End Function

Private Function DumpOnePartSourceDimension( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal dimensionName As String, _
    ByVal roleName As String, _
    ByVal activeConfiguration As String) As Boolean

    On Error GoTo Failed

    Dim dimension As SldWorks.Dimension
    Set dimension = swPart.Parameter(dimensionName)

    If dimension Is Nothing Then
        ProbeLog "R23_SECTION_SOURCE_TOLERANCE" & _
            "|role=" & ProbeToken(roleName) & _
            "|dimension=" & ProbeToken(dimensionName) & _
            "|status=ParameterReturnedNothing"
        Exit Function
    End If

    Dim nominalM As Double
    Dim nominalAvailable As Boolean
    nominalAvailable = TryReadNominal(dimension, nominalM)

    Dim sourceH7Proven As Boolean
    Dim toleranceSummary As String
    toleranceSummary = _
        ReadDimensionTolerance(dimension, sourceH7Proven)

    ProbeLog "R23_SECTION_SOURCE_TOLERANCE" & _
        "|role=" & ProbeToken(roleName) & _
        "|dimension=" & ProbeToken(dimensionName) & _
        "|fullName=" & ProbeToken(SafeDimensionFullName(dimension)) & _
        "|configuration=" & ProbeToken(activeConfiguration) & _
        "|nominalAvailable=" & CStr(nominalAvailable) & _
        "|nominalM=" & FormatProbeNumber(nominalM) & _
        "|sourceH7Proven=" & CStr(sourceH7Proven) & _
        "|" & toleranceSummary

    DumpOnePartSourceDimension = sourceH7Proven
    Exit Function

Failed:
    ProbeLog "R23_SECTION_SOURCE_TOLERANCE" & _
        "|role=" & ProbeToken(roleName) & _
        "|dimension=" & ProbeToken(dimensionName) & _
        "|status=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Function

Private Function SafeDimensionFullName( _
    ByRef dimension As SldWorks.Dimension) As String

    On Error GoTo Failed
    SafeDimensionFullName = dimension.FullName
    Exit Function

Failed:
    SafeDimensionFullName = "Unavailable:" & CStr(Err.Number)
End Function

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
        If VariantItemCount(values) < 1 Then Exit Function
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

Private Function ReadDimensionTolerance( _
    ByRef dimension As SldWorks.Dimension, _
    ByRef h7Proven As Boolean) As String

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

    Dim minimumM As Double
    Dim maximumM As Double
    Dim minimumStatus As Long
    Dim maximumStatus As Long

    minimumStatus = tolerance.GetMinValue2(minimumM)
    maximumStatus = tolerance.GetMaxValue2(maximumM)

    Dim holeFit As String
    Dim shaftFit As String
    Dim fitValues As String

    holeFit = tolerance.GetHoleFitValue
    shaftFit = tolerance.GetShaftFitValue

    On Error Resume Next
    fitValues = dimension.GetToleranceFitValues
    On Error GoTo Failed

    If InStr(1, holeFit, "H7", vbTextCompare) > 0 Or _
       InStr(1, fitValues, "H7", vbTextCompare) > 0 Then

        h7Proven = True
    End If

    ReadDimensionTolerance = _
        "toleranceType=" & CStr(tolerance.Type) & _
        "|fitType=" & CStr(tolerance.FitType) & _
        "|fitDisplayStyle=" & CStr(tolerance.FitDisplayStyle) & _
        "|holeFit=" & ProbeToken(holeFit) & _
        "|shaftFit=" & ProbeToken(shaftFit) & _
        "|fitValues=" & ProbeToken(fitValues) & _
        "|minimumStatus=" & CStr(minimumStatus) & _
        "|minimumM=" & FormatProbeNumber(minimumM) & _
        "|maximumStatus=" & CStr(maximumStatus) & _
        "|maximumM=" & FormatProbeNumber(maximumM)
    Exit Function

Failed:
    ReadDimensionTolerance = _
        "toleranceReadError=" & CStr(Err.Number) & _
        ":" & ProbeToken(Err.Description)
End Function

Private Sub DumpDisplayDimensionGeometry( _
    ByRef displayDimension As SldWorks.DisplayDimension, _
    ByVal targetName As String)

    On Error GoTo Failed

    Dim displayData As SldWorks.DisplayData
    Set displayData = displayDimension.GetDisplayData

    If displayData Is Nothing Then
        ProbeLog "R23_SECTION_DISPLAY" & _
            "|target=" & ProbeToken(targetName) & _
            "|status=Nothing"
        Exit Sub
    End If

    ProbeLog "R23_SECTION_DISPLAY" & _
        "|target=" & ProbeToken(targetName) & _
        "|lineCount=" & CStr(displayData.GetLineCount) & _
        "|arrowCount=" & CStr(displayData.GetArrowHeadCount) & _
        "|textCount=" & CStr(displayData.GetTextCount) & _
        "|arcCount=" & CStr(displayData.GetArcCount) & _
        "|ellipseCount=" & CStr(displayData.GetEllipseCount) & _
        "|polylineCount=" & CStr(displayData.GetPolyLineCount)

    Dim i As Long
    For i = 0 To displayData.GetLineCount - 1
        ProbeLog "R23_SECTION_DISPLAY_LINE" & _
            "|target=" & ProbeToken(targetName) & _
            "|index=" & CStr(i) & _
            "|data=" & ArrayToken( _
                displayData.GetLineAtIndex3(i), 16)
    Next i

    For i = 0 To displayData.GetArrowHeadCount - 1
        ProbeLog "R23_SECTION_DISPLAY_ARROW" & _
            "|target=" & ProbeToken(targetName) & _
            "|index=" & CStr(i) & _
            "|data=" & ArrayToken( _
                displayData.GetArrowHeadAtIndex2(i), 16)
    Next i

    For i = 0 To displayData.GetTextCount - 1
        ProbeLog "R23_SECTION_DISPLAY_TEXT" & _
            "|target=" & ProbeToken(targetName) & _
            "|index=" & CStr(i) & _
            "|text=" & ProbeToken( _
                displayData.GetTextAtIndex(i)) & _
            "|position=" & ArrayToken( _
                displayData.GetTextPositionAtIndex(i), 8) & _
            "|height=" & FormatProbeNumber( _
                displayData.GetTextHeightAtIndex(i)) & _
            "|angle=" & FormatProbeNumber( _
                displayData.GetTextAngleAtIndex(i))
    Next i
    Exit Sub

Failed:
    ProbeLog "R23_SECTION_DISPLAY_ERROR" & _
        "|target=" & ProbeToken(targetName) & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

Private Function DumpSectionLineGeometry( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef primaryView As SldWorks.View, _
    ByRef sectionView As SldWorks.View) As Boolean

    On Error GoTo Failed

    DumpSectionConstructionCapture

    Dim sectionData As SldWorks.DrSection
    Set sectionData = sectionView.GetSection

    If sectionData Is Nothing Then
        ProbeLog "R23_SECTION_DATA|status=Nothing"
        Exit Function
    End If

    ProbeLog "R23_SECTION_DATA" & _
        "|label=" & ProbeToken(sectionData.GetLabel) & _
        "|name=" & ProbeToken(sectionData.GetName) & _
        "|aligned=" & CStr(sectionData.IsAligned) & _
        "|sectionDepthM=" & FormatProbeNumber( _
            sectionData.SectionDepth) & _
        "|lineInfo=" & ArrayToken(sectionData.GetLineInfo, 80) & _
        "|arrowInfo=" & ArrayToken(sectionData.GetArrowInfo, 80) & _
        "|textInfo=" & ArrayToken(sectionData.GetTextInfo, 80)

    Dim reportedSize As Long
    Dim lineCount As Long
    lineCount = primaryView.GetSectionLineCount2(reportedSize)

    Dim sectionInfo As Variant
    sectionInfo = primaryView.GetSectionLineInfo2

    ProbeLog "R23_SECTION_LINE_READBACK" & _
        "|sourceView=" & ProbeToken(SafeViewName(primaryView)) & _
        "|lineCount=" & CStr(lineCount) & _
        "|reportedSize=" & CStr(reportedSize) & _
        "|arrayItems=" & CStr(VariantItemCount(sectionInfo)) & _
        "|raw=" & ArrayToken(sectionInfo, 200)

    If lineCount < 1 Or reportedSize < 1 Then Exit Function
    If Not IsArray(sectionInfo) Then Exit Function

    Set mParsedSegmentPoints = New Collection
    Set mParsedArrowPoints = New Collection
    Set mParsedLabelPoints = New Collection
    mParsedLabelTextHeight = 0#

    Dim parsedOk As Boolean
    parsedOk = ParseSectionLineInfo( _
        sectionInfo, SafeViewName(primaryView), lineCount)

    If parsedOk Then
        ProveSectionPayloadFrame
        TransformSegmentsAndEvaluateClearance swDraw, primaryView
    End If

    DumpSectionLineGeometry = parsedOk
    Exit Function

Failed:
    ProbeLog "R23_SECTION_LINE_ERROR" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Function

' Logs the construction-time capture recorded by the disposable
' Module2_DrawingPipeline overlay: the original page-frame path points and the
' converted view-sketch values actually passed to CreateLine.  The page values
' are labelled PageAtConstruction because later layout moves can relocate the
' view; current-page values are recomputed from the current transform.
Private Sub DumpSectionConstructionCapture()
    On Error GoTo Failed

    If Not Module2_DrawingPipeline.R23_SectionPlanCaptured Then
        ProbeLog "R23_JJ_CONSTRUCTION|status=NotCaptured"
        Exit Sub
    End If

    Dim pagePoints As Collection
    Dim sketchPoints As Collection
    Set pagePoints = Module2_DrawingPipeline.R23_SectionPagePoints
    Set sketchPoints = Module2_DrawingPipeline.R23_SectionSketchPoints

    ProbeLog "R23_JJ_CONSTRUCTION" & _
        "|sourceView=" & ProbeToken( _
            Module2_DrawingPipeline.R23_SectionPlanSourceView) & _
        "|pointCount=" & CStr(pagePoints.Count) & _
        "|status=Captured"

    Dim i As Long
    For i = 1 To pagePoints.Count
        Dim pagePoint As Variant
        Dim sketchPoint As Variant
        pagePoint = pagePoints(i)
        sketchPoint = sketchPoints(i)

        ProbeLog "R23_JJ_CONSTRUCTION_POINT" & _
            "|index=" & CStr(i) & _
            "|frame=PageAtConstruction" & _
            "|pageX=" & FormatProbeNumber(CDbl(pagePoint(0))) & _
            "|pageY=" & FormatProbeNumber(CDbl(pagePoint(1))) & _
            "|frame2=ViewSketch" & _
            "|sketchX=" & FormatProbeNumber(CDbl(sketchPoint(0))) & _
            "|sketchY=" & FormatProbeNumber(CDbl(sketchPoint(1)))
    Next i
    Exit Sub

Failed:
    ProbeLog "R23_JJ_CONSTRUCTION|status=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

' Proves which frame the returned segment payload uses by comparing it with
' the captured view-sketch CreateLine inputs.  An exact match proves the
' payload segments are view-sketch coordinates, so exactly one sketch-to-page
' conversion is required before any clearance comparison.
Private Sub ProveSectionPayloadFrame()
    On Error GoTo Failed

    If Not Module2_DrawingPipeline.R23_SectionPlanCaptured Then
        ProbeLog "R23_JJ_FRAME_PROOF|status=NoConstructionCapture"
        Exit Sub
    End If

    Dim sketchPoints As Collection
    Set sketchPoints = Module2_DrawingPipeline.R23_SectionSketchPoints

    If sketchPoints.Count <> mParsedSegmentPoints.Count Then
        ProbeLog "R23_JJ_FRAME_PROOF" & _
            "|status=PointCountMismatch" & _
            "|constructionPoints=" & CStr(sketchPoints.Count) & _
            "|payloadPoints=" & CStr(mParsedSegmentPoints.Count)
        Exit Sub
    End If

    Dim allMatch As Boolean
    allMatch = True

    Dim i As Long
    For i = 1 To sketchPoints.Count
        Dim sketchPoint As Variant
        Dim payloadPoint As Variant
        sketchPoint = sketchPoints(i)
        payloadPoint = mParsedSegmentPoints(i)

        Dim deltaM As Double
        deltaM = Sqr( _
            (CDbl(sketchPoint(0)) - CDbl(payloadPoint(0))) ^ 2 + _
            (CDbl(sketchPoint(1)) - CDbl(payloadPoint(1))) ^ 2)

        If deltaM > PAGE_MATCH_TOLERANCE_M Then allMatch = False

        ProbeLog "R23_JJ_FRAME_PROOF_POINT" & _
            "|index=" & CStr(i) & _
            "|sketchX=" & FormatProbeNumber(CDbl(sketchPoint(0))) & _
            "|sketchY=" & FormatProbeNumber(CDbl(sketchPoint(1))) & _
            "|payloadX=" & FormatProbeNumber(CDbl(payloadPoint(0))) & _
            "|payloadY=" & FormatProbeNumber(CDbl(payloadPoint(1))) & _
            "|deltaM=" & FormatProbeNumber(deltaM)
    Next i

    ProbeLog "R23_JJ_FRAME_PROOF" & _
        "|payloadSegmentFrame=" & _
            IIf(allMatch, "ViewSketchProved", "Unproved") & _
        "|matchToleranceM=" & FormatProbeNumber(PAGE_MATCH_TOLERANCE_M) & _
        "|status=" & IIf(allMatch, "SUCCESS", "MISMATCH")
    Exit Sub

Failed:
    ProbeLog "R23_JJ_FRAME_PROOF|status=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

' Converts every parsed payload segment endpoint from the view-sketch frame to
' the current page frame exactly once, using the exact inverse of the
' page-to-sketch conversion applied at construction (IView.GetXform origin and
' scale plus IView.Angle rotation), then evaluates page-frame clearance.
Private Sub TransformSegmentsAndEvaluateClearance( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef primaryView As SldWorks.View)

    On Error GoTo Failed

    Dim xform As Variant
    xform = primaryView.GetXform

    If IsEmpty(xform) Or Not IsArray(xform) Then
        ProbeLog "R23_JJ_PAGE_TRANSFORM|status=XformUnavailable"
        Exit Sub
    End If

    If UBound(xform) - LBound(xform) + 1 < 3 Then
        ProbeLog "R23_JJ_PAGE_TRANSFORM|status=XformTooShort"
        Exit Sub
    End If

    Dim originX As Double
    Dim originY As Double
    Dim viewScale As Double
    Dim viewAngle As Double

    originX = CDbl(xform(LBound(xform)))
    originY = CDbl(xform(LBound(xform) + 1))
    viewScale = CDbl(xform(LBound(xform) + 2))
    viewAngle = primaryView.Angle

    ProbeLog "R23_JJ_PAGE_TRANSFORM" & _
        "|sourceView=" & ProbeToken(SafeViewName(primaryView)) & _
        "|xformSheetX=" & FormatProbeNumber(originX) & _
        "|xformSheetY=" & FormatProbeNumber(originY) & _
        "|scale=" & FormatProbeNumber(viewScale) & _
        "|angle=" & FormatProbeNumber(viewAngle) & _
        "|direction=ViewSketchToPage|applications=ExactlyOnce"

    If viewScale <= 0# Then
        ProbeLog "R23_JJ_PAGE_TRANSFORM|status=InvalidScale"
        Exit Sub
    End If

    Dim pageSegments As New Collection

    Dim i As Long
    For i = 1 To mParsedSegmentPoints.Count
        Dim payloadPoint As Variant
        payloadPoint = mParsedSegmentPoints(i)

        Dim pageX As Double
        Dim pageY As Double
        SketchPointToPage _
            CDbl(payloadPoint(0)), CDbl(payloadPoint(1)), _
            originX, originY, viewScale, viewAngle, pageX, pageY

        Dim pagePoint(0 To 1) As Double
        pagePoint(0) = pageX
        pagePoint(1) = pageY
        pageSegments.Add pagePoint

        ProbeLog "R23_JJ_SEGMENT_POINT" & _
            "|index=" & CStr(i) & _
            "|frame=ViewSketch" & _
            "|sketchX=" & FormatProbeNumber(CDbl(payloadPoint(0))) & _
            "|sketchY=" & FormatProbeNumber(CDbl(payloadPoint(1))) & _
            "|frame2=Page" & _
            "|pageX=" & FormatProbeNumber(pageX) & _
            "|pageY=" & FormatProbeNumber(pageY)
    Next i

    DumpPageFrameRegions swDraw
    EvaluateJJClearance pageSegments
    Exit Sub

Failed:
    ProbeLog "R23_JJ_PAGE_TRANSFORM|status=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

' Exact inverse of Module2_DrawingPipeline.SheetToViewSketchCoordinates.
Private Sub SketchPointToPage( _
    ByVal sketchX As Double, _
    ByVal sketchY As Double, _
    ByVal originX As Double, _
    ByVal originY As Double, _
    ByVal viewScale As Double, _
    ByVal viewAngle As Double, _
    ByRef pageX As Double, _
    ByRef pageY As Double)

    Dim deltaX As Double
    Dim deltaY As Double
    deltaX = sketchX * Cos(viewAngle) - sketchY * Sin(viewAngle)
    deltaY = sketchX * Sin(viewAngle) + sketchY * Cos(viewAngle)

    pageX = originX + deltaX * viewScale
    pageY = originY + deltaY * viewScale
End Sub

Private Function ParseSectionLineInfo( _
    ByVal sectionInfo As Variant, _
    ByVal viewName As String, _
    ByVal expectedLineCount As Long) As Boolean

    On Error GoTo Failed

    Dim cursor As Long
    Dim upperIndex As Long
    cursor = LBound(sectionInfo)
    upperIndex = UBound(sectionInfo)

    If cursor > upperIndex Then
        RecordSectionParseFailure viewName, "EmptyArray", cursor
        Exit Function
    End If

    Dim lineCount As Long
    lineCount = CLng(sectionInfo(cursor))
    cursor = cursor + 1

    If lineCount < 1 Then
        RecordSectionParseFailure _
            viewName, "InvalidLineCount:" & CStr(lineCount), cursor
        Exit Function
    End If

    If lineCount <> expectedLineCount Then
        RecordSectionParseFailure _
            viewName, _
            "LineCountMismatch:array=" & CStr(lineCount) & _
                ":reported=" & CStr(expectedLineCount), _
            cursor
        Exit Function
    End If

    Dim lineIndex As Long
    For lineIndex = 1 To lineCount
        If cursor + 1 > upperIndex Then
            RecordSectionParseFailure _
                viewName, "TruncatedLineHeader", cursor
            Exit Function
        End If

        Dim layerValue As Double
        Dim segmentCount As Long
        layerValue = CDbl(sectionInfo(cursor))
        cursor = cursor + 1
        segmentCount = CLng(sectionInfo(cursor))
        cursor = cursor + 1

        If segmentCount < 1 Then
            RecordSectionParseFailure _
                viewName, "InvalidSegmentCount", cursor
            Exit Function
        End If

        Dim segmentIndex As Long
        For segmentIndex = 1 To segmentCount
            If cursor + 6 > upperIndex Then
                RecordSectionParseFailure _
                    viewName, "TruncatedSegment", cursor
                Exit Function
            End If

            ProbeLog "R23_SECTION_SEGMENT" & _
                "|view=" & ProbeToken(viewName) & _
                "|line=" & CStr(lineIndex) & _
                "|segment=" & CStr(segmentIndex) & _
                "|frame=RawPayload" & _
                "|startX=" & FormatProbeNumber( _
                    CDbl(sectionInfo(cursor + 1))) & _
                "|startY=" & FormatProbeNumber( _
                    CDbl(sectionInfo(cursor + 2))) & _
                "|endX=" & FormatProbeNumber( _
                    CDbl(sectionInfo(cursor + 4))) & _
                "|endY=" & FormatProbeNumber( _
                    CDbl(sectionInfo(cursor + 5)))

            CaptureParsedPoint mParsedSegmentPoints, _
                CDbl(sectionInfo(cursor + 1)), _
                CDbl(sectionInfo(cursor + 2))
            CaptureParsedPoint mParsedSegmentPoints, _
                CDbl(sectionInfo(cursor + 4)), _
                CDbl(sectionInfo(cursor + 5))

            cursor = cursor + 7
        Next segmentIndex

        If cursor + 24 > upperIndex Then
            RecordSectionParseFailure _
                viewName, "TruncatedArrowOrLabelBlock", cursor
            Exit Function
        End If

        ProbeLog "R23_SECTION_ARROW" & _
            "|view=" & ProbeToken(viewName) & _
            "|line=" & CStr(lineIndex) & _
            "|index=1" & _
            "|frame=Page" & _
            "|startX=" & FormatProbeNumber(CDbl(sectionInfo(cursor))) & _
            "|startY=" & FormatProbeNumber(CDbl(sectionInfo(cursor + 1))) & _
            "|endX=" & FormatProbeNumber(CDbl(sectionInfo(cursor + 3))) & _
            "|endY=" & FormatProbeNumber(CDbl(sectionInfo(cursor + 4)))
        CaptureParsedPoint mParsedArrowPoints, _
            CDbl(sectionInfo(cursor)), CDbl(sectionInfo(cursor + 1))
        CaptureParsedPoint mParsedArrowPoints, _
            CDbl(sectionInfo(cursor + 3)), CDbl(sectionInfo(cursor + 4))
        cursor = cursor + 9

        ProbeLog "R23_SECTION_ARROW" & _
            "|view=" & ProbeToken(viewName) & _
            "|line=" & CStr(lineIndex) & _
            "|index=2" & _
            "|frame=Page" & _
            "|startX=" & FormatProbeNumber(CDbl(sectionInfo(cursor))) & _
            "|startY=" & FormatProbeNumber(CDbl(sectionInfo(cursor + 1))) & _
            "|endX=" & FormatProbeNumber(CDbl(sectionInfo(cursor + 3))) & _
            "|endY=" & FormatProbeNumber(CDbl(sectionInfo(cursor + 4)))
        CaptureParsedPoint mParsedArrowPoints, _
            CDbl(sectionInfo(cursor)), CDbl(sectionInfo(cursor + 1))
        CaptureParsedPoint mParsedArrowPoints, _
            CDbl(sectionInfo(cursor + 3)), CDbl(sectionInfo(cursor + 4))
        cursor = cursor + 9

        ProbeLog "R23_SECTION_LABEL" & _
            "|view=" & ProbeToken(viewName) & _
            "|line=" & CStr(lineIndex) & _
            "|frame=Page" & _
            "|text1X=" & FormatProbeNumber(CDbl(sectionInfo(cursor))) & _
            "|text1Y=" & FormatProbeNumber(CDbl(sectionInfo(cursor + 1))) & _
            "|text2X=" & FormatProbeNumber(CDbl(sectionInfo(cursor + 3))) & _
            "|text2Y=" & FormatProbeNumber(CDbl(sectionInfo(cursor + 4))) & _
            "|textHeight=" & FormatProbeNumber( _
                Abs(CDbl(sectionInfo(cursor + 6)))) & _
            "|layer=" & FormatProbeNumber(layerValue) & _
            "|segments=" & CStr(segmentCount)
        CaptureParsedPoint mParsedLabelPoints, _
            CDbl(sectionInfo(cursor)), CDbl(sectionInfo(cursor + 1))
        CaptureParsedPoint mParsedLabelPoints, _
            CDbl(sectionInfo(cursor + 3)), CDbl(sectionInfo(cursor + 4))
        mParsedLabelTextHeight = Abs(CDbl(sectionInfo(cursor + 6)))
        cursor = cursor + 7
    Next lineIndex

    ProbeLog "R23_SECTION_PARSE" & _
        "|view=" & ProbeToken(viewName) & _
        "|lineCount=" & CStr(lineCount) & _
        "|finalCursor=" & CStr(cursor) & _
        "|upperIndex=" & CStr(upperIndex) & _
        "|status=SUCCESS"

    ParseSectionLineInfo = True
    Exit Function

Failed:
    RecordSectionParseFailure _
        viewName, _
        "APIError:" & CStr(Err.Number) & ":" & Err.Description, _
        cursor
End Function

Private Sub RecordSectionParseFailure( _
    ByVal viewName As String, _
    ByVal reason As String, _
    ByVal cursor As Long)

    ProbeLog "R23_SECTION_PARSE" & _
        "|view=" & ProbeToken(viewName) & _
        "|reason=" & ProbeToken(reason) & _
        "|cursor=" & CStr(cursor) & _
        "|status=FAILED"
End Sub

Private Sub CaptureParsedPoint( _
    ByRef target As Collection, _
    ByVal x As Double, _
    ByVal y As Double)

    On Error Resume Next
    If target Is Nothing Then Exit Sub

    Dim capturedPoint(0 To 1) As Double
    capturedPoint(0) = x
    capturedPoint(1) = y
    target.Add capturedPoint
    On Error GoTo 0
End Sub

' Logs every page-frame reserved region used by the clearance comparison,
' with its provenance.  All values are in sheet/page coordinates.
Private Sub DumpPageFrameRegions(ByRef swDraw As SldWorks.DrawingDoc)
    On Error GoTo Failed

    mContentRegionValid = False
    mTitleRegionValid = False
    mPartIdRegionValid = False

    Dim evidence As CRunEvidence
    Set evidence = Module1_Main.GlobalEvidence

    If evidence Is Nothing Then
        ProbeLog "R23_JJ_REGION|status=EvidenceUnavailable"
    ElseIf Not evidence.LayoutBoundariesProven Then
        ProbeLog "R23_JJ_REGION|status=LayoutBoundariesNotProven"
    Else
        mContentRegionValid = True
        mContentLeft = evidence.ContentBorderLeft
        mContentBottom = evidence.ContentBorderBottom
        mContentRight = evidence.ContentBorderRight
        mContentTop = evidence.ContentBorderTop

        mTitleRegionValid = True
        mTitleLeft = evidence.TitleBlockLeft
        mTitleBottom = evidence.TitleBlockBottom
        mTitleRight = evidence.TitleBlockRight
        mTitleTop = evidence.TitleBlockTop

        ProbeLog "R23_JJ_REGION|region=Sheet|frame=Page" & _
            "|widthM=" & FormatProbeNumber(evidence.SheetWidth) & _
            "|heightM=" & FormatProbeNumber(evidence.SheetHeight) & _
            "|source=ISheet.GetSize"

        ProbeLog "R23_JJ_REGION|region=ContentBorder|frame=Page" & _
            "|left=" & FormatProbeNumber(mContentLeft) & _
            "|bottom=" & FormatProbeNumber(mContentBottom) & _
            "|right=" & FormatProbeNumber(mContentRight) & _
            "|top=" & FormatProbeNumber(mContentTop) & _
            "|source=ISheet.GetZoneMargin" & _
            "|note=OutsideIsZoneNumberBand"

        ProbeLog "R23_JJ_REGION|region=TitleBlock|frame=Page" & _
            "|left=" & FormatProbeNumber(mTitleLeft) & _
            "|bottom=" & FormatProbeNumber(mTitleBottom) & _
            "|right=" & FormatProbeNumber(mTitleRight) & _
            "|top=" & FormatProbeNumber(mTitleTop) & _
            "|source=MeasureControlledSheetRegions"
    End If

    FindPartIdentificationNoteExtent swDraw

    DumpAllViewOutlines swDraw
    Exit Sub

Failed:
    ProbeLog "R23_JJ_REGION|status=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

' The probe drawing's part-identification band is the template sheet-format
' linked note that renders *P-0251-14A-001*.  Its extent is measured with
' INote.GetExtent, which returns sheet-space box extents.
Private Sub FindPartIdentificationNoteExtent( _
    ByRef swDraw As SldWorks.DrawingDoc)

    On Error GoTo Failed

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    Do While Not swView Is Nothing
        Dim note As SldWorks.Note
        Set note = swView.GetFirstNote

        Do While Not note Is Nothing
            Dim renderedText As String
            renderedText = vbNullString
            On Error Resume Next
            renderedText = note.GetText
            On Error GoTo Failed

            If InStr(1, renderedText, "*P-0251-14A-001*", _
                     vbTextCompare) > 0 Then

                Dim extent As Variant
                extent = note.GetExtent

                If IsArray(extent) Then
                    If VariantItemCount(extent) >= 6 Then
                        Dim baseIndex As Long
                        baseIndex = LBound(extent)

                        mPartIdLeft = CDbl(extent(baseIndex))
                        mPartIdBottom = CDbl(extent(baseIndex + 1))
                        mPartIdRight = CDbl(extent(baseIndex + 3))
                        mPartIdTop = CDbl(extent(baseIndex + 4))

                        If mPartIdRight < mPartIdLeft Then
                            SwapDoubles mPartIdLeft, mPartIdRight
                        End If
                        If mPartIdTop < mPartIdBottom Then
                            SwapDoubles mPartIdBottom, mPartIdTop
                        End If

                        mPartIdRegionValid = _
                            (mPartIdRight > mPartIdLeft) And _
                            (mPartIdTop > mPartIdBottom)

                        ProbeLog "R23_JJ_REGION" & _
                            "|region=PartIdentification|frame=Page" & _
                            "|left=" & FormatProbeNumber(mPartIdLeft) & _
                            "|bottom=" & FormatProbeNumber(mPartIdBottom) & _
                            "|right=" & FormatProbeNumber(mPartIdRight) & _
                            "|top=" & FormatProbeNumber(mPartIdTop) & _
                            "|owner=" & ProbeToken(SafeViewName(swView)) & _
                            "|source=INote.GetExtent" & _
                            "|noteText=" & ProbeToken(renderedText)
                        Exit Sub
                    End If
                End If
            End If

            Set note = note.GetNext
        Loop

        Set swView = swView.GetNextView
    Loop

    ProbeLog "R23_JJ_REGION|region=PartIdentification" & _
        "|status=NoteNotFound" & _
        "|note=ClearanceAgainstPartIdUnavailable"
    Exit Sub

Failed:
    ProbeLog "R23_JJ_REGION|region=PartIdentification" & _
        "|status=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

Private Sub SwapDoubles(ByRef firstValue As Double, _
    ByRef secondValue As Double)

    Dim holdValue As Double
    holdValue = firstValue
    firstValue = secondValue
    secondValue = holdValue
End Sub

Private Sub DumpAllViewOutlines(ByRef swDraw As SldWorks.DrawingDoc)
    On Error GoTo Failed

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        Dim outline As Variant
        outline = swView.GetOutline

        If IsArray(outline) Then
            If VariantItemCount(outline) = 4 Then
                Dim baseIndex As Long
                baseIndex = LBound(outline)

                ProbeLog "R23_JJ_REGION|region=ViewOutline|frame=Page" & _
                    "|view=" & ProbeToken(SafeViewName(swView)) & _
                    "|left=" & FormatProbeNumber( _
                        CDbl(outline(baseIndex))) & _
                    "|bottom=" & FormatProbeNumber( _
                        CDbl(outline(baseIndex + 1))) & _
                    "|right=" & FormatProbeNumber( _
                        CDbl(outline(baseIndex + 2))) & _
                    "|top=" & FormatProbeNumber( _
                        CDbl(outline(baseIndex + 3))) & _
                    "|source=IView.GetOutline" & _
                    "|note=OutlineIncludesAnnotations"
            End If
        End If

        Set swView = swView.GetNextView
    Loop
    Exit Sub

Failed:
    ProbeLog "R23_JJ_REGION|region=ViewOutline|status=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

' Compares only page-frame J-J geometry against the page-frame reserved
' regions.  Segments arrive already converted (exactly once); arrow and label
' coordinates are page-frame directly from the payload.
Private Sub EvaluateJJClearance(ByRef pageSegments As Collection)
    On Error GoTo Failed

    Dim violationCount As Long
    Dim itemCount As Long
    Dim unavailableChecks As Long

    Dim i As Long
    For i = 1 To pageSegments.Count - 1 Step 2
        itemCount = itemCount + 1
        EvaluateSegmentClearance _
            "Segment" & CStr((i + 1) \ 2), _
            pageSegments(i), pageSegments(i + 1), _
            violationCount, unavailableChecks
    Next i

    For i = 1 To mParsedArrowPoints.Count - 1 Step 2
        itemCount = itemCount + 1
        EvaluateSegmentClearance _
            "Arrow" & CStr((i + 1) \ 2), _
            mParsedArrowPoints(i), mParsedArrowPoints(i + 1), _
            violationCount, unavailableChecks
    Next i

    For i = 1 To mParsedLabelPoints.Count
        itemCount = itemCount + 1
        EvaluateLabelClearance _
            "Label" & CStr(i), mParsedLabelPoints(i), _
            violationCount, unavailableChecks
    Next i

    ProbeLog "R23_JJ_CLEARANCE_SUMMARY" & _
        "|frame=Page" & _
        "|items=" & CStr(itemCount) & _
        "|violations=" & CStr(violationCount) & _
        "|unavailableChecks=" & CStr(unavailableChecks) & _
        "|allClear=" & CStr( _
            violationCount = 0 And unavailableChecks = 0) & _
        "|note=TruthfulMeasurementNotAcceptance"
    Exit Sub

Failed:
    ProbeLog "R23_JJ_CLEARANCE_SUMMARY|status=ReadError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & ProbeToken(Err.Description)
End Sub

Private Sub EvaluateSegmentClearance( _
    ByVal itemName As String, _
    ByVal startPoint As Variant, _
    ByVal endPoint As Variant, _
    ByRef violationCount As Long, _
    ByRef unavailableChecks As Long)

    Dim contentResult As String
    Dim titleResult As String
    Dim partIdResult As String

    If mContentRegionValid Then
        If PointInRect(CDbl(startPoint(0)), CDbl(startPoint(1)), _
               mContentLeft, mContentBottom, _
               mContentRight, mContentTop) And _
           PointInRect(CDbl(endPoint(0)), CDbl(endPoint(1)), _
               mContentLeft, mContentBottom, _
               mContentRight, mContentTop) Then
            contentResult = "Inside"
        Else
            contentResult = "VIOLATION_OutsideContentBorder"
            violationCount = violationCount + 1
        End If
    Else
        contentResult = "Unavailable"
        unavailableChecks = unavailableChecks + 1
    End If

    If mTitleRegionValid Then
        If SegmentTouchesRect(startPoint, endPoint, _
               mTitleLeft, mTitleBottom, mTitleRight, mTitleTop) Then
            titleResult = "VIOLATION_TitleBlock"
            violationCount = violationCount + 1
        Else
            titleResult = "Clear"
        End If
    Else
        titleResult = "Unavailable"
        unavailableChecks = unavailableChecks + 1
    End If

    If mPartIdRegionValid Then
        If SegmentTouchesRect(startPoint, endPoint, _
               mPartIdLeft, mPartIdBottom, mPartIdRight, mPartIdTop) Then
            partIdResult = "VIOLATION_PartIdentification"
            violationCount = violationCount + 1
        Else
            partIdResult = "Clear"
        End If
    Else
        partIdResult = "Unavailable"
        unavailableChecks = unavailableChecks + 1
    End If

    ProbeLog "R23_JJ_CLEARANCE" & _
        "|item=" & ProbeToken(itemName) & _
        "|frame=Page" & _
        "|startX=" & FormatProbeNumber(CDbl(startPoint(0))) & _
        "|startY=" & FormatProbeNumber(CDbl(startPoint(1))) & _
        "|endX=" & FormatProbeNumber(CDbl(endPoint(0))) & _
        "|endY=" & FormatProbeNumber(CDbl(endPoint(1))) & _
        "|contentBorder=" & ProbeToken(contentResult) & _
        "|titleBlock=" & ProbeToken(titleResult) & _
        "|partIdentification=" & ProbeToken(partIdResult)
End Sub

' The label box approximates one J character anchored at the reported text
' position: lower-left anchor, square of side textHeight.  The approximation
' is logged so the screenshot review can weigh it.
Private Sub EvaluateLabelClearance( _
    ByVal itemName As String, _
    ByVal anchorPoint As Variant, _
    ByRef violationCount As Long, _
    ByRef unavailableChecks As Long)

    Dim boxLeft As Double
    Dim boxBottom As Double
    Dim boxRight As Double
    Dim boxTop As Double

    boxLeft = CDbl(anchorPoint(0))
    boxBottom = CDbl(anchorPoint(1))
    boxRight = boxLeft + mParsedLabelTextHeight
    boxTop = boxBottom + mParsedLabelTextHeight

    Dim contentResult As String
    Dim titleResult As String
    Dim partIdResult As String

    If mContentRegionValid Then
        If boxLeft >= mContentLeft And boxBottom >= mContentBottom And _
           boxRight <= mContentRight And boxTop <= mContentTop Then
            contentResult = "Inside"
        Else
            contentResult = "VIOLATION_OutsideContentBorder"
            violationCount = violationCount + 1
        End If
    Else
        contentResult = "Unavailable"
        unavailableChecks = unavailableChecks + 1
    End If

    If mTitleRegionValid Then
        If RectsOverlap(boxLeft, boxBottom, boxRight, boxTop, _
               mTitleLeft, mTitleBottom, mTitleRight, mTitleTop) Then
            titleResult = "VIOLATION_TitleBlock"
            violationCount = violationCount + 1
        Else
            titleResult = "Clear"
        End If
    Else
        titleResult = "Unavailable"
        unavailableChecks = unavailableChecks + 1
    End If

    If mPartIdRegionValid Then
        If RectsOverlap(boxLeft, boxBottom, boxRight, boxTop, _
               mPartIdLeft, mPartIdBottom, mPartIdRight, mPartIdTop) Then
            partIdResult = "VIOLATION_PartIdentification"
            violationCount = violationCount + 1
        Else
            partIdResult = "Clear"
        End If
    Else
        partIdResult = "Unavailable"
        unavailableChecks = unavailableChecks + 1
    End If

    ProbeLog "R23_JJ_CLEARANCE" & _
        "|item=" & ProbeToken(itemName) & _
        "|frame=Page" & _
        "|boxModel=AnchorLowerLeftSquareSideTextHeight" & _
        "|left=" & FormatProbeNumber(boxLeft) & _
        "|bottom=" & FormatProbeNumber(boxBottom) & _
        "|right=" & FormatProbeNumber(boxRight) & _
        "|top=" & FormatProbeNumber(boxTop) & _
        "|contentBorder=" & ProbeToken(contentResult) & _
        "|titleBlock=" & ProbeToken(titleResult) & _
        "|partIdentification=" & ProbeToken(partIdResult)
End Sub

Private Function PointInRect( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal rectLeft As Double, _
    ByVal rectBottom As Double, _
    ByVal rectRight As Double, _
    ByVal rectTop As Double) As Boolean

    PointInRect = _
        (x >= rectLeft) And (x <= rectRight) And _
        (y >= rectBottom) And (y <= rectTop)
End Function

Private Function RectsOverlap( _
    ByVal aLeft As Double, ByVal aBottom As Double, _
    ByVal aRight As Double, ByVal aTop As Double, _
    ByVal bLeft As Double, ByVal bBottom As Double, _
    ByVal bRight As Double, ByVal bTop As Double) As Boolean

    RectsOverlap = Not ( _
        aRight < bLeft Or bRight < aLeft Or _
        aTop < bBottom Or bTop < aBottom)
End Function

Private Function SegmentTouchesRect( _
    ByVal startPoint As Variant, _
    ByVal endPoint As Variant, _
    ByVal rectLeft As Double, _
    ByVal rectBottom As Double, _
    ByVal rectRight As Double, _
    ByVal rectTop As Double) As Boolean

    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    x1 = CDbl(startPoint(0))
    y1 = CDbl(startPoint(1))
    x2 = CDbl(endPoint(0))
    y2 = CDbl(endPoint(1))

    If PointInRect(x1, y1, rectLeft, rectBottom, rectRight, rectTop) Then
        SegmentTouchesRect = True
        Exit Function
    End If

    If PointInRect(x2, y2, rectLeft, rectBottom, rectRight, rectTop) Then
        SegmentTouchesRect = True
        Exit Function
    End If

    SegmentTouchesRect = _
        SegmentsIntersect(x1, y1, x2, y2, _
            rectLeft, rectBottom, rectRight, rectBottom) Or _
        SegmentsIntersect(x1, y1, x2, y2, _
            rectRight, rectBottom, rectRight, rectTop) Or _
        SegmentsIntersect(x1, y1, x2, y2, _
            rectRight, rectTop, rectLeft, rectTop) Or _
        SegmentsIntersect(x1, y1, x2, y2, _
            rectLeft, rectTop, rectLeft, rectBottom)
End Function

Private Function SegmentsIntersect( _
    ByVal ax1 As Double, ByVal ay1 As Double, _
    ByVal ax2 As Double, ByVal ay2 As Double, _
    ByVal bx1 As Double, ByVal by1 As Double, _
    ByVal bx2 As Double, ByVal by2 As Double) As Boolean

    Dim d1 As Double
    Dim d2 As Double
    Dim d3 As Double
    Dim d4 As Double

    d1 = CrossDirection(bx1, by1, bx2, by2, ax1, ay1)
    d2 = CrossDirection(bx1, by1, bx2, by2, ax2, ay2)
    d3 = CrossDirection(ax1, ay1, ax2, ay2, bx1, by1)
    d4 = CrossDirection(ax1, ay1, ax2, ay2, bx2, by2)

    SegmentsIntersect = _
        (((d1 > 0# And d2 < 0#) Or (d1 < 0# And d2 > 0#)) And _
         ((d3 > 0# And d4 < 0#) Or (d3 < 0# And d4 > 0#)))
End Function

Private Function CrossDirection( _
    ByVal originX As Double, ByVal originY As Double, _
    ByVal throughX As Double, ByVal throughY As Double, _
    ByVal pointX As Double, ByVal pointY As Double) As Double

    CrossDirection = _
        (throughX - originX) * (pointY - originY) - _
        (throughY - originY) * (pointX - originX)
End Function

Private Function CountDisplayDimensions( _
    ByRef swView As SldWorks.View) As Long

    On Error GoTo Failed
    CountDisplayDimensions = VariantItemCount(swView.GetDisplayDimensions)
    Exit Function

Failed:
    CountDisplayDimensions = 0
End Function

Private Function SafeDisplayDimensionName( _
    ByRef displayDimension As SldWorks.DisplayDimension) As String

    On Error GoTo Failed
    SafeDisplayDimensionName = _
        displayDimension.GetNameForSelection
    Exit Function

Failed:
    SafeDisplayDimensionName = _
        "Unavailable:" & CStr(Err.Number)
End Function

Private Function SafeAnnotationPosition( _
    ByRef annotation As SldWorks.Annotation) As Variant

    On Error GoTo Failed
    If annotation Is Nothing Then Exit Function
    SafeAnnotationPosition = annotation.GetPosition
    Exit Function

Failed:
    SafeAnnotationPosition = Empty
End Function

Private Function SafeViewName( _
    ByRef swView As SldWorks.View) As String

    If swView Is Nothing Then
        SafeViewName = "Nothing"
        Exit Function
    End If

    On Error Resume Next
    SafeViewName = swView.GetName2
    If Len(SafeViewName) = 0 Then SafeViewName = swView.Name
    On Error GoTo 0
End Function

Private Function DecodeOrdinateResult( _
    ByVal resultCode As Long) As String

    Select Case resultCode
        Case 0: DecodeOrdinateResult = "0 Success"
        Case 1: DecodeOrdinateResult = "1 GeneralFailure"
        Case 2: DecodeOrdinateResult = "2 NoInternalDimensions"
        Case 3: DecodeOrdinateResult = "3 BadSelection"
        Case 4: DecodeOrdinateResult = "4 ModelNotLoaded"
        Case 5: DecodeOrdinateResult = "5 SamePartOnly"
        Case 6: DecodeOrdinateResult = "6 ExtraSelection"
        Case 7: DecodeOrdinateResult = "7 OrdinateFailure"
        Case 8: DecodeOrdinateResult = "8 Duplicate"
        Case 9: DecodeOrdinateResult = "9 BadDirection"
        Case -1: DecodeOrdinateResult = "-1 Undefined"
        Case -9999: DecodeOrdinateResult = "NotCalled"
        Case Else
            DecodeOrdinateResult = _
                CStr(resultCode) & " Unknown"
    End Select
End Function

Private Sub FinishDrawingProbe( _
    ByVal probeName As String, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal modelSaveFlagBefore As Boolean)

    On Error Resume Next
    Module4_ModelItemImporter.R23_ClearImportProbe

    Dim cleanupSelectionCount As Long

    If Not swDrawModel Is Nothing Then
        swDrawModel.SetPickMode
        swDrawModel.ClearSelection2 True
        cleanupSelectionCount = _
            swDrawModel.SelectionManager.GetSelectedObjectCount2(-1)

        If Not swDraw Is Nothing Then
            Module8_RuntimeSupport.RestoreSheetContext _
                swDrawModel, swDraw
        End If
    Else
        cleanupSelectionCount = -1
    End If

    Dim modelSaveFlagAfter As Boolean
    If Not swPart Is Nothing Then
        modelSaveFlagAfter = NormalizeSwBoolean(swPart.GetSaveFlag)
    End If

    ProbeLog "R23_DRAWING_PROBE_END" & _
        "|name=" & ProbeToken(probeName) & _
        "|cleanupSelection=" & CStr(cleanupSelectionCount) & _
        "|setPickModeCalled=True" & _
        "|modelSaveFlagBefore=" & CStr(modelSaveFlagBefore) & _
        "|modelSaveFlagAfter=" & CStr(modelSaveFlagAfter) & _
        "|modelUnchanged=" & CStr( _
            modelSaveFlagBefore = modelSaveFlagAfter) & _
        "|drawingSaved=False" & _
        "|status=COMPLETE"

    CloseProbeLog
    On Error GoTo 0
End Sub

Private Function VariantItemCount(ByVal value As Variant) As Long
    On Error GoTo Failed

    If IsArray(value) Then
        VariantItemCount = UBound(value) - LBound(value) + 1
    ElseIf IsEmpty(value) Or IsNull(value) Then
        VariantItemCount = 0
    Else
        VariantItemCount = 1
    End If
    Exit Function

Failed:
    VariantItemCount = 0
End Function

Private Function ArrayToken( _
    ByVal value As Variant, _
    ByVal maximumItems As Long) As String

    On Error GoTo Failed

    If IsEmpty(value) Then
        ArrayToken = "Empty"
        Exit Function
    End If

    If IsNull(value) Then
        ArrayToken = "Null"
        Exit Function
    End If

    If Not IsArray(value) Then
        ArrayToken = "Scalar:" & ProbeToken(CStr(value))
        Exit Function
    End If

    Dim count As Long
    count = VariantItemCount(value)

    Dim result As String
    result = "Count" & CStr(count) & ":["

    Dim shown As Long
    Dim i As Long
    For i = LBound(value) To UBound(value)
        If shown >= maximumItems Then
            result = result & ",..."
            Exit For
        End If

        If shown > 0 Then result = result & ","
        result = result & VariantValueToken(value(i))
        shown = shown + 1
    Next i

    ArrayToken = result & "]"
    Exit Function

Failed:
    ArrayToken = _
        "Unavailable:" & CStr(Err.Number)
End Function

Private Function VariantValueToken(ByVal value As Variant) As String
    If IsObject(value) Then
        Dim target As Object
        Set target = value

        If target Is Nothing Then
            VariantValueToken = "Nothing"
        Else
            VariantValueToken = _
                ProbeToken(TypeName(target)) & ":" & CStr(ObjPtr(target))
        End If
    ElseIf IsNull(value) Then
        VariantValueToken = "Null"
    ElseIf IsEmpty(value) Then
        VariantValueToken = "Empty"
    ElseIf IsNumeric(value) Then
        VariantValueToken = FormatProbeNumber(CDbl(value))
    Else
        VariantValueToken = ProbeToken(CStr(value))
    End If
End Function

Private Function FormatProbeNumber(ByVal value As Double) As String
    FormatProbeNumber = Format$(value, "0.000000000")
End Function

Private Function ProbeToken(ByVal value As String) As String
    ProbeToken = Replace$(value, "|", "/")
    ProbeToken = Replace$(ProbeToken, vbCr, " ")
    ProbeToken = Replace$(ProbeToken, vbLf, " ")
End Function

Private Sub StartProbeLog(ByVal probeName As String)
    On Error GoTo Failed

    mR23DrawingProbeLogPath = _
        R23_LOG_DIRECTORY & "\R23_" & probeName & "_" & _
        Format$(Now, "yyyymmdd_hhnnss") & ".log"

    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open mR23DrawingProbeLogPath For Output As #fileNumber
    Close #fileNumber
    Exit Sub

Failed:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    mR23DrawingProbeLogPath = vbNullString
    Debug.Print "R23_DRAWING_PROBE_LOG_ERROR" & _
        "|operation=Start" & _
        "|error=" & CStr(Err.Number)
    On Error GoTo 0
End Sub

Private Sub ProbeLog(ByVal message As String)
    Debug.Print message

    If Len(mR23DrawingProbeLogPath) = 0 Then Exit Sub
    On Error GoTo Failed

    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open mR23DrawingProbeLogPath For Append As #fileNumber
    Print #fileNumber, message
    Close #fileNumber
    Exit Sub

Failed:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    Debug.Print "R23_DRAWING_PROBE_LOG_ERROR" & _
        "|operation=Append" & _
        "|path=" & ProbeToken(mR23DrawingProbeLogPath) & _
        "|error=" & CStr(Err.Number)
    On Error GoTo 0
End Sub

Private Sub CloseProbeLog()
    mR23DrawingProbeLogPath = vbNullString
End Sub
