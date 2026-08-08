Option Explicit

' Controlled non-modal fallback for Hole Wizard families when
' InsertModelAnnotations4 produces no native IDisplayDimension hole callout.
' Every note is sourced from typed model data and is attached only after an
' edge owned by that same feature maps into an orthographic drawing view.

Private Const POINT_TOL_M As Double = 0.000001
Private Const POSITION_TOL_M As Double = 0.0001
Private Const RADIUS_TOL_M As Double = 0.000001

' swObjectEquality, MCP-confirmed and installed SW2025-validated 2026-08-08.
Private Const swObjectSame As Long = 1

' swDimensionTextParts_e, MCP-confirmed 2026-08-08.
Private Const swDimensionTextPrefix As Long = 1

' swUserPreferenceStringValue_e, installed SW2025 swconst.tlb-confirmed
' 2026-08-08. The API MCP exposed the member but not its numeric value.
Private Const swFileLocationsHoleCalloutFormatFile As Long = 26

Private Type HoleCalloutDefinition
    FeatureName As String
    HoleType As Long
    Quantity As Long
    ThruDiameterM As Double
    TapDrillDiameterM As Double
    TapDrillDepthM As Double
    CounterBoreDiameterM As Double
    CounterBoreDepthM As Double
    ThreadDepthM As Double
    FastenerSize As String
    StandardName As String
    ThreadCallout As String
    QuantitySource As String
    ThreadCalloutSource As String
    CalloutText As String
    DefinitionKind As String
End Type

Private mRequiredFamilies As Long
Private mCreatedCallouts As Long
Private mFailureCount As Long
Private mNativeCalloutsAtStart As Long
Private mNativeCoveredFamilies As Long
Private mDiagnostics As String

Public Sub EnsureControlledHoleCallouts( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal holes As Collection)

    ResetRunState
    On Error GoTo Failed

    Dim requiredNames As Object
    Set requiredNames = CollectRequiredFeatureNames(holes)
    mRequiredFamilies = requiredNames.Count
    mNativeCalloutsAtStart = _
        Module6_QAEngine.CountAllViewHoleCallouts(swDraw)

    AppendDiagnostic "CONTROLLED_CALLOUT_BEGIN|requiredFamilies=" & _
        CStr(mRequiredFamilies) & "|nativeAtStart=" & _
        CStr(mNativeCalloutsAtStart)

    If mRequiredFamilies = 0 Then
        AppendDiagnostic _
            "CONTROLLED_CALLOUT_SKIP|reason=NoHoleWizardFamilies"
        Exit Sub
    End If

    Dim processed As Object
    Set processed = CreateObject("Scripting.Dictionary")
    processed.CompareMode = vbTextCompare

    Dim swFeature As SldWorks.Feature
    Set swFeature = swPart.FirstFeature

    Do While Not swFeature Is Nothing
        Dim featureName As String
        featureName = swFeature.Name

        If requiredNames.Exists(featureName) And _
           Not processed.Exists(featureName) Then

            processed.Add featureName, True

            If NativeCalloutCoversFeature( _
                swApp, swDraw, swFeature) Then

                mNativeCoveredFamilies = mNativeCoveredFamilies + 1
                AppendDiagnostic _
                    "CONTROLLED_CALLOUT_RETAIN_NATIVE|feature=" & _
                    SafeToken(featureName) & _
                    "|identity=AttachedOwnedDrawingEntity"
            Else
                ProcessHoleWizardFeature swApp, swPart, swDrawModel, swDraw, _
                    swFeature, CLng(requiredNames(featureName))
            End If
        End If

        Set swFeature = swFeature.GetNextFeature
    Loop

    If processed.Count <> mRequiredFamilies Then
        mFailureCount = mFailureCount + _
            (mRequiredFamilies - processed.Count)
        AppendDiagnostic _
            "CONTROLLED_CALLOUT_FAILURE|reason=FeatureTraversalShortfall" & _
            "|processed=" & CStr(processed.Count) & _
            "|required=" & CStr(mRequiredFamilies)
    End If

    If mNativeCoveredFamilies < mNativeCalloutsAtStart Then
        mFailureCount = mFailureCount + _
            (mNativeCalloutsAtStart - mNativeCoveredFamilies)
        AppendDiagnostic _
            "CONTROLLED_CALLOUT_FAILURE|reason=NativeCalloutUnattributed" & _
            "|nativeAtStart=" & CStr(mNativeCalloutsAtStart) & _
            "|coveredFamilies=" & CStr(mNativeCoveredFamilies)
    End If

    AppendDiagnostic "CONTROLLED_CALLOUT_END|created=" & _
        CStr(mCreatedCallouts) & _
        "|nativeCoveredFamilies=" & CStr(mNativeCoveredFamilies) & _
        "|failures=" & CStr(mFailureCount)
    Exit Sub

Failed:
    mFailureCount = mFailureCount + 1
    AppendDiagnostic "CONTROLLED_CALLOUT_FATAL|error=" & _
        CStr(Err.Number) & "|description=" & SafeToken(Err.Description)
    On Error Resume Next
    swDrawModel.ClearSelection2 True
End Sub

Private Function NativeCalloutCoversFeature( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swFeature As SldWorks.Feature) As Boolean

    On Error GoTo Failed

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        Dim dimensions As Variant
        dimensions = swView.GetDisplayDimensions

        If IsArray(dimensions) Then
            Dim i As Long
            For i = LBound(dimensions) To UBound(dimensions)
                Dim swDispDim As SldWorks.DisplayDimension
                Set swDispDim = Nothing
                On Error Resume Next
                Set swDispDim = dimensions(i)
                On Error GoTo Failed

                If Not swDispDim Is Nothing Then
                    If Not (swDispDim.IsHoleCallout = False) Then
                        Dim annotation As SldWorks.Annotation
                        Set annotation = swDispDim.GetAnnotation

                        If Not annotation Is Nothing Then
                            Dim attachments As Variant
                            attachments = annotation.GetAttachedEntities3

                            If FeatureOwnsAnyAttachment( _
                                swApp, swFeature, swView, attachments) Then

                                AppendDiagnostic _
                                    "CONTROLLED_CALLOUT_NATIVE_MATCH|feature=" & _
                                    SafeToken(swFeature.Name) & _
                                    "|view=" & SafeToken(swView.Name) & _
                                    "|attachments=" & _
                                    CStr(CountVariantItems(attachments)) & _
                                    "|identity=ISldWorks.IsSame"
                                NativeCalloutCoversFeature = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next i
        End If

        Set swView = swView.GetNextView
    Loop
    Exit Function

Failed:
    AppendDiagnostic _
        "CONTROLLED_CALLOUT_NATIVE_MATCH_ERROR|feature=" & _
        SafeToken(swFeature.Name) & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & SafeToken(Err.Description)
    NativeCalloutCoversFeature = False
End Function

Private Function FeatureOwnsAnyAttachment( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swFeature As SldWorks.Feature, _
    ByRef swView As SldWorks.View, _
    ByVal attachments As Variant) As Boolean

    On Error GoTo Failed
    If Not IsArray(attachments) Then Exit Function

    Dim faces As Variant
    faces = swFeature.GetFaces
    If Not IsArray(faces) Then Exit Function

    Dim f As Long
    For f = LBound(faces) To UBound(faces)
        Dim swFace As SldWorks.Face2
        Set swFace = faces(f)
        If swFace Is Nothing Then GoTo ContinueFace

        Dim edges As Variant
        edges = swFace.GetEdges
        If Not IsArray(edges) Then GoTo ContinueFace

        Dim e As Long
        For e = LBound(edges) To UBound(edges)
            Dim modelEdge As SldWorks.Edge
            Set modelEdge = edges(e)
            If modelEdge Is Nothing Then GoTo ContinueEdge

            Dim mapped As Object
            Set mapped = swView.GetCorrespondingEntity(modelEdge)
            If mapped Is Nothing Then GoTo ContinueEdge

            Dim a As Long
            For a = LBound(attachments) To UBound(attachments)
                Dim attached As Object
                Set attached = Nothing
                On Error Resume Next
                Set attached = attachments(a)
                On Error GoTo Failed

                If Not attached Is Nothing Then
                    If ObjectsAreSame(swApp, mapped, attached) Then
                        FeatureOwnsAnyAttachment = True
                        Exit Function
                    End If
                End If
            Next a

ContinueEdge:
        Next e
ContinueFace:
    Next f
    Exit Function

Failed:
    FeatureOwnsAnyAttachment = False
End Function

Private Function ObjectsAreSame( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal first As Object, _
    ByVal second As Object) As Boolean

    On Error GoTo Failed
    If first Is Nothing Or second Is Nothing Then Exit Function
    ObjectsAreSame = (CLng(swApp.IsSame(first, second)) = swObjectSame)
    Exit Function

Failed:
    ObjectsAreSame = False
End Function

Private Sub ProcessHoleWizardFeature( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swFeature As SldWorks.Feature, _
    ByVal quantity As Long)

    On Error GoTo Failed

    Dim definition As HoleCalloutDefinition
    If Not ReadHoleWizardDefinition( _
        swApp, swPart, swDraw, swFeature, quantity, definition) Then

        mFailureCount = mFailureCount + 1
        AppendDiagnostic _
            "CONTROLLED_CALLOUT_FAILURE|feature=" & _
            SafeToken(swFeature.Name) & _
            "|reason=IncompleteTypedDefinition"
        Exit Sub
    End If

    Dim anchorView As SldWorks.View
    Dim anchorEdge As SldWorks.Edge
    If Not FindOwnedDrawingCircle( _
        swFeature, swDraw, anchorView, anchorEdge) Then

        mFailureCount = mFailureCount + 1
        AppendDiagnostic _
            "CONTROLLED_CALLOUT_FAILURE|feature=" & _
            SafeToken(swFeature.Name) & _
            "|kind=" & definition.DefinitionKind & _
            "|reason=NoMappedOwnedCircularEdge"
        Exit Sub
    End If

    If Not InsertAttachedCallout( _
        swDrawModel, swDraw, anchorView, anchorEdge, definition, _
        mCreatedCallouts) Then

        mFailureCount = mFailureCount + 1
        Exit Sub
    End If

    mCreatedCallouts = mCreatedCallouts + 1
    Exit Sub

Failed:
    mFailureCount = mFailureCount + 1
    AppendDiagnostic "CONTROLLED_CALLOUT_FAILURE|feature=" & _
        SafeToken(swFeature.Name) & "|reason=UnhandledFeatureError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & SafeToken(Err.Description)
End Sub

Private Function ReadHoleWizardDefinition( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swFeature As SldWorks.Feature, _
    ByVal auditQuantity As Long, _
    ByRef definition As HoleCalloutDefinition) As Boolean

    Dim accessGranted As Boolean
    Dim holeData As SldWorks.WizardHoleFeatureData2
    On Error GoTo Failed

    definition.FeatureName = swFeature.Name

    Set holeData = swFeature.GetDefinition
    If holeData Is Nothing Then Exit Function

    accessGranted = Not ( _
        holeData.AccessSelections(swPart, Nothing) = False)
    If Not accessGranted Then Exit Function

    definition.HoleType = holeData.Type
    definition.ThruDiameterM = ReadDoubleMember( _
        holeData, "ThruHoleDiameter")
    definition.TapDrillDiameterM = ReadDoubleMember( _
        holeData, "TapDrillDiameter")
    definition.TapDrillDepthM = ReadDoubleMember( _
        holeData, "TapDrillDepth")
    definition.CounterBoreDiameterM = ReadDoubleMember( _
        holeData, "CounterBoreDiameter")
    definition.CounterBoreDepthM = ReadDoubleMember( _
        holeData, "CounterBoreDepth")
    definition.ThreadDepthM = ReadDoubleMember( _
        holeData, "ThreadDepth")
    definition.FastenerSize = Trim$(CStr(holeData.FastenerSize))
    definition.StandardName = Trim$(CStr(holeData.Standard))
    definition.Quantity = holeData.GetSketchPointCount
    definition.QuantitySource = _
        "IWizardHoleFeatureData2.GetSketchPointCount"

    holeData.ReleaseSelectionAccess
    accessGranted = False

    Dim targetRadiusM As Double
    If definition.CounterBoreDiameterM > 0# Then
        targetRadiusM = definition.ThruDiameterM / 2#
    Else
        targetRadiusM = definition.TapDrillDiameterM / 2#
    End If

    Dim derivedQuantity As Long
    derivedQuantity = CountDerivedPatternInstances( _
        swPart, swFeature, targetRadiusM)
    If derivedQuantity > 0 Then
        definition.Quantity = definition.Quantity + derivedQuantity
        definition.QuantitySource = definition.QuantitySource & _
            "+IFace2.GetSeedFeature/CylinderParams"
    End If

    If definition.CounterBoreDiameterM > 0# Then
        If definition.ThruDiameterM <= 0# Or _
           definition.CounterBoreDepthM <= 0# Then Exit Function

        definition.DefinitionKind = "Counterbore"
        definition.CalloutText = CStr(definition.Quantity) & _
            "x <MOD-DIAM>" & _
            FormatMillimetres(definition.ThruDiameterM) & " THRU" & _
            vbCrLf & "<HOLE-SPOT><MOD-DIAM>" & _
            FormatMillimetres(definition.CounterBoreDiameterM) & _
            " <HOLE-DEPTH>" & _
            FormatMillimetres(definition.CounterBoreDepthM)
    ElseIf definition.ThreadDepthM > 0# Then
        If definition.TapDrillDiameterM <= 0# Or _
           definition.TapDrillDepthM <= 0# Or _
           Len(definition.FastenerSize) = 0 Then Exit Function

        definition.ThreadCallout = FindInstalledThreadCallout( _
            swApp, definition.StandardName, definition.FastenerSize)
        If Len(definition.ThreadCallout) > 0 Then
            definition.ThreadCalloutSource = _
                "InstalledCalloutFormat/TAP-BLIND"
        Else
            definition.ThreadCallout = FindImportedThreadCallout( _
                swDraw, definition.FastenerSize)
            If Len(definition.ThreadCallout) > 0 Then
                definition.ThreadCalloutSource = _
                    "IView.GetDisplayDimensions/GetText(Prefix)"
            Else
                definition.ThreadCallout = FindCosmeticThreadCallout( _
                    swPart, definition.FastenerSize)
                definition.ThreadCalloutSource = _
                    "ICosmeticThreadFeatureData.ThreadCallout"
            End If
        End If
        If Not IsSpecificThreadCallout( _
            definition.ThreadCallout, definition.FastenerSize) Then _
            Exit Function

        definition.DefinitionKind = "Tapped"
        definition.CalloutText = CStr(definition.Quantity) & _
            "x <MOD-DIAM>" & _
            FormatMillimetres(definition.TapDrillDiameterM) & _
            " <HOLE-DEPTH>" & _
            FormatMillimetres(definition.TapDrillDepthM) & vbCrLf & _
            definition.ThreadCallout & " <HOLE-DEPTH>" & _
            FormatMillimetres(definition.ThreadDepthM)
    Else
        Exit Function
    End If

    AppendDiagnostic "CONTROLLED_CALLOUT_DEFINITION|feature=" & _
        SafeToken(definition.FeatureName) & _
        "|type=" & CStr(definition.HoleType) & _
        "|kind=" & definition.DefinitionKind & _
        "|quantity=" & CStr(definition.Quantity) & _
        "|quantitySource=" & definition.QuantitySource & _
        "|legacyAuditQuantity=" & CStr(auditQuantity) & _
        "|thruDiameterM=" & Format$(definition.ThruDiameterM, _
            "0.000000000") & _
        "|tapDrillDiameterM=" & Format$( _
            definition.TapDrillDiameterM, "0.000000000") & _
        "|tapDrillDepthM=" & Format$( _
            definition.TapDrillDepthM, "0.000000000") & _
        "|counterBoreDiameterM=" & Format$( _
            definition.CounterBoreDiameterM, "0.000000000") & _
        "|counterBoreDepthM=" & Format$( _
            definition.CounterBoreDepthM, "0.000000000") & _
        "|threadDepthM=" & Format$( _
            definition.ThreadDepthM, "0.000000000") & _
        "|standard=" & SafeToken(definition.StandardName) & _
        "|threadCallout=" & SafeToken(definition.ThreadCallout) & _
        "|threadCalloutSource=" & definition.ThreadCalloutSource

    ReadHoleWizardDefinition = True
    Exit Function

Failed:
    If accessGranted Then
        On Error Resume Next
        holeData.ReleaseSelectionAccess
        On Error GoTo 0
    End If
End Function

Private Function FindInstalledThreadCallout( _
    ByRef swApp As SldWorks.SldWorks, _
    ByVal standardName As String, _
    ByVal fastenerSize As String) As String

    Dim fileNumber As Integer
    Dim fileOpened As Boolean
    Dim formatPath As String
    On Error GoTo Failed

    formatPath = Trim$(swApp.GetUserPreferenceStringValue( _
        swFileLocationsHoleCalloutFormatFile))
    If Len(formatPath) = 0 Then GoTo SafeExit

    If LCase$(Right$(formatPath, 4)) <> ".txt" Then
        If Right$(formatPath, 1) <> "\" Then _
            formatPath = formatPath & "\"
        formatPath = formatPath & "calloutformat.txt"
    End If
    If Len(Dir$(formatPath)) = 0 Then GoTo SafeExit

    fileNumber = FreeFile
    Open formatPath For Input As #fileNumber
    fileOpened = True

    Dim inRequestedStandard As Boolean
    Dim inTappedHoles As Boolean
    Dim sourceLine As String

    Do Until EOF(fileNumber)
        Dim currentLine As String
        Line Input #fileNumber, currentLine

        Dim upperLine As String
        upperLine = UCase$(Trim$(currentLine))

        If InStr(1, upperLine, "[", vbBinaryCompare) > 0 And _
           InStr(1, upperLine, "]", vbBinaryCompare) > 0 Then

            inRequestedStandard = InStr(1, upperLine, _
                "[" & UCase$(standardName) & "]", _
                vbBinaryCompare) > 0
            inTappedHoles = False
        ElseIf inRequestedStandard And _
               InStr(1, upperLine, "TAPPED HOLES", _
                   vbBinaryCompare) > 0 Then

            inTappedHoles = True
        ElseIf inRequestedStandard And inTappedHoles And _
               Left$(upperLine, 10) = "TAP-BLIND=" Then

            Do While Right$(Trim$(currentLine), 1) = "\" And _
                     Not EOF(fileNumber)
                Line Input #fileNumber, currentLine
                If InStr(1, currentLine, "<hw-threaddesc>", _
                    vbTextCompare) > 0 Then
                    sourceLine = currentLine
                    Exit Do
                End If
            Loop
            Exit Do
        End If
    Loop

    If Len(sourceLine) > 0 Then
        Dim resolved As String
        resolved = Replace$(sourceLine, "<hw-threaddesc>", _
            fastenerSize, 1, 1, vbTextCompare)

        Dim depthAt As Long
        depthAt = InStr(1, resolved, "<HOLE-DEPTH>", vbTextCompare)
        If depthAt > 0 Then resolved = Left$(resolved, depthAt - 1)

        resolved = Trim$(Replace$(resolved, vbTab, " "))
        If IsSpecificThreadCallout(resolved, fastenerSize) Then
            FindInstalledThreadCallout = resolved
        End If
    End If

SafeExit:
    On Error Resume Next
    If fileOpened Then Close #fileNumber
    On Error GoTo 0

    AppendDiagnostic _
        "CONTROLLED_CALLOUT_THREAD_SOURCE|source=InstalledCalloutFormat" & _
        "|standard=" & SafeToken(standardName) & _
        "|path=" & SafeToken(formatPath) & _
        "|found=" & CStr(Len(FindInstalledThreadCallout) > 0)
    Exit Function

Failed:
    AppendDiagnostic _
        "CONTROLLED_CALLOUT_THREAD_SOURCE|source=InstalledCalloutFormat" & _
        "|standard=" & SafeToken(standardName) & _
        "|path=" & SafeToken(formatPath) & _
        "|result=Error|error=" & CStr(Err.Number) & _
        "|description=" & SafeToken(Err.Description)
    Resume SafeExit
End Function

Private Function CountDerivedPatternInstances( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef seedFeature As SldWorks.Feature, _
    ByVal targetRadiusM As Double) As Long

    On Error GoTo Failed
    If targetRadiusM <= 0# Then Exit Function

    Dim cylinderKeys As Object
    Set cylinderKeys = CreateObject("Scripting.Dictionary")
    cylinderKeys.CompareMode = vbTextCompare

    Dim swFeature As SldWorks.Feature
    Set swFeature = swPart.FirstFeature

    Do While Not swFeature Is Nothing
        If StrComp(swFeature.Name, seedFeature.Name, _
            vbTextCompare) <> 0 Then

            Dim faces As Variant
            faces = swFeature.GetFaces
            If IsArray(faces) Then
                Dim faceIndex As Long
                For faceIndex = LBound(faces) To UBound(faces)
                    Dim swFace As SldWorks.Face2
                    Set swFace = faces(faceIndex)
                    If swFace Is Nothing Then GoTo ContinueFace

                    Dim faceSeed As SldWorks.Feature
                    Set faceSeed = Nothing
                    On Error Resume Next
                    Set faceSeed = swFace.GetSeedFeature
                    On Error GoTo Failed
                    If faceSeed Is Nothing Then GoTo ContinueFace
                    If StrComp(faceSeed.Name, seedFeature.Name, _
                        vbTextCompare) <> 0 Then GoTo ContinueFace

                    Dim surface As SldWorks.Surface
                    Set surface = swFace.GetSurface
                    If surface Is Nothing Then GoTo ContinueFace
                    If surface.IsCylinder = False Then GoTo ContinueFace

                    Dim cylinderData As Variant
                    cylinderData = surface.CylinderParams
                    If Not IsArray(cylinderData) Then GoTo ContinueFace
                    If CountVariantItems(cylinderData) <> 7 Then _
                        GoTo ContinueFace
                    If Abs(Abs(CDbl(cylinderData(6))) - _
                        targetRadiusM) > RADIUS_TOL_M Then GoTo ContinueFace

                    Dim cylinderKey As String
                    cylinderKey = CylinderIdentityKey(cylinderData)
                    If Not cylinderKeys.Exists(cylinderKey) Then
                        cylinderKeys.Add cylinderKey, True
                    End If
ContinueFace:
                Next faceIndex
            End If
        End If
        Set swFeature = swFeature.GetNextFeature
    Loop

    CountDerivedPatternInstances = cylinderKeys.Count
    Exit Function

Failed:
    CountDerivedPatternInstances = 0
End Function

Private Function CylinderIdentityKey(ByVal cylinderData As Variant) As String
    CylinderIdentityKey = _
        Format$(CDbl(cylinderData(0)), "0.000000") & "," & _
        Format$(CDbl(cylinderData(1)), "0.000000") & "," & _
        Format$(CDbl(cylinderData(2)), "0.000000") & "|" & _
        Format$(Abs(CDbl(cylinderData(3))), "0.000000") & "," & _
        Format$(Abs(CDbl(cylinderData(4))), "0.000000") & "," & _
        Format$(Abs(CDbl(cylinderData(5))), "0.000000") & "|" & _
        Format$(Abs(CDbl(cylinderData(6))), "0.000000")
End Function

Private Function FindImportedThreadCallout( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal fastenerSize As String) As String

    On Error GoTo Failed

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        Dim dimensions As Variant
        dimensions = swView.GetDisplayDimensions

        If IsArray(dimensions) Then
            Dim i As Long
            For i = LBound(dimensions) To UBound(dimensions)
                Dim displayDimension As SldWorks.DisplayDimension
                Set displayDimension = dimensions(i)
                If displayDimension Is Nothing Then GoTo ContinueDimension

                Dim prefix As String
                prefix = displayDimension.GetText(swDimensionTextPrefix)

                FindImportedThreadCallout = ExtractThreadDesignation( _
                    prefix, fastenerSize)
                If Len(FindImportedThreadCallout) > 0 Then Exit Function
ContinueDimension:
            Next i
        End If

        Set swView = swView.GetNextView
    Loop
    Exit Function

Failed:
    FindImportedThreadCallout = vbNullString
End Function

Private Function ExtractThreadDesignation( _
    ByVal sourceText As String, _
    ByVal fastenerSize As String) As String

    Dim startAt As Long
    startAt = InStr(1, sourceText, fastenerSize, vbTextCompare)
    If startAt = 0 Then Exit Function

    Dim candidate As String
    candidate = Mid$(sourceText, startAt)

    Dim depthAt As Long
    depthAt = InStr(1, candidate, "<HOLE-DEPTH>", vbTextCompare)
    If depthAt > 0 Then candidate = Left$(candidate, depthAt - 1)

    candidate = Trim$(candidate)
    If IsSpecificThreadCallout(candidate, fastenerSize) Then
        ExtractThreadDesignation = candidate
    End If
End Function

Private Function IsSpecificThreadCallout( _
    ByVal calloutText As String, _
    ByVal fastenerSize As String) As Boolean

    Dim compactCallout As String
    Dim compactFastener As String
    compactCallout = UCase$(CompactText(calloutText))
    compactFastener = UCase$(CompactText(fastenerSize))

    If Len(compactCallout) <= Len(compactFastener) Then Exit Function
    If InStr(1, compactCallout, compactFastener, _
        vbBinaryCompare) = 0 Then Exit Function
    If InStr(1, compactCallout, "TAPPEDHOLE", _
        vbBinaryCompare) > 0 Then Exit Function

    IsSpecificThreadCallout = True
End Function

Private Function FindCosmeticThreadCallout( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal fastenerSize As String) As String

    On Error GoTo Failed

    Dim swFeature As SldWorks.Feature
    Set swFeature = swPart.FirstFeature

    Do While Not swFeature Is Nothing
        FindCosmeticThreadCallout = FindThreadCalloutInBranch( _
            swPart, swFeature, fastenerSize)
        If Len(FindCosmeticThreadCallout) > 0 Then Exit Function
        Set swFeature = swFeature.GetNextFeature
    Loop
    Exit Function

Failed:
    FindCosmeticThreadCallout = vbNullString
End Function

Private Function FindThreadCalloutInBranch( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByVal fastenerSize As String) As String

    On Error GoTo Failed

    If UCase$(ResolvedFeatureType(swFeature)) = "COSMETICTHREAD" Then
        FindThreadCalloutInBranch = _
            ReadCosmeticThreadText(swPart, swFeature, fastenerSize)
        If Len(FindThreadCalloutInBranch) > 0 Then Exit Function
    End If

    Dim child As SldWorks.Feature
    Set child = swFeature.GetFirstSubFeature

    Do While Not child Is Nothing
        FindThreadCalloutInBranch = FindThreadCalloutInBranch( _
            swPart, child, fastenerSize)
        If Len(FindThreadCalloutInBranch) > 0 Then Exit Function
        Set child = child.GetNextSubFeature
    Loop
    Exit Function

Failed:
    FindThreadCalloutInBranch = vbNullString
End Function

Private Function ReadCosmeticThreadText( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByVal fastenerSize As String) As String

    Dim accessGranted As Boolean
    Dim threadData As SldWorks.CosmeticThreadFeatureData
    On Error GoTo Failed

    Set threadData = swFeature.GetDefinition
    If threadData Is Nothing Then Exit Function

    accessGranted = Not ( _
        threadData.AccessSelections(swPart, Nothing) = False)
    If Not accessGranted Then Exit Function

    Dim calloutText As String
    calloutText = Trim$(CStr(threadData.ThreadCallout))

    threadData.ReleaseSelectionAccess
    accessGranted = False

    If InStr(1, CompactText(calloutText), _
        CompactText(fastenerSize), vbTextCompare) > 0 Then

        ReadCosmeticThreadText = calloutText
    End If
    Exit Function

Failed:
    If accessGranted Then
        On Error Resume Next
        threadData.ReleaseSelectionAccess
        On Error GoTo 0
    End If
End Function

Private Function FindOwnedDrawingCircle( _
    ByRef swFeature As SldWorks.Feature, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef anchorView As SldWorks.View, _
    ByRef anchorEdge As SldWorks.Edge) As Boolean

    On Error GoTo Failed

    Dim faces As Variant
    faces = swFeature.GetFaces
    If Not IsArray(faces) Then Exit Function

    Dim swView As SldWorks.View
    Dim preferredRole As Long
    preferredRole = PreferredRoleForFeature(swFeature)

    Dim searchPass As Long
    For searchPass = 0 To 1
        Set swView = swDraw.GetFirstView
        If Not swView Is Nothing Then Set swView = swView.GetNextView

        Do While Not swView Is Nothing
            Dim role As Long
            role = Module8_ViewClassifier.ClassifyView(swView)

            If AllowsCalloutAnchor(role) And _
               ((searchPass = 0 And role = preferredRole) Or _
                (searchPass = 1 And role <> preferredRole)) Then

                Dim faceIndex As Long
                For faceIndex = LBound(faces) To UBound(faces)
                    Dim swFace As SldWorks.Face2
                    Set swFace = faces(faceIndex)
                    If swFace Is Nothing Then GoTo ContinueFace

                    Dim edges As Variant
                    edges = swFace.GetEdges
                    If Not IsArray(edges) Then GoTo ContinueFace

                    Dim edgeIndex As Long
                    For edgeIndex = LBound(edges) To UBound(edges)
                        Dim modelEdge As SldWorks.Edge
                        Set modelEdge = edges(edgeIndex)
                        If modelEdge Is Nothing Then GoTo ContinueEdge
                        If Not IsCompleteCircularEdge(modelEdge) Then _
                            GoTo ContinueEdge

                        Dim mapped As Object
                        Set mapped = swView.GetCorrespondingEntity(modelEdge)
                        If mapped Is Nothing Then GoTo ContinueEdge

                        Dim drawingEdge As SldWorks.Edge
                        On Error Resume Next
                        Set drawingEdge = mapped
                        On Error GoTo Failed
                        If drawingEdge Is Nothing Then GoTo ContinueEdge
                        If Not IsCompleteCircularEdge(drawingEdge) Then _
                            GoTo ContinueEdge

                        Set anchorView = swView
                        Set anchorEdge = drawingEdge
                        AppendDiagnostic _
                            "CONTROLLED_CALLOUT_ANCHOR|feature=" & _
                            SafeToken(swFeature.Name) & _
                            "|view=" & SafeToken(swView.Name) & _
                            "|role=" & _
                            Module8_ViewClassifier.RoleName(role) & _
                            "|preferredRole=" & _
                            Module8_ViewClassifier.RoleName(preferredRole) & _
                            "|ownership=IFeature.GetFaces/IFace2.GetEdges" & _
                            "|mapping=IView.GetCorrespondingEntity"
                        FindOwnedDrawingCircle = True
                        Exit Function

ContinueEdge:
                    Next edgeIndex
ContinueFace:
                Next faceIndex
            End If

            Set swView = swView.GetNextView
        Loop
    Next searchPass
    Exit Function

Failed:
    FindOwnedDrawingCircle = False
End Function

Private Function PreferredRoleForFeature( _
    ByRef swFeature As SldWorks.Feature) As Long

    On Error GoTo Failed

    Dim faces As Variant
    faces = swFeature.GetFaces
    If Not IsArray(faces) Then Exit Function

    Dim i As Long
    For i = LBound(faces) To UBound(faces)
        Dim swFace As SldWorks.Face2
        Set swFace = faces(i)
        If swFace Is Nothing Then GoTo ContinueFace

        Dim surface As SldWorks.Surface
        Set surface = swFace.GetSurface
        If surface Is Nothing Then GoTo ContinueFace
        If surface.IsCylinder = False Then GoTo ContinueFace

        Dim cylinderData As Variant
        cylinderData = surface.CylinderParams
        If Not IsArray(cylinderData) Then GoTo ContinueFace
        If CountVariantItems(cylinderData) <> 7 Then GoTo ContinueFace

        Dim axisX As Double
        Dim axisY As Double
        Dim axisZ As Double
        axisX = Abs(CDbl(cylinderData(3)))
        axisY = Abs(CDbl(cylinderData(4)))
        axisZ = Abs(CDbl(cylinderData(5)))

        If axisX >= axisY And axisX >= axisZ Then
            PreferredRoleForFeature = _
                Module8_ViewClassifier.VIEW_ROLE_RIGHT
        ElseIf axisY >= axisX And axisY >= axisZ Then
            PreferredRoleForFeature = _
                Module8_ViewClassifier.VIEW_ROLE_TOP
        Else
            PreferredRoleForFeature = _
                Module8_ViewClassifier.VIEW_ROLE_FRONT
        End If
        Exit Function
ContinueFace:
    Next i
    Exit Function

Failed:
    PreferredRoleForFeature = Module8_ViewClassifier.VIEW_ROLE_UNKNOWN
End Function

Private Function IsCompleteCircularEdge( _
    ByRef swEdge As SldWorks.Edge) As Boolean

    On Error GoTo Failed

    Dim swCurve As SldWorks.Curve
    Set swCurve = swEdge.GetCurve
    If swCurve Is Nothing Then Exit Function
    If swCurve.IsCircle = False Then Exit Function

    Dim params As SldWorks.CurveParamData
    Set params = swEdge.GetCurveParams3
    If params Is Nothing Then Exit Function

    Dim startPoint As Variant
    Dim endPoint As Variant
    startPoint = params.StartPoint
    endPoint = params.EndPoint
    If Not IsArray(startPoint) Or Not IsArray(endPoint) Then Exit Function

    IsCompleteCircularEdge = _
        Abs(CDbl(startPoint(0)) - CDbl(endPoint(0))) <= POINT_TOL_M And _
        Abs(CDbl(startPoint(1)) - CDbl(endPoint(1))) <= POINT_TOL_M And _
        Abs(CDbl(startPoint(2)) - CDbl(endPoint(2))) <= POINT_TOL_M
    Exit Function

Failed:
    IsCompleteCircularEdge = False
End Function

Private Function InsertAttachedCallout( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swView As SldWorks.View, _
    ByRef drawingEdge As SldWorks.Edge, _
    ByRef definition As HoleCalloutDefinition, _
    ByVal laneIndex As Long) As Boolean

    On Error GoTo Failed

    swDrawModel.ClearSelection2 True

    Dim activated As Boolean
    activated = Not (swDraw.ActivateView(swView.Name) = False)
    If Not activated Then GoTo ActivationFailed

    Dim activeView As SldWorks.View
    Set activeView = swDraw.ActiveDrawingView
    If activeView Is Nothing Then GoTo ActivationFailed
    If StrComp(activeView.Name, swView.Name, vbTextCompare) <> 0 Then _
        GoTo ActivationFailed

    Dim selectionManager As SldWorks.SelectionMgr
    Set selectionManager = swDrawModel.SelectionManager
    If selectionManager Is Nothing Then GoTo SelectionFailed

    Dim selectData As SldWorks.SelectData
    Set selectData = selectionManager.CreateSelectData
    If selectData Is Nothing Then GoTo SelectionFailed

    ' ISelectData.View is propertyput on the installed build. Plain assignment
    ' is intentional; using Set raised error 91 in earlier revisions.
    selectData.View = swView

    Dim entity As SldWorks.Entity
    Set entity = drawingEdge

    Dim selected As Boolean
    selected = Not (entity.Select4(False, selectData) = False)
    If Not selected Or _
       selectionManager.GetSelectedObjectCount2(-1) <> 1 Then
        GoTo SelectionFailed
    End If

    Dim calloutNote As SldWorks.Note
    Set calloutNote = swDrawModel.InsertNote(definition.CalloutText)
    swDrawModel.ClearSelection2 True
    If calloutNote Is Nothing Then GoTo NoteFailed

    Dim annotation As SldWorks.Annotation
    Set annotation = calloutNote.GetAnnotation
    If annotation Is Nothing Then GoTo NoteFailed

    Dim attachments As Variant
    attachments = annotation.GetAttachedEntities3
    Dim attachmentCount As Long
    attachmentCount = CountVariantItems(attachments)
    If attachmentCount < 1 Then GoTo AttachmentFailed

    Dim leaderCount As Long
    leaderCount = annotation.GetLeaderCount
    If leaderCount < 1 Then GoTo AttachmentFailed

    Dim targetX As Double
    Dim targetY As Double
    If Not ResolveNotePosition( _
        swView, laneIndex, targetX, targetY) Then GoTo PositionFailed

    Dim positionReturned As Boolean
    positionReturned = Not ( _
        annotation.SetPosition2(targetX, targetY, 0#) = False)

    swDrawModel.ForceRebuild3 False

    Dim readback As Variant
    readback = annotation.GetPosition
    If Not IsArray(readback) Then GoTo PositionFailed
    If Abs(CDbl(readback(0)) - targetX) > POSITION_TOL_M Or _
       Abs(CDbl(readback(1)) - targetY) > POSITION_TOL_M Then
        GoTo PositionFailed
    End If

    AppendDiagnostic "CONTROLLED_CALLOUT_CREATED|feature=" & _
        SafeToken(definition.FeatureName) & _
        "|kind=" & definition.DefinitionKind & _
        "|view=" & SafeToken(swView.Name) & _
        "|attachments=" & CStr(attachmentCount) & _
        "|leaders=" & CStr(leaderCount) & _
        "|positionReturned=" & CStr(positionReturned) & _
        "|positionX=" & Format$(CDbl(readback(0)), "0.000000") & _
        "|positionY=" & Format$(CDbl(readback(1)), "0.000000") & _
        "|text=" & SafeToken(definition.CalloutText)

    InsertAttachedCallout = True
    Exit Function

ActivationFailed:
    AppendDiagnostic "CONTROLLED_CALLOUT_FAILURE|feature=" & _
        SafeToken(definition.FeatureName) & _
        "|reason=ViewActivationFailed|view=" & SafeToken(swView.Name)
    Exit Function

SelectionFailed:
    AppendDiagnostic "CONTROLLED_CALLOUT_FAILURE|feature=" & _
        SafeToken(definition.FeatureName) & _
        "|reason=OwnedDrawingEdgeSelectionFailed|view=" & _
        SafeToken(swView.Name)
    swDrawModel.ClearSelection2 True
    Exit Function

NoteFailed:
    AppendDiagnostic "CONTROLLED_CALLOUT_FAILURE|feature=" & _
        SafeToken(definition.FeatureName) & _
        "|reason=InsertNoteOrAnnotationFailed|view=" & _
        SafeToken(swView.Name)
    Exit Function

AttachmentFailed:
    AppendDiagnostic "CONTROLLED_CALLOUT_FAILURE|feature=" & _
        SafeToken(definition.FeatureName) & _
        "|reason=AttachmentOrLeaderReadbackFailed|view=" & _
        SafeToken(swView.Name)
    Exit Function

PositionFailed:
    AppendDiagnostic "CONTROLLED_CALLOUT_FAILURE|feature=" & _
        SafeToken(definition.FeatureName) & _
        "|reason=PositionReadbackFailed|view=" & SafeToken(swView.Name)
    Exit Function

Failed:
    On Error Resume Next
    swDrawModel.ClearSelection2 True
    AppendDiagnostic "CONTROLLED_CALLOUT_FAILURE|feature=" & _
        SafeToken(definition.FeatureName) & _
        "|reason=UnhandledInsertError|error=" & CStr(Err.Number) & _
        "|description=" & SafeToken(Err.Description)
End Function

Private Function ResolveNotePosition( _
    ByRef swView As SldWorks.View, _
    ByVal laneIndex As Long, _
    ByRef targetX As Double, _
    ByRef targetY As Double) As Boolean

    On Error GoTo Failed

    Dim outline As Variant
    outline = swView.GetOutline
    If Not IsArray(outline) Then Exit Function
    If CountVariantItems(outline) <> 4 Then Exit Function

    Dim role As Long
    role = Module8_ViewClassifier.ClassifyView(swView)

    Select Case role
        Case Module8_ViewClassifier.VIEW_ROLE_FRONT
            targetX = CDbl(outline(2)) + 0.02
            targetY = CDbl(outline(3)) + 0.02 + _
                (CDbl(laneIndex) * 0.018)
        Case Module8_ViewClassifier.VIEW_ROLE_TOP, _
             Module8_ViewClassifier.VIEW_ROLE_BOTTOM
            targetX = CDbl(outline(0)) + 0.012
            targetY = CDbl(outline(1)) - 0.008 - _
                (CDbl(laneIndex) * 0.014)
        Case Else
            targetX = CDbl(outline(2)) + 0.008
            targetY = CDbl(outline(3)) - 0.002 - _
                (CDbl(laneIndex) * 0.014)
    End Select
    ResolveNotePosition = True
    Exit Function

Failed:
    ResolveNotePosition = False
End Function

Private Function CollectRequiredFeatureNames( _
    ByVal holes As Collection) As Object

    Dim names As Object
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbTextCompare

    On Error GoTo SafeExit
    If holes Is Nothing Then GoTo SafeExit

    Dim i As Long
    For i = 1 To holes.Count
        If Module3_ModelAudit.IsHoleWizardItem(holes(i)) Then
            Dim featureName As String
            featureName = Module3_ModelAudit.GetHoleFeatureName(holes(i))
            If Len(featureName) > 0 Then
                If names.Exists(featureName) Then
                    names(featureName) = CLng(names(featureName)) + 1
                Else
                    names.Add featureName, 1
                End If
            End If
        End If
    Next i

SafeExit:
    Set CollectRequiredFeatureNames = names
End Function

Private Function AllowsCalloutAnchor(ByVal role As Long) As Boolean
    Select Case role
        Case Module8_ViewClassifier.VIEW_ROLE_FRONT, _
             Module8_ViewClassifier.VIEW_ROLE_TOP, _
             Module8_ViewClassifier.VIEW_ROLE_BOTTOM, _
             Module8_ViewClassifier.VIEW_ROLE_RIGHT, _
             Module8_ViewClassifier.VIEW_ROLE_LEFT, _
             Module8_ViewClassifier.VIEW_ROLE_BACK

            AllowsCalloutAnchor = True
    End Select
End Function

Private Function ResolvedFeatureType( _
    ByRef swFeature As SldWorks.Feature) As String

    On Error Resume Next
    ResolvedFeatureType = swFeature.GetTypeName2
    If UCase$(ResolvedFeatureType) = "ICE" Or _
       Len(ResolvedFeatureType) = 0 Then
        ResolvedFeatureType = swFeature.GetTypeName
    End If
    On Error GoTo 0
End Function

Private Function ReadDoubleMember( _
    ByVal source As Object, _
    ByVal memberName As String) As Double

    On Error GoTo Failed
    Dim value As Variant
    value = CallByName(source, memberName, VbGet)
    If IsNumeric(value) Then ReadDoubleMember = CDbl(value)
Failed:
End Function

Private Function FormatMillimetres(ByVal valueM As Double) As String
    Dim rendered As String
    rendered = Format$(valueM * 1000#, "0.000")

    Do While Len(rendered) > 0 And Right$(rendered, 1) = "0"
        rendered = Left$(rendered, Len(rendered) - 1)
    Loop

    If Len(rendered) > 0 Then
        If Right$(rendered, 1) = "." Or _
           Right$(rendered, 1) = "," Then
            rendered = Left$(rendered, Len(rendered) - 1)
        End If
    End If

    FormatMillimetres = rendered
End Function

Private Function CompactText(ByVal value As String) As String
    CompactText = Replace$(Trim$(value), " ", vbNullString)
End Function

Private Function CountVariantItems(ByVal items As Variant) As Long
    On Error GoTo SafeExit
    If IsArray(items) Then
        CountVariantItems = UBound(items) - LBound(items) + 1
    ElseIf Not IsEmpty(items) And Not IsNull(items) Then
        CountVariantItems = 1
    End If
SafeExit:
End Function

Private Sub ResetRunState()
    mRequiredFamilies = 0
    mCreatedCallouts = 0
    mFailureCount = 0
    mNativeCalloutsAtStart = 0
    mNativeCoveredFamilies = 0
    mDiagnostics = vbNullString
End Sub

Private Sub AppendDiagnostic(ByVal value As String)
    If Len(mDiagnostics) > 0 Then mDiagnostics = mDiagnostics & vbCrLf
    mDiagnostics = mDiagnostics & value
    Debug.Print value
End Sub

Private Function SafeToken(ByVal value As String) As String
    SafeToken = Replace$(value, "|", "/")
    SafeToken = Replace$(SafeToken, vbCr, "/")
    SafeToken = Replace$(SafeToken, vbLf, "/")
End Function

Public Function RequiredFamilyCount() As Long
    RequiredFamilyCount = mRequiredFamilies
End Function

Public Function CreatedControlledCalloutCount() As Long
    CreatedControlledCalloutCount = mCreatedCallouts
End Function

Public Function NativeCoveredFamilyCount() As Long
    NativeCoveredFamilyCount = mNativeCoveredFamilies
End Function

Public Function ControlledCalloutFailureCount() As Long
    ControlledCalloutFailureCount = mFailureCount
End Function

Public Function DescribeControlledCallouts() As String
    DescribeControlledCallouts = "Controlled hole-callout transactions:" & _
        vbCrLf
    If Len(mDiagnostics) = 0 Then
        DescribeControlledCallouts = DescribeControlledCallouts & _
            "  none recorded"
    Else
        DescribeControlledCallouts = DescribeControlledCallouts & _
            "  " & Replace$(mDiagnostics, vbCrLf, vbCrLf & "  ")
    End If
End Function

' Compile-failure localisation no-op called by
' Module20_ProbeRunner.R23_TouchAllModules.
Public Sub R23_CompileTouch()
End Sub
