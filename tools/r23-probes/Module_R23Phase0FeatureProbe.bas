Option Explicit

Private Const R23_PROBE_BUILD As String = "20260731.2-curve-and-suppression-evidence"

' R23 Phase 0 read-only probe.
' This module does not modify a feature definition, rebuild, or save a model.
' Run only with one of the three authorized fixture parts active.

Private Const swDocPART As Long = 1
Private Const swSpecifyConfiguration As Long = 3
Private Const MAX_FEATURE_DEPTH As Long = 64
Private Const R23_LOG_DIRECTORY As String = _
    "C:\Users\V.T\Documents\VBA 3D TO 2D\test_assets\iteration_evidence\r23\20260730-075811\live-probes"
Private mR23ProbeLogPath As String

Public Sub R23_ProbeActivePartFeaturesAndCurves()
    On Error GoTo Failed

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Debug.Print "R23_PROBE_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swPart As SldWorks.ModelDoc2
    Set swPart = swApp.ActiveDoc

    If swPart Is Nothing Then
        Debug.Print "R23_PROBE_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swPart.GetType <> swDocPART Then
        Debug.Print "R23_PROBE_FATAL|reason=ActiveDocumentNotPart"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not IsAuthorizedProbeFixture(partPath) Then
        Debug.Print "R23_PROBE_FATAL|reason=UnauthorizedFixture|path=" & _
            EvidenceToken(partPath)
        Exit Sub
    End If

    Dim configurationName As String
    configurationName = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    StartProbeLog

    ProbeLog "R23_PROBE_BEGIN|build=" & R23_PROBE_BUILD & _
        "|name=FeatureAndCurveContract" & _
        "|part=" & EvidenceToken(partPath) & _
        "|configuration=" & EvidenceToken(configurationName) & _
        "|solidWorksRevision=" & EvidenceToken(swApp.RevisionNumber) & _
        "|logPath=" & EvidenceToken(mR23ProbeLogPath)

    Dim visited As Object
    Set visited = CreateObject("Scripting.Dictionary")
    visited.CompareMode = vbTextCompare

    Dim swFeature As SldWorks.Feature
    Set swFeature = swPart.FirstFeature

    Do While Not swFeature Is Nothing
        ProbeFeatureTree _
            swPart, swFeature, configurationName, visited, 0
        Set swFeature = swFeature.GetNextFeature
    Loop

    ProbeLog "R23_PROBE_END|name=FeatureAndCurveContract" & _
        "|visitedFeatures=" & CStr(visited.Count) & _
        "|status=COMPLETE"
    CloseProbeLog
    Exit Sub

Failed:
    ProbeLog "R23_PROBE_FATAL|reason=UnhandledError" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & EvidenceToken(Err.Description)
    CloseProbeLog
End Sub

Private Sub ProbeFeatureTree( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByVal configurationName As String, _
    ByRef visited As Object, _
    ByVal depth As Long)

    If swFeature Is Nothing Then Exit Sub

    If depth > MAX_FEATURE_DEPTH Then
        ProbeLog "R23_FEATURE_REJECT|reason=MaximumDepthExceeded" & _
            "|feature=" & EvidenceToken(SafeFeatureName(swFeature))
        Exit Sub
    End If

    Dim featureKey As String
    ' ObjPtr identifies the transient VBA COM wrapper, not a durable
    ' SOLIDWORKS feature identity. Different feature wrappers can reuse the
    ' same address while walking the tree, so include stable diagnostics in
    ' the recursion guard.
    featureKey = _
        SafeFeatureName(swFeature) & "|" & _
        SafeGetTypeName2(swFeature) & "|" & _
        SafeObjectKey(swFeature)

    If visited.Exists(featureKey) Then Exit Sub
    visited.Add featureKey, True

    ProbeOneFeature _
        swPart, swFeature, configurationName, depth

    On Error GoTo SubfeatureReadFailed

    Dim subFeature As SldWorks.Feature
    Set subFeature = swFeature.GetFirstSubFeature

    Do While Not subFeature Is Nothing
        ProbeFeatureTree _
            swPart, subFeature, configurationName, visited, depth + 1
        Set subFeature = subFeature.GetNextSubFeature
    Loop
    Exit Sub

SubfeatureReadFailed:
    ProbeLog "R23_FEATURE_SUBTREE_ERROR" & _
        "|feature=" & EvidenceToken(SafeFeatureName(swFeature)) & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & EvidenceToken(Err.Description)
End Sub

Private Sub ProbeOneFeature( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByVal configurationName As String, _
    ByVal depth As Long)

    On Error GoTo Failed

    Dim featureName As String
    Dim rawType2 As String
    Dim rawType1 As String
    Dim effectiveType As String
    Dim suppressionState As String

    featureName = SafeFeatureName(swFeature)
    rawType2 = SafeGetTypeName2(swFeature)
    rawType1 = SafeGetTypeName(swFeature)
    effectiveType = NormalizeEffectiveType(rawType2, rawType1)
    suppressionState = ReadSuppressionState( _
        swFeature, configurationName)

    Dim definition As Object
    Dim definitionType As String
    Set definition = SafeGetDefinition(swFeature)
    definitionType = SafeObjectTypeName(definition)

    ProbeLog "R23_FEATURE" & _
        "|depth=" & CStr(depth) & _
        "|name=" & EvidenceToken(featureName) & _
        "|rawType2=" & EvidenceToken(rawType2) & _
        "|rawType1=" & EvidenceToken(rawType1) & _
        "|effectiveType=" & EvidenceToken(effectiveType) & _
        "|definitionType=" & EvidenceToken(definitionType) & _
        "|configuration=" & EvidenceToken(configurationName) & _
        "|suppression=" & EvidenceToken(suppressionState)

    If Len(effectiveType) = 0 Then
        ProbeLog "R23_FEATURE_REJECT" & _
            "|feature=" & EvidenceToken(featureName) & _
            "|reason=EffectiveTypeUnresolved"
    Else
        Select Case UCase$(effectiveType)
            Case "CUT", "CUTTHIN", "EXTRUSION"
                ProbeExtrudeDefinition _
                    swPart, swFeature, definition, effectiveType

            Case "HOLEWZD"
                ProbeWizardHoleDefinition _
                    swPart, swFeature, definition
        End Select
    End If

    ProbeOwnedFaces swFeature, effectiveType
    Exit Sub

Failed:
    ProbeLog "R23_FEATURE_ERROR" & _
        "|feature=" & EvidenceToken(SafeFeatureName(swFeature)) & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & EvidenceToken(Err.Description)
End Sub

Private Sub ProbeExtrudeDefinition( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As Object, _
    ByVal effectiveType As String)

    Dim featureName As String
    featureName = SafeFeatureName(swFeature)

    If definition Is Nothing Then
        ProbeLog "R23_EXTRUDE_DEFINITION" & _
            "|feature=" & EvidenceToken(featureName) & _
            "|effectiveType=" & EvidenceToken(effectiveType) & _
            "|typed=False|reason=DefinitionNothing"
        Exit Sub
    End If

    On Error GoTo TypeMismatch

    Dim extrudeData As SldWorks.ExtrudeFeatureData2
    Set extrudeData = definition

    If extrudeData Is Nothing Then GoTo TypeMismatch

    Dim selectionsAccessed As Boolean
    selectionsAccessed = CBool( _
        extrudeData.AccessSelections(swPart, Nothing))

    ProbeLog "R23_EXTRUDE_DEFINITION" & _
        "|feature=" & EvidenceToken(featureName) & _
        "|effectiveType=" & EvidenceToken(effectiveType) & _
        "|typed=True" & _
        "|accessSelections=" & CStr(selectionsAccessed)

    If Not selectionsAccessed Then Exit Sub

    On Error GoTo ReadFailed

    Dim contourCount As Long
    Dim contourState As String
    Dim profileSubfeatureCount As Long
    ReadExtrudeContourState _
        extrudeData, contourCount, contourState
    profileSubfeatureCount = CountProfileSubfeatures(swFeature)

    ProbeLog "R23_EXTRUDE_DATA" & _
        "|feature=" & EvidenceToken(featureName) & _
        "|isBoss=" & CStr(CBool(extrudeData.IsBossFeature)) & _
        "|forwardDepthM=" & FormatProbeNumber( _
            extrudeData.GetDepth(True)) & _
        "|reverseDepthM=" & FormatProbeNumber( _
            extrudeData.GetDepth(False)) & _
        "|forwardEndCondition=" & CStr( _
            extrudeData.GetEndCondition(True)) & _
        "|reverseEndCondition=" & CStr( _
            extrudeData.GetEndCondition(False)) & _
        "|contourCount=" & CStr(contourCount) & _
        "|contours=" & EvidenceToken(contourState) & _
        "|profileSubfeatures=" & CStr(profileSubfeatureCount)

ReleaseSelections:
    On Error Resume Next
    extrudeData.ReleaseSelectionAccess
    If Err.Number <> 0 Then
        ProbeLog "R23_EXTRUDE_RELEASE_ERROR" & _
            "|feature=" & EvidenceToken(featureName) & _
            "|error=" & CStr(Err.Number) & _
            "|description=" & EvidenceToken(Err.Description)
    Else
        ProbeLog "R23_EXTRUDE_RELEASE" & _
            "|feature=" & EvidenceToken(featureName) & _
            "|status=SUCCESS"
    End If
    On Error GoTo 0
    Exit Sub

ReadFailed:
    ProbeLog "R23_EXTRUDE_READ_ERROR" & _
        "|feature=" & EvidenceToken(featureName) & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & EvidenceToken(Err.Description)
    Resume ReleaseSelections

TypeMismatch:
    ProbeLog "R23_EXTRUDE_DEFINITION" & _
        "|feature=" & EvidenceToken(featureName) & _
        "|effectiveType=" & EvidenceToken(effectiveType) & _
        "|typed=False|reason=IExtrudeFeatureData2Unavailable" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & EvidenceToken(Err.Description)
End Sub

Private Sub ReadExtrudeContourState( _
    ByRef extrudeData As SldWorks.ExtrudeFeatureData2, _
    ByRef contourCount As Long, _
    ByRef contourState As String)

    On Error GoTo Failed

    contourCount = extrudeData.GetContoursCount

    Dim contours As Variant
    contours = extrudeData.Contours

    If IsArray(contours) Then
        contourState = "ArrayCount:" & CStr(VariantItemCount(contours))
    ElseIf IsEmpty(contours) Then
        contourState = "Empty"
    ElseIf IsNull(contours) Then
        contourState = "Null"
    ElseIf IsObject(contours) Then
        contourState = "Object:" & SafeObjectTypeName(contours)
    Else
        contourState = "ScalarType:" & CStr(VarType(contours))
    End If
    Exit Sub

Failed:
    contourCount = -1
    contourState = _
        "Error:" & CStr(Err.Number) & ":" & Err.Description
End Sub

Private Function CountProfileSubfeatures( _
    ByRef swFeature As SldWorks.Feature) As Long

    On Error GoTo Failed

    Dim child As SldWorks.Feature
    Set child = swFeature.GetFirstSubFeature

    Do While Not child Is Nothing
        If UCase$(SafeGetTypeName2(child)) = "PROFILEFEATURE" Then
            CountProfileSubfeatures = CountProfileSubfeatures + 1
        End If
        Set child = child.GetNextSubFeature
    Loop
    Exit Function

Failed:
    CountProfileSubfeatures = -1
End Function

Private Sub ProbeWizardHoleDefinition( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swFeature As SldWorks.Feature, _
    ByRef definition As Object)

    Dim featureName As String
    featureName = SafeFeatureName(swFeature)

    If definition Is Nothing Then
        ProbeLog "R23_HOLEWZD_DEFINITION" & _
            "|feature=" & EvidenceToken(featureName) & _
            "|typed=False|reason=DefinitionNothing"
        Exit Sub
    End If

    On Error GoTo TypeMismatch

    Dim holeData As SldWorks.WizardHoleFeatureData2
    Set holeData = definition

    If holeData Is Nothing Then GoTo TypeMismatch

    Dim selectionsAccessed As Boolean
    selectionsAccessed = CBool( _
        holeData.AccessSelections(swPart, Nothing))

    ProbeLog "R23_HOLEWZD_DEFINITION" & _
        "|feature=" & EvidenceToken(featureName) & _
        "|typed=True" & _
        "|accessSelections=" & CStr(selectionsAccessed)

    If Not selectionsAccessed Then Exit Sub

    On Error GoTo ReadFailed

    ProbeLog "R23_HOLEWZD_DATA" & _
        "|feature=" & EvidenceToken(featureName) & _
        "|type=" & ReadObjectProperty(definition, "Type") & _
        "|standard2=" & ReadObjectProperty(definition, "Standard2") & _
        "|fastenerSize=" & ReadObjectProperty( _
            definition, "FastenerSize") & _
        "|holeDiameterM=" & ReadObjectProperty( _
            definition, "HoleDiameter") & _
        "|holeDepthM=" & ReadObjectProperty( _
            definition, "HoleDepth") & _
        "|counterBoreDiameterM=" & ReadObjectProperty( _
            definition, "CounterBoreDiameter") & _
        "|counterBoreDepthM=" & ReadObjectProperty( _
            definition, "CounterBoreDepth") & _
        "|threadDiameterM=" & ReadObjectProperty( _
            definition, "ThreadDiameter") & _
        "|threadDepthM=" & ReadObjectProperty( _
            definition, "ThreadDepth") & _
        "|threadClass=" & ReadObjectProperty( _
            definition, "ThreadClass") & _
        "|sketchPointCount=" & ReadObjectMethod0( _
            definition, "GetSketchPointCount")

ReleaseSelections:
    On Error Resume Next
    holeData.ReleaseSelectionAccess
    If Err.Number <> 0 Then
        ProbeLog "R23_HOLEWZD_RELEASE_ERROR" & _
            "|feature=" & EvidenceToken(featureName) & _
            "|error=" & CStr(Err.Number) & _
            "|description=" & EvidenceToken(Err.Description)
    Else
        ProbeLog "R23_HOLEWZD_RELEASE" & _
            "|feature=" & EvidenceToken(featureName) & _
            "|status=SUCCESS"
    End If
    On Error GoTo 0
    Exit Sub

ReadFailed:
    ProbeLog "R23_HOLEWZD_READ_ERROR" & _
        "|feature=" & EvidenceToken(featureName) & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & EvidenceToken(Err.Description)
    Resume ReleaseSelections

TypeMismatch:
    ProbeLog "R23_HOLEWZD_DEFINITION" & _
        "|feature=" & EvidenceToken(featureName) & _
        "|typed=False|reason=IWizardHoleFeatureData2Unavailable" & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & EvidenceToken(Err.Description)
End Sub

Private Sub ProbeOwnedFaces( _
    ByRef swFeature As SldWorks.Feature, _
    ByVal effectiveType As String)

    On Error GoTo Failed

    Dim faces As Variant
    faces = swFeature.GetFaces

    Dim faceCount As Long
    faceCount = VariantItemCount(faces)

    Dim cylinderCount As Long
    Dim i As Long

    If faceCount > 0 Then
        For i = LBound(faces) To UBound(faces)
            Dim swFace As SldWorks.Face2
            Set swFace = faces(i)

            If Not swFace Is Nothing Then
                If FaceIsCylindrical(swFace) Then
                    cylinderCount = cylinderCount + 1
                    ProbeCylindricalFace _
                        swFeature, swFace, i - LBound(faces)
                End If
            End If
        Next i
    End If

    ProbeLog "R23_FEATURE_FACES" & _
        "|feature=" & EvidenceToken(SafeFeatureName(swFeature)) & _
        "|effectiveType=" & EvidenceToken(effectiveType) & _
        "|ownedFaces=" & CStr(faceCount) & _
        "|ownedCylinders=" & CStr(cylinderCount)
    Exit Sub

Failed:
    ProbeLog "R23_FEATURE_FACE_ERROR" & _
        "|feature=" & EvidenceToken(SafeFeatureName(swFeature)) & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & EvidenceToken(Err.Description)
End Sub

Private Sub ProbeCylindricalFace( _
    ByRef swFeature As SldWorks.Feature, _
    ByRef swFace As SldWorks.Face2, _
    ByVal faceIndex As Long)

    On Error GoTo Failed

    Dim swSurface As SldWorks.Surface
    Set swSurface = swFace.GetSurface

    Dim cylinderData As Variant
    cylinderData = swSurface.CylinderParams

    Dim cylinderRadius As String
    cylinderRadius = "Unavailable"

    If IsArray(cylinderData) Then
        If VariantItemCount(cylinderData) = 7 Then
            cylinderRadius = FormatProbeNumber( _
                CDbl(cylinderData(LBound(cylinderData) + 6)))
        End If
    End If

    Dim seedName As String
    seedName = ReadSeedFeatureName(swFace)

    ProbeLog "R23_CYLINDER_FACE" & _
        "|feature=" & EvidenceToken(SafeFeatureName(swFeature)) & _
        "|faceIndex=" & CStr(faceIndex) & _
        "|faceInSurfaceSense=" & CStr(CBool( _
            swFace.FaceInSurfaceSense)) & _
        "|cylinderParamCount=" & CStr( _
            VariantItemCount(cylinderData)) & _
        "|radiusM=" & cylinderRadius & _
        "|seedFeature=" & EvidenceToken(seedName)

    Dim edges As Variant
    edges = swFace.GetEdges

    If VariantItemCount(edges) = 0 Then Exit Sub

    Dim i As Long
    For i = LBound(edges) To UBound(edges)
        Dim swEdge As SldWorks.Edge
        Set swEdge = edges(i)

        If Not swEdge Is Nothing Then
            ProbeCurveReadOrders _
                swFeature, faceIndex, i - LBound(edges), swEdge
        End If
    Next i
    Exit Sub

Failed:
    ProbeLog "R23_CYLINDER_FACE_ERROR" & _
        "|feature=" & EvidenceToken(SafeFeatureName(swFeature)) & _
        "|faceIndex=" & CStr(faceIndex) & _
        "|error=" & CStr(Err.Number) & _
        "|description=" & EvidenceToken(Err.Description)
End Sub

Private Sub ProbeCurveReadOrders( _
    ByRef swFeature As SldWorks.Feature, _
    ByVal faceIndex As Long, _
    ByVal edgeIndex As Long, _
    ByRef swEdge As SldWorks.Edge)

    Dim paramsFirstResult As String
    Dim circleFirstResult As String

    paramsFirstResult = ReadCurveContract(swEdge, True)
    circleFirstResult = ReadCurveContract(swEdge, False)

    ProbeLog "R23_CURVE_ORDER" & _
        "|feature=" & EvidenceToken(SafeFeatureName(swFeature)) & _
        "|faceIndex=" & CStr(faceIndex) & _
        "|edgeIndex=" & CStr(edgeIndex) & _
        "|paramsFirst=" & EvidenceToken(paramsFirstResult) & _
        "|circleFirst=" & EvidenceToken(circleFirstResult)
End Sub

Private Function ReadCurveContract( _
    ByRef swEdge As SldWorks.Edge, _
    ByVal parametersFirst As Boolean) As String

    On Error GoTo Failed

    Dim swCurve As SldWorks.Curve
    Set swCurve = swEdge.GetCurve

    If swCurve Is Nothing Then
        ReadCurveContract = "CurveNothing"
        Exit Function
    End If

    Dim curveParameters As SldWorks.CurveParamData
    Dim circleData As Variant
    Dim isCircleBefore As Boolean
    Dim isCircleAfter As Boolean
    Dim parameterState As String
    Dim circleState As String

    If parametersFirst Then
        Set curveParameters = swEdge.GetCurveParams3
        parameterState = CurveParameterState(curveParameters)

        isCircleBefore = CBool(swCurve.IsCircle)
        circleState = ReadCircleState(swCurve, isCircleBefore, circleData)
    Else
        isCircleBefore = CBool(swCurve.IsCircle)
        circleState = ReadCircleState(swCurve, isCircleBefore, circleData)

        Set curveParameters = swEdge.GetCurveParams3
        parameterState = CurveParameterState(curveParameters)
    End If

    isCircleAfter = CBool(swCurve.IsCircle)

    ReadCurveContract = _
        "isCircleBefore=" & CStr(isCircleBefore) & _
        ",isCircleAfter=" & CStr(isCircleAfter) & _
        ",circle=" & circleState & _
        ",parameters=" & parameterState
    Exit Function

Failed:
    ReadCurveContract = _
        "Error:" & CStr(Err.Number) & ":" & EvidenceToken(Err.Description)
End Function

Private Function ReadCircleState( _
    ByRef swCurve As SldWorks.Curve, _
    ByVal isCircle As Boolean, _
    ByRef circleData As Variant) As String

    If Not isCircle Then
        ReadCircleState = "SkippedNotCircle"
        Exit Function
    End If

    On Error GoTo Failed

    circleData = swCurve.CircleParams

    If Not IsArray(circleData) Then
        ReadCircleState = "NotArray"
    ElseIf VariantItemCount(circleData) <> 7 Then
        ReadCircleState = _
            "WrongCount:" & CStr(VariantItemCount(circleData))
    Else
        ReadCircleState = "Count7,RadiusM=" & FormatProbeNumber( _
            CDbl(circleData(LBound(circleData) + 6)))
    End If
    Exit Function

Failed:
    ReadCircleState = _
        "Error:" & CStr(Err.Number) & ":" & EvidenceToken(Err.Description)
End Function

Private Function CurveParameterState( _
    ByRef curveParameters As SldWorks.CurveParamData) As String

    If curveParameters Is Nothing Then
        CurveParameterState = "Nothing"
        Exit Function
    End If

    On Error GoTo Failed

    Dim startPoint As Variant
    Dim endPoint As Variant
    startPoint = curveParameters.StartPoint
    endPoint = curveParameters.EndPoint

    Dim endpointDistance As Double
    endpointDistance = PointDistance(startPoint, endPoint)

    CurveParameterState = _
        "UMin=" & FormatProbeNumber(curveParameters.UMinValue) & _
        ",UMax=" & FormatProbeNumber(curveParameters.UMaxValue) & _
        ",EndpointDistanceM=" & FormatProbeNumber(endpointDistance)
    Exit Function

Failed:
    CurveParameterState = _
        "Error:" & CStr(Err.Number) & ":" & EvidenceToken(Err.Description)
End Function

Private Function PointDistance( _
    ByVal firstPoint As Variant, _
    ByVal secondPoint As Variant) As Double

    If Not IsArray(firstPoint) Or Not IsArray(secondPoint) Then
        PointDistance = -1#
        Exit Function
    End If

    If VariantItemCount(firstPoint) < 3 Or _
       VariantItemCount(secondPoint) < 3 Then

        PointDistance = -1#
        Exit Function
    End If

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
End Function

Private Function FaceIsCylindrical( _
    ByRef swFace As SldWorks.Face2) As Boolean

    If swFace Is Nothing Then Exit Function

    On Error GoTo Failed

    Dim swSurface As SldWorks.Surface
    Set swSurface = swFace.GetSurface

    If swSurface Is Nothing Then Exit Function
    FaceIsCylindrical = CBool(swSurface.IsCylinder)
    Exit Function

Failed:
    FaceIsCylindrical = False
End Function

Private Function ReadSeedFeatureName( _
    ByRef swFace As SldWorks.Face2) As String

    On Error Resume Next

    Dim seedFeature As SldWorks.Feature
    Set seedFeature = swFace.GetSeedFeature

    If seedFeature Is Nothing Then
        Set seedFeature = swFace.GetPatternSeedFeature
    End If

    If Err.Number <> 0 Then
        ReadSeedFeatureName = _
            "Error:" & CStr(Err.Number) & ":" & Err.Description
    ElseIf seedFeature Is Nothing Then
        ReadSeedFeatureName = "None"
    Else
        ReadSeedFeatureName = SafeFeatureName(seedFeature)
    End If

    On Error GoTo 0
End Function

Private Function NormalizeEffectiveType( _
    ByVal rawType2 As String, _
    ByVal rawType1 As String) As String

    If UCase$(Trim$(rawType2)) = "ICE" Then
        NormalizeEffectiveType = Trim$(rawType1)
    Else
        NormalizeEffectiveType = Trim$(rawType2)
    End If
End Function

Private Function SafeGetTypeName2( _
    ByRef swFeature As SldWorks.Feature) As String

    On Error GoTo Failed
    SafeGetTypeName2 = Trim$(swFeature.GetTypeName2)
    Exit Function

Failed:
    SafeGetTypeName2 = _
        "Error:" & CStr(Err.Number) & ":" & Err.Description
End Function

Private Function SafeGetTypeName( _
    ByRef swFeature As SldWorks.Feature) As String

    On Error GoTo Failed
    SafeGetTypeName = Trim$(swFeature.GetTypeName)
    Exit Function

Failed:
    SafeGetTypeName = _
        "Error:" & CStr(Err.Number) & ":" & Err.Description
End Function

Private Function SafeGetDefinition( _
    ByRef swFeature As SldWorks.Feature) As Object

    On Error GoTo Failed
    Set SafeGetDefinition = swFeature.GetDefinition
    Exit Function

Failed:
    Set SafeGetDefinition = Nothing
End Function

Private Function ReadSuppressionState( _
    ByRef swFeature As SldWorks.Feature, _
    ByVal configurationName As String) As String

    On Error GoTo Failed

    Dim states As Variant
    states = swFeature.IsSuppressed2( _
        swSpecifyConfiguration, Array(configurationName))

    If IsArray(states) Then
        If VariantItemCount(states) < 1 Then
            ReadSuppressionState = "Unknown:EmptyArray"
        ElseIf CBool(states(LBound(states))) Then
            ReadSuppressionState = "Suppressed:Array"
        Else
            ReadSuppressionState = "Active:Array"
        End If
    ElseIf IsEmpty(states) Or IsNull(states) Then
        ReadSuppressionState = ReadCurrentSuppressionFallback(swFeature)
    ElseIf CBool(states) Then
        ReadSuppressionState = "Suppressed:Scalar"
    Else
        ReadSuppressionState = "Active:Scalar"
    End If
    Exit Function

Failed:
    ReadSuppressionState = _
        "Unknown:Error:" & CStr(Err.Number) & ":" & Err.Description
End Function

Private Function ReadCurrentSuppressionFallback( _
    ByRef swFeature As SldWorks.Feature) As String

    On Error GoTo Failed

    If swFeature.IsSuppressed Then
        ReadCurrentSuppressionFallback = _
            "Suppressed:CurrentFallbackAfterEmptySpecify"
    Else
        ReadCurrentSuppressionFallback = _
            "Active:CurrentFallbackAfterEmptySpecify"
    End If
    Exit Function

Failed:
    ReadCurrentSuppressionFallback = _
        "Unknown:EmptySpecifyAndCurrentError:" & CStr(Err.Number) & _
        ":" & Err.Description
End Function

Private Function ReadObjectProperty( _
    ByRef target As Object, _
    ByVal propertyName As String) As String

    On Error GoTo Failed

    Dim value As Variant
    value = CallByName(target, propertyName, VbGet)
    ReadObjectProperty = VariantValueToken(value)
    Exit Function

Failed:
    ReadObjectProperty = _
        "Unavailable:" & CStr(Err.Number)
End Function

Private Function ReadObjectMethod0( _
    ByRef target As Object, _
    ByVal methodName As String) As String

    On Error GoTo Failed

    Dim value As Variant
    value = CallByName(target, methodName, VbMethod)
    ReadObjectMethod0 = VariantValueToken(value)
    Exit Function

Failed:
    ReadObjectMethod0 = _
        "Unavailable:" & CStr(Err.Number)
End Function

Private Function VariantValueToken(ByVal value As Variant) As String
    If IsNull(value) Then
        VariantValueToken = "Null"
    ElseIf IsEmpty(value) Then
        VariantValueToken = "Empty"
    ElseIf IsArray(value) Then
        VariantValueToken = _
            "ArrayCount:" & CStr(VariantItemCount(value))
    ElseIf VarType(value) = vbBoolean Then
        VariantValueToken = CStr(CBool(value))
    ElseIf IsNumeric(value) Then
        VariantValueToken = FormatProbeNumber(CDbl(value))
    Else
        VariantValueToken = EvidenceToken(CStr(value))
    End If
End Function

Private Function VariantItemCount(ByVal items As Variant) As Long
    If Not IsArray(items) Then Exit Function

    On Error GoTo Failed
    VariantItemCount = UBound(items) - LBound(items) + 1
    Exit Function

Failed:
    VariantItemCount = 0
End Function

Private Function SafeFeatureName( _
    ByRef swFeature As SldWorks.Feature) As String

    If swFeature Is Nothing Then
        SafeFeatureName = "Nothing"
        Exit Function
    End If

    On Error GoTo Failed
    SafeFeatureName = swFeature.Name
    Exit Function

Failed:
    SafeFeatureName = _
        "Error:" & CStr(Err.Number)
End Function

Private Function SafeObjectTypeName(ByVal target As Object) As String
    If target Is Nothing Then
        SafeObjectTypeName = "Nothing"
        Exit Function
    End If

    On Error GoTo Failed
    SafeObjectTypeName = TypeName(target)
    Exit Function

Failed:
    SafeObjectTypeName = _
        "Error:" & CStr(Err.Number)
End Function

Private Function SafeObjectKey(ByVal target As Object) As String
    If target Is Nothing Then
        SafeObjectKey = "Nothing"
        Exit Function
    End If

    On Error GoTo Failed
    SafeObjectKey = _
        SafeObjectTypeName(target) & ":" & CStr(ObjPtr(target))
    Exit Function

Failed:
    SafeObjectKey = _
        SafeObjectTypeName(target) & ":Error:" & CStr(Err.Number)
End Function

Private Function IsAuthorizedProbeFixture( _
    ByVal partPath As String) As Boolean

    Dim normalized As String
    normalized = LCase$(Replace$(Trim$(partPath), "/", "\"))

    IsAuthorizedProbeFixture = _
        (normalized = LCase$( _
            "C:\Users\V.T\Documents\VBA 3D TO 2D\" & _
            "test_assets\models\P-0251-14A-001.SLDPRT")) Or _
        (normalized = LCase$( _
            "C:\Users\V.T\Documents\VBA 3D TO 2D\" & _
            "test_assets\models\P-0252-01-001.SLDPRT")) Or _
        (normalized = LCase$( _
            "C:\Users\V.T\Documents\VBA 3D TO 2D\" & _
            "test_assets\models\P-0252-01-013.SLDPRT"))
End Function

Private Function FormatProbeNumber(ByVal value As Double) As String
    FormatProbeNumber = Format$(value, "0.000000000")
End Function

Private Function EvidenceToken(ByVal value As String) As String
    EvidenceToken = Replace$(value, "|", "/")
    EvidenceToken = Replace$(EvidenceToken, vbCr, " ")
    EvidenceToken = Replace$(EvidenceToken, vbLf, " ")
End Function

Private Sub StartProbeLog()
    On Error GoTo Failed

    mR23ProbeLogPath = _
        R23_LOG_DIRECTORY & "\R23_FEATURE_" & _
        Format$(Now, "yyyymmdd_hhnnss") & ".log"

    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open mR23ProbeLogPath For Output As #fileNumber
    Close #fileNumber
    Exit Sub

Failed:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    mR23ProbeLogPath = vbNullString
    On Error GoTo 0
End Sub

Private Sub ProbeLog(ByVal message As String)
    Debug.Print message

    If Len(mR23ProbeLogPath) = 0 Then Exit Sub

    On Error GoTo Failed

    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open mR23ProbeLogPath For Append As #fileNumber
    Print #fileNumber, message
    Close #fileNumber
    Exit Sub

Failed:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    Debug.Print "R23_PROBE_LOG_ERROR|path=" & _
        EvidenceToken(mR23ProbeLogPath) & _
        "|error=" & CStr(Err.Number)
    On Error GoTo 0
End Sub

Private Sub CloseProbeLog()
    mR23ProbeLogPath = vbNullString
End Sub
