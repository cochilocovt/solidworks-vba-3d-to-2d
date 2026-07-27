Attribute VB_Name = "Module3_ModelAudit"

Option Explicit

Private Const IDX_NAME As Long = 0
Private Const IDX_TYPE As Long = 1
Private Const IDX_X As Long = 2
Private Const IDX_Y As Long = 3
Private Const IDX_DIA As Long = 4
Private Const IDX_LABEL As Long = 5
Private Const IDX_IS_HOLEWIZARD As Long = 6
Private Const IDX_IS_SIMPLECUT As Long = 7

Private Const DUP_POS_TOL As Double = 0.000001
Private Const DUP_DIA_TOL As Double = 0.000001

Public Function GetAllHoleLikeFeatures(ByRef swPart As SldWorks.ModelDoc2) As Collection
    On Error GoTo Failed

    Dim results As Collection
    Set results = New Collection

    If swPart Is Nothing Then
        Set GetAllHoleLikeFeatures = results
        Exit Function
    End If

    Dim swFeat As SldWorks.Feature
    Set swFeat = swPart.FirstFeature

    Do While Not swFeat Is Nothing
        CollectHoleInfoForFeature swFeat, results
        Set swFeat = swFeat.GetNextFeature
    Loop

    Set GetAllHoleLikeFeatures = results
    Exit Function

Failed:
    Dim emptyResults As Collection
    Set emptyResults = New Collection
    Set GetAllHoleLikeFeatures = emptyResults
End Function

Public Function CountHoles(ByVal holes As Collection) As Long
    On Error Resume Next
    If holes Is Nothing Then
        CountHoles = 0
    Else
        CountHoles = holes.Count
    End If
End Function

Public Function GetHoleLabel(ByVal holeInfo As Variant) As String
    On Error Resume Next
    GetHoleLabel = CStr(holeInfo(IDX_LABEL))
End Function

Public Function GetHoleX(ByVal holeInfo As Variant) As Double
    On Error Resume Next
    GetHoleX = CDbl(holeInfo(IDX_X))
End Function

Public Function GetHoleY(ByVal holeInfo As Variant) As Double
    On Error Resume Next
    GetHoleY = CDbl(holeInfo(IDX_Y))
End Function

Public Function GetHoleDiameter(ByVal holeInfo As Variant) As Double
    On Error Resume Next
    GetHoleDiameter = CDbl(holeInfo(IDX_DIA))
End Function

Public Function IsHoleWizardItem(ByVal holeInfo As Variant) As Boolean
    On Error Resume Next
    IsHoleWizardItem = CBool(holeInfo(IDX_IS_HOLEWIZARD))
End Function

Public Function DescribeFeatureType(ByRef swFeat As SldWorks.Feature) As String
    Dim t2 As String
    Dim t1 As String

    On Error Resume Next
    t2 = swFeat.GetTypeName2
    t1 = swFeat.GetTypeName
    On Error GoTo 0

    If UCase$(t2) = "ICE" And Len(t1) > 0 Then
        DescribeFeatureType = t1 & " (TypeName2=ICE)"
    ElseIf Len(t2) > 0 Then
        DescribeFeatureType = t2
    Else
        DescribeFeatureType = t1
    End If
End Function

Private Sub CollectHoleInfoForFeature(ByRef swFeat As SldWorks.Feature, ByRef results As Collection)
    Dim resolvedType As String
    resolvedType = ResolveFeatureType(swFeat)

    If IsHoleWizardType(resolvedType) Then
        CollectHoleWizardInfo swFeat, results
    ElseIf IsSimpleCutCandidate(resolvedType) Then
        CollectSimpleCutInfo swFeat, results
    End If
End Sub

Private Function ResolveFeatureType(ByRef swFeat As SldWorks.Feature) As String
    Dim t2 As String
    Dim t1 As String

    On Error Resume Next
    t2 = swFeat.GetTypeName2
    t1 = swFeat.GetTypeName
    On Error GoTo 0

    If UCase$(t2) = "ICE" Then
        ResolveFeatureType = t1
    ElseIf Len(t2) > 0 Then
        ResolveFeatureType = t2
    Else
        ResolveFeatureType = t1
    End If
End Function

Private Function IsHoleWizardType(ByVal featureType As String) As Boolean
    IsHoleWizardType = (UCase$(featureType) = "HOLEWZD")
End Function

Private Function IsSimpleCutCandidate(ByVal featureType As String) As Boolean
    featureType = UCase$(featureType)
    IsSimpleCutCandidate = (InStr(1, featureType, "CUT", vbTextCompare) > 0)
End Function

Private Sub CollectHoleWizardInfo(ByRef swFeat As SldWorks.Feature, ByRef results As Collection)
    On Error GoTo SafeExit

    Dim locSketch As SldWorks.Sketch
    Set locSketch = FindLocationSketch(swFeat)
    If locSketch Is Nothing Then
        AddIfUnique results, Array(swFeat.Name, ResolveFeatureType(swFeat), 0#, 0#, 0#, swFeat.Name & " (Hole Wizard)", True, False)
        Exit Sub
    End If

    Dim pts As Variant
    pts = locSketch.GetSketchPoints2

    If IsEmpty(pts) Then
        AddIfUnique results, Array(swFeat.Name, ResolveFeatureType(swFeat), 0#, 0#, 0#, swFeat.Name & " (Hole Wizard)", True, False)
        Exit Sub
    End If

    Dim i As Long
    For i = LBound(pts) To UBound(pts)
        Dim swPt As SldWorks.SketchPoint
        Set swPt = pts(i)
        If Not swPt Is Nothing Then
            AddIfUnique results, Array( _
                swFeat.Name, _
                ResolveFeatureType(swFeat), _
                swPt.x, _
                swPt.y, _
                0#, _
                swFeat.Name & " @ X:" & Format(swPt.x * 1000#, "0.0") & " Y:" & Format(swPt.y * 1000#, "0.0"), _
                True, _
                False)
        End If
    Next i

SafeExit:
End Sub

Private Sub CollectSimpleCutInfo(ByRef swFeat As SldWorks.Feature, ByRef results As Collection)
    On Error GoTo SafeExit

    Dim swSketch As SldWorks.Sketch
    Set swSketch = FindDrivingSketch(swFeat)
    If swSketch Is Nothing Then Exit Sub

    Dim segs As Variant
    segs = swSketch.GetSketchSegments
    If IsEmpty(segs) Then Exit Sub

    Dim i As Long
    For i = LBound(segs) To UBound(segs)
        Dim swSeg As SldWorks.SketchSegment
        Set swSeg = segs(i)
        If Not swSeg Is Nothing Then
            If swSeg.GetType = 1 Then
                Dim swArc As SldWorks.SketchArc
                Set swArc = swSeg
                If swArc.IsCircle = 1 Then
                    Dim swCtr As SldWorks.SketchPoint
                    Set swCtr = swArc.GetCenterPoint2
                    If Not swCtr Is Nothing Then
                        AddIfUnique results, Array( _
                            swFeat.Name, _
                            ResolveFeatureType(swFeat), _
                            swCtr.x, _
                            swCtr.y, _
                            swArc.GetRadius * 2#, _
                            swFeat.Name & " O" & Format(swArc.GetRadius * 2000#, "0.0"), _
                            False, _
                            True)
                    End If
                End If
            End If
        End If
    Next i

SafeExit:
End Sub

Private Function FindLocationSketch(ByRef swFeat As SldWorks.Feature) As SldWorks.Sketch
    On Error GoTo SafeExit

    Dim swSubFeat As SldWorks.Feature
    Set swSubFeat = swFeat.GetFirstSubFeature

    Do While Not swSubFeat Is Nothing
        If UCase$(swSubFeat.GetTypeName2) = "PROFILEFEATURE" Then
            Dim sk As Object
            Set sk = swSubFeat.GetSpecificFeature2
            If Not sk Is Nothing Then
                If TypeOf sk Is SldWorks.Sketch Then
                    Dim pts As Variant
                    pts = sk.GetSketchPoints2
                    If Not IsEmpty(pts) Then
                        Set FindLocationSketch = sk
                        Exit Function
                    End If
                End If
            End If
        End If
        Set swSubFeat = swSubFeat.GetNextSubFeature
    Loop

SafeExit:
End Function

Private Function FindDrivingSketch(ByRef swFeat As SldWorks.Feature) As SldWorks.Sketch
    On Error GoTo SafeExit

    Dim swSubFeat As SldWorks.Feature
    Set swSubFeat = swFeat.GetFirstSubFeature

    Do While Not swSubFeat Is Nothing
        If UCase$(swSubFeat.GetTypeName2) = "PROFILEFEATURE" Then
            Dim obj As Object
            Set obj = swSubFeat.GetSpecificFeature2
            If Not obj Is Nothing Then
                If TypeOf obj Is SldWorks.Sketch Then
                    Set FindDrivingSketch = obj
                    Exit Function
                End If
            End If
        End If
        Set swSubFeat = swSubFeat.GetNextSubFeature
    Loop

SafeExit:
End Function

Private Sub AddIfUnique(ByRef results As Collection, ByVal candidate As Variant)
    If Not isDuplicate(results, candidate) Then results.Add candidate
End Sub

Private Function isDuplicate(ByVal results As Collection, ByVal candidate As Variant) As Boolean
    On Error GoTo SafeExit

    Dim i As Long
    For i = 1 To results.Count
        Dim cur As Variant
        cur = results.item(i)

        If CStr(cur(IDX_NAME)) = CStr(candidate(IDX_NAME)) Then
            If Abs(CDbl(cur(IDX_X)) - CDbl(candidate(IDX_X))) < DUP_POS_TOL _
            And Abs(CDbl(cur(IDX_Y)) - CDbl(candidate(IDX_Y))) < DUP_POS_TOL _
            And Abs(CDbl(cur(IDX_DIA)) - CDbl(candidate(IDX_DIA))) < DUP_DIA_TOL Then
                isDuplicate = True
                Exit Function
            End If
        End If
    Next i

SafeExit:
End Function

