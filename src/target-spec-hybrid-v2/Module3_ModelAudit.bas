Option Explicit

Private Const swThisConfiguration As Long = 1
Private Const swSpecifyConfiguration As Long = 3

Public Function GetAllHoleLikeFeatures( _
    ByRef swPart As SldWorks.ModelDoc2) As Collection

    Dim results As New Collection
    Set GetAllHoleLikeFeatures = results

    If swPart Is Nothing Then Exit Function

    On Error GoTo Failed

    Dim swFeature As SldWorks.Feature
    Set swFeature = swPart.FirstFeature

    Do While Not swFeature Is Nothing
        If Not IsFeatureSuppressed(swFeature) Then
            If IsOwnedHoleFeature(swFeature) Then results.Add swFeature
        End If

        Set swFeature = swFeature.GetNextFeature
    Loop
    Exit Function

Failed:
    Debug.Print "Module3 GetAllHoleLikeFeatures error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CountHoles(ByVal holes As Collection) As Long
    If holes Is Nothing Then Exit Function
    CountHoles = holes.Count
End Function

Public Function DescribeFeatureType( _
    ByRef swFeature As SldWorks.Feature) As String

    If swFeature Is Nothing Then Exit Function
    DescribeFeatureType = ResolveFeatureType(swFeature)
End Function

Public Function IsOwnedHoleFeature( _
    ByRef swFeature As SldWorks.Feature) As Boolean

    If swFeature Is Nothing Then Exit Function

    Dim featureType As String
    featureType = UCase$(ResolveFeatureType(swFeature))

    Select Case featureType
        Case "HOLEWZD", "ADVHOLEWZD", "SKETCHHOLE"
            IsOwnedHoleFeature = True

        Case Else
            If InStr(featureType, "CUT") > 0 Then
                IsOwnedHoleFeature = HasInternalCylindricalFace(swFeature)
            End If
    End Select
End Function

Public Function FeatureContainsFace( _
    ByRef swFeature As SldWorks.Feature, _
    ByRef targetFace As SldWorks.Face2) As Boolean

    If swFeature Is Nothing Then Exit Function
    If targetFace Is Nothing Then Exit Function

    On Error GoTo Failed

    Dim faces As Variant
    faces = swFeature.GetFaces
    If IsEmpty(faces) Or Not IsArray(faces) Then Exit Function

    Dim i As Long
    For i = LBound(faces) To UBound(faces)
        Dim candidateFace As SldWorks.Face2
        Set candidateFace = faces(i)

        If Not candidateFace Is Nothing Then
            If targetFace.IsSame(candidateFace) Then
                FeatureContainsFace = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

Failed:
    FeatureContainsFace = False
End Function

Public Function IsInternalCylindricalFace( _
    ByRef swFace As SldWorks.Face2) As Boolean

    If swFace Is Nothing Then Exit Function

    On Error GoTo Failed

    Dim swSurface As SldWorks.Surface
    Set swSurface = swFace.GetSurface

    If swSurface Is Nothing Then Exit Function
    IsInternalCylindricalFace = _
        swSurface.IsCylinder And swFace.FaceInSurfaceSense
    Exit Function

Failed:
    IsInternalCylindricalFace = False
End Function

Public Function IsFeatureActiveInCurrentConfiguration( _
    ByRef swFeature As SldWorks.Feature) As Boolean

    If swFeature Is Nothing Then Exit Function
    IsFeatureActiveInCurrentConfiguration = Not IsFeatureSuppressed(swFeature)
End Function

Public Function IsFeatureActiveInConfiguration( _
    ByRef swFeature As SldWorks.Feature, _
    ByVal configurationName As String) As Boolean

    If swFeature Is Nothing Then Exit Function
    If Len(Trim$(configurationName)) = 0 Then Exit Function

    On Error GoTo UnknownState

    Dim configurationNames As Variant
    configurationNames = Array(configurationName)

    Dim states As Variant
    states = swFeature.IsSuppressed2( _
        swSpecifyConfiguration, configurationNames)

    If Not IsArray(states) Then GoTo UnknownState
    If UBound(states) < LBound(states) Then GoTo UnknownState

    IsFeatureActiveInConfiguration = _
        Not CBool(states(LBound(states)))
    Exit Function

UnknownState:
    IsFeatureActiveInConfiguration = False
End Function

Private Function ResolveFeatureType( _
    ByRef swFeature As SldWorks.Feature) As String

    On Error Resume Next
    ResolveFeatureType = Trim$(swFeature.GetTypeName2)

    If Len(ResolveFeatureType) = 0 Then
        ResolveFeatureType = Trim$(swFeature.GetTypeName)
    End If
    On Error GoTo 0
End Function

Private Function IsFeatureSuppressed( _
    ByRef swFeature As SldWorks.Feature) As Boolean

    On Error GoTo UnknownState

    Dim states As Variant
    states = swFeature.IsSuppressed2(swThisConfiguration, Empty)

    If IsArray(states) Then
        If UBound(states) >= LBound(states) Then
            IsFeatureSuppressed = CBool(states(LBound(states)))
        End If
    ElseIf Not IsEmpty(states) Then
        IsFeatureSuppressed = CBool(states)
    End If
    Exit Function

UnknownState:
    ' Fail closed for the preview count. Runtime geometry ownership performs
    ' its own proof and records rejection evidence.
    IsFeatureSuppressed = True
End Function

Private Function HasInternalCylindricalFace( _
    ByRef swFeature As SldWorks.Feature) As Boolean

    On Error GoTo Failed

    Dim faces As Variant
    faces = swFeature.GetFaces

    If IsEmpty(faces) Or Not IsArray(faces) Then Exit Function

    Dim i As Long
    For i = LBound(faces) To UBound(faces)
        Dim swFace As SldWorks.Face2
        Set swFace = faces(i)

        If Not swFace Is Nothing Then
            If IsInternalCylindricalFace(swFace) Then
                HasInternalCylindricalFace = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

Failed:
    HasInternalCylindricalFace = False
End Function
