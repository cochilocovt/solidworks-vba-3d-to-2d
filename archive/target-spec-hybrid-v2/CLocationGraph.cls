Option Explicit

' R23 indexed store for feature definitions, physical locations, view
' projections and imported annotations.
' CRunEvidence continues to record run results; it is no longer the
' production object store. Consolidation of coaxial cylinders happens here,
' in ResolveOrCreatePhysicalLocation, so there is exactly one place where the
' physical-identity invariants are applied.

Private mFeatures As Collection
Private mFeaturesByKey As Object
Private mFeatureKeyByFaceKey As Object

Private mLocations As Collection
Private mLocationsByPhysicalKey As Object
Private mLocationsByLineKey As Object
Private mLocationsByFamilyKey As Object

Private mProjections As Collection
Private mProjectionsByKey As Object
Private mProjectionsByView As Object

Private mAnnotations As Collection
Private mAnnotationsBySourceIdentity As Object
Private mAnnotationsByModelDimensionIdentity As Object

Private mConsolidationCount As Long

Private Sub Class_Initialize()
    Set mFeatures = New Collection
    Set mFeaturesByKey = CreateObject("Scripting.Dictionary")
    Set mFeatureKeyByFaceKey = CreateObject("Scripting.Dictionary")

    Set mLocations = New Collection
    Set mLocationsByPhysicalKey = CreateObject("Scripting.Dictionary")
    Set mLocationsByLineKey = CreateObject("Scripting.Dictionary")
    Set mLocationsByFamilyKey = CreateObject("Scripting.Dictionary")

    Set mProjections = New Collection
    Set mProjectionsByKey = CreateObject("Scripting.Dictionary")
    Set mProjectionsByView = CreateObject("Scripting.Dictionary")

    Set mAnnotations = New Collection
    Set mAnnotationsBySourceIdentity = CreateObject("Scripting.Dictionary")
    Set mAnnotationsByModelDimensionIdentity = _
        CreateObject("Scripting.Dictionary")

    mConsolidationCount = 0
End Sub

Public Sub AddFeatureDefinition( _
    ByRef featureDefinition As CFeatureDefinition)

    If featureDefinition Is Nothing Then Exit Sub

    mFeatures.Add featureDefinition

    Dim featureKey As String
    featureKey = featureDefinition.FeatureIdentityKey()

    If Not mFeaturesByKey.Exists(featureKey) Then
        mFeaturesByKey.Add featureKey, featureDefinition
    End If
End Sub

' Records which feature owns a face, so a later drawing entity can be traced
' back to feature ownership rather than accepted on appearance.
Public Sub RegisterFaceOwner( _
    ByVal faceKey As String, _
    ByRef featureDefinition As CFeatureDefinition)

    If featureDefinition Is Nothing Then Exit Sub
    If Len(Trim$(faceKey)) = 0 Then Exit Sub

    If Not mFeatureKeyByFaceKey.Exists(faceKey) Then
        mFeatureKeyByFaceKey.Add faceKey, _
            featureDefinition.FeatureIdentityKey()
    End If
End Sub

Public Function FeatureKeyForFace(ByVal faceKey As String) As String
    If mFeatureKeyByFaceKey.Exists(faceKey) Then
        FeatureKeyForFace = mFeatureKeyByFaceKey(faceKey)
    End If
End Function

' The single consolidation point.
' A candidate joins an existing location only when CanConsolidateWith is
' satisfied, which requires the same infinite line AND meeting axial
' intervals. Coaxial counterbore steps therefore merge, while opposite blind
' holes on one infinite line remain separate locations.
Public Function ResolveOrCreatePhysicalLocation( _
    ByRef candidate As CPhysicalHoleLocation) As CPhysicalHoleLocation

    If candidate Is Nothing Then Exit Function

    If Not candidate.IdentityProven Then
        Set ResolveOrCreatePhysicalLocation = Nothing
        Exit Function
    End If

    candidate.RecomputeKeys

    Dim bucket As Collection

    If mLocationsByLineKey.Exists(candidate.LineKey) Then
        Set bucket = mLocationsByLineKey(candidate.LineKey)

        Dim i As Long
        For i = 1 To bucket.Count
            Dim existing As CPhysicalHoleLocation
            Set existing = bucket(i)

            If existing.CanConsolidateWith(candidate) Then
                existing.MergeFrom candidate
                ReindexLocation existing
                mConsolidationCount = mConsolidationCount + 1
                Set ResolveOrCreatePhysicalLocation = existing
                Exit Function
            End If
        Next i
    Else
        Set bucket = New Collection
        mLocationsByLineKey.Add candidate.LineKey, bucket
    End If

    bucket.Add candidate
    mLocations.Add candidate
    IndexLocation candidate

    Set ResolveOrCreatePhysicalLocation = candidate
End Function

Private Sub IndexLocation(ByRef location As CPhysicalHoleLocation)
    If location Is Nothing Then Exit Sub

    If Len(location.PhysicalInstanceKey) > 0 Then
        If Not mLocationsByPhysicalKey.Exists( _
            location.PhysicalInstanceKey) Then

            mLocationsByPhysicalKey.Add _
                location.PhysicalInstanceKey, location
        End If
        location.IndexedPhysicalKey = location.PhysicalInstanceKey
    End If

    AddToFamilyIndex location
End Sub

' Merging changes the axial interval, which changes the physical key. The
' stale key is removed so a location is never reachable under two keys.
Private Sub ReindexLocation(ByRef location As CPhysicalHoleLocation)
    If location Is Nothing Then Exit Sub

    Dim previousKey As String
    previousKey = location.IndexedPhysicalKey

    If Len(previousKey) > 0 Then
        If StrComp(previousKey, location.PhysicalInstanceKey, _
            vbBinaryCompare) <> 0 Then

            If mLocationsByPhysicalKey.Exists(previousKey) Then
                mLocationsByPhysicalKey.Remove previousKey
            End If
        End If
    End If

    IndexLocation location
End Sub

Private Sub AddToFamilyIndex(ByRef location As CPhysicalHoleLocation)
    Dim familyKey As String
    familyKey = location.SemanticFamilyKey
    If Len(Trim$(familyKey)) = 0 Then Exit Sub

    Dim familyBucket As Collection

    If mLocationsByFamilyKey.Exists(familyKey) Then
        Set familyBucket = mLocationsByFamilyKey(familyKey)
    Else
        Set familyBucket = New Collection
        mLocationsByFamilyKey.Add familyKey, familyBucket
    End If

    Dim i As Long
    For i = 1 To familyBucket.Count
        If familyBucket(i) Is location Then Exit Sub
    Next i

    familyBucket.Add location
End Sub

Public Sub AddProjection(ByRef projection As CViewHoleProjection)
    If projection Is Nothing Then Exit Sub

    mProjections.Add projection

    Dim projectionKey As String
    projectionKey = projection.ProjectionKey()

    If Not mProjectionsByKey.Exists(projectionKey) Then
        mProjectionsByKey.Add projectionKey, projection
    End If

    Dim viewKey As String
    viewKey = Module11_GeometryIdentity.IdentityToken(projection.ViewName)

    Dim viewBucket As Collection
    If mProjectionsByView.Exists(viewKey) Then
        Set viewBucket = mProjectionsByView(viewKey)
    Else
        Set viewBucket = New Collection
        mProjectionsByView.Add viewKey, viewBucket
    End If

    viewBucket.Add projection
End Sub

Public Sub AddImportedAnnotation( _
    ByRef importedAnnotation As CImportedAnnotation)

    If importedAnnotation Is Nothing Then Exit Sub

    mAnnotations.Add importedAnnotation

    Dim identityKey As String
    identityKey = importedAnnotation.AnnotationIdentityKey()

    If Not mAnnotationsBySourceIdentity.Exists(identityKey) Then
        mAnnotationsBySourceIdentity.Add identityKey, importedAnnotation
    End If

    Dim modelIdentity As String
    modelIdentity = Module11_GeometryIdentity.IdentityToken( _
        importedAnnotation.SourceIdentity)

    If Len(modelIdentity) = 0 Then Exit Sub

    Dim modelBucket As Collection
    If mAnnotationsByModelDimensionIdentity.Exists(modelIdentity) Then
        Set modelBucket = mAnnotationsByModelDimensionIdentity(modelIdentity)
    Else
        Set modelBucket = New Collection
        mAnnotationsByModelDimensionIdentity.Add modelIdentity, modelBucket
    End If

    modelBucket.Add importedAnnotation
End Sub

' An annotation inventory is a snapshot of the sheet, not a cumulative
' event log. Final semantic QA must re-read it after callout/title/layout
' work without retaining stale COM objects from the pre-creation snapshot.
Public Sub ClearImportedAnnotations()
    Set mAnnotations = New Collection
    Set mAnnotationsBySourceIdentity = _
        CreateObject("Scripting.Dictionary")
    Set mAnnotationsByModelDimensionIdentity = _
        CreateObject("Scripting.Dictionary")
End Sub

Public Function FeatureCount() As Long
    FeatureCount = mFeatures.Count
End Function

Public Function LocationCount() As Long
    LocationCount = mLocations.Count
End Function

Public Function ProjectionCount() As Long
    ProjectionCount = mProjections.Count
End Function

Public Function AnnotationCount() As Long
    AnnotationCount = mAnnotations.Count
End Function

Public Function ConsolidationCount() As Long
    ConsolidationCount = mConsolidationCount
End Function

Public Function Features() As Collection
    Set Features = mFeatures
End Function

Public Function Locations() As Collection
    Set Locations = mLocations
End Function

Public Function Projections() As Collection
    Set Projections = mProjections
End Function

Public Function Annotations() As Collection
    Set Annotations = mAnnotations
End Function

Public Function LocationByPhysicalKey( _
    ByVal physicalKey As String) As CPhysicalHoleLocation

    If mLocationsByPhysicalKey.Exists(physicalKey) Then
        Set LocationByPhysicalKey = mLocationsByPhysicalKey(physicalKey)
    End If
End Function

Public Function LocationsForFamily( _
    ByVal familyKey As String) As Collection

    If mLocationsByFamilyKey.Exists(familyKey) Then
        Set LocationsForFamily = mLocationsByFamilyKey(familyKey)
    Else
        Set LocationsForFamily = New Collection
    End If
End Function

Public Function LocationsOnLine(ByVal lineKey As String) As Collection
    If mLocationsByLineKey.Exists(lineKey) Then
        Set LocationsOnLine = mLocationsByLineKey(lineKey)
    Else
        Set LocationsOnLine = New Collection
    End If
End Function

Public Function ProjectionsForView( _
    ByVal viewName As String) As Collection

    Dim viewKey As String
    viewKey = Module11_GeometryIdentity.IdentityToken(viewName)

    If mProjectionsByView.Exists(viewKey) Then
        Set ProjectionsForView = mProjectionsByView(viewKey)
    Else
        Set ProjectionsForView = New Collection
    End If
End Function

Public Function AnnotationsForModelDimension( _
    ByVal sourceIdentity As String) As Collection

    Dim modelIdentity As String
    modelIdentity = Module11_GeometryIdentity.IdentityToken(sourceIdentity)

    If mAnnotationsByModelDimensionIdentity.Exists(modelIdentity) Then
        Set AnnotationsForModelDimension = _
            mAnnotationsByModelDimensionIdentity(modelIdentity)
    Else
        Set AnnotationsForModelDimension = New Collection
    End If
End Function

Public Function FamilyKeys() As Collection
    Dim results As New Collection
    Dim keys As Variant
    Dim i As Long

    Set FamilyKeys = results
    If mLocationsByFamilyKey.Count = 0 Then Exit Function

    keys = mLocationsByFamilyKey.keys
    For i = LBound(keys) To UBound(keys)
        results.Add CStr(keys(i))
    Next i
End Function

Public Function GraphSummary() As String
    GraphSummary = _
        "features=" & CStr(FeatureCount()) & _
        "|physicalLocations=" & CStr(LocationCount()) & _
        "|consolidations=" & CStr(ConsolidationCount()) & _
        "|families=" & CStr(mLocationsByFamilyKey.Count) & _
        "|projections=" & CStr(ProjectionCount()) & _
        "|annotations=" & CStr(AnnotationCount())
End Function
